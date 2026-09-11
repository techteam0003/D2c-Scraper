#!/usr/bin/env python3
"""
Multi-Platform Dealer Inventory Scraper - local web app.

Run:
    pip install -r requirements.txt
    python app.py
Then open:
    http://localhost:5000

Paste one or more dealer LISTING-page URLs (one per line), watch live
progress, then browse combined results and download CSV or Excel.

WHY THIS VERSION EXISTS
------------------------
The original tool was built around one specific SM360 URL shape:
    /used/{slug}-id{stock_id}.html
That pattern is used by some SM360 dealer sites, but NOT all of them.
Inspecting a second dealership (Hamel Chevrolet Buick GMC, also SM360 --
confirmed by its <img src="https://img.sm360.ca/..."> assets and by the
schema.org/Car JSON-LD blocks on its vehicle pages) showed a DIFFERENT
detail-page URL shape:
    /en/used-inventory/{make}/{model}/{year}-{make}-{model}-id{stock_id}
(no ".html", and the path prefix is "used-inventory" not "used").
Its listing pages also paginate with ?page=N (seen alongside
namedSorting=default&limit=12 in the query string), rather than the
generic page/pageNumber/p guesswork used before.

A third dealer platform (seen in a sample CSV of garagetardif.com, a
WordPress-based site, NOT SM360) uses yet another shape:
    /en/pre-owned/{year}/{make}/{model}/{numeric_id}

Rather than hard-coding one shape, this version tries several known
detail-URL patterns against each listing page and uses whichever one
actually finds vehicle links -- logging which "platform" it detected.
New patterns can be added to DETAIL_URL_PATTERNS without touching the
rest of the pipeline.
"""

import csv
import io
import json
import queue
import re
import threading
import time
import uuid
from dataclasses import dataclass, asdict, fields
from urllib.parse import urljoin, urlparse, parse_qs, urlencode

import requests
from bs4 import BeautifulSoup
from flask import Flask, Response, jsonify, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

app = Flask(__name__)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-CA,en;q=0.9,fr-CA;q=0.8,fr;q=0.7",
}
REQUEST_DELAY = 1.2
REQUEST_TIMEOUT = 15
DEFAULT_MAX_PAGES = 40

# --------------------------------------------------------------------------
# Platform detection: each entry is (platform_name, compiled_regex).
# The regex must have its first capturing group be the numeric stock/listing
# id. Tried in order against a listing page's HTML; the first pattern that
# matches at least one href wins for that source URL (and every subsequent
# paginated page from that same source URL reuses the same pattern).
# --------------------------------------------------------------------------
DETAIL_URL_PATTERNS = [
    ("sm360-html",
     re.compile(r"/used/[^\"'>\s]+-id(\d+)\.html", re.IGNORECASE)),

    ("sm360-inventory-path",
     # e.g. /en/used-inventory/ford/maverick/2023-ford-maverick-id35313906
     # also matches /en/new-inventory/... on the same platform
     re.compile(r"/(?:used|new|all)-inventory/[a-z0-9\-]+/[a-z0-9\-]+/"
                r"[a-z0-9\-]*-id(\d+)(?=[/?#\"'>\s]|$)", re.IGNORECASE)),

    ("syncauto-numeric-path",
     # e.g. /en/pre-owned/2004/nissan/350z/2745
     re.compile(r"/(?:pre-owned|used|inventory)/(?:19|20)\d{2}/"
                r"[a-z0-9\-]+/[a-z0-9\-]+/(\d{3,9})(?=[/?#\"'>\s]|$)",
                re.IGNORECASE)),

    ("generic-id-suffix",
     # catch-all: any href ending in "...-id123456" regardless of path,
     # optionally followed by .html
     re.compile(r"-id(\d{4,10})(?:\.html)?(?=[/?#\"'>\s]|$)", re.IGNORECASE)),
]

# Params various listing pages use for pagination, tried in order.
PAGE_PARAM_CANDIDATES = ["page", "pageNumber", "p"]

CSV_FIELDNAMES = [
    "source_url", "domain", "url", "stock_id", "title", "year", "make",
    "model", "trim", "price", "mileage", "vin", "stock_number",
    "exterior_color", "interior_color", "transmission", "engine",
    "fuel_type", "drivetrain", "body_type", "image_url", "extraction_notes",
]

# job_id -> {
#   "queue": Queue,
#   "status": "running|done|error",
#   "url_jobs": [ {"url":..., "domain":..., "status":..., "platform":...,
#                  "vehicles":[...]} , ... ],
# }
JOBS = {}


@dataclass
class Vehicle:
    source_url: str = ""
    domain: str = ""
    url: str = ""
    stock_id: str = ""
    title: str = ""
    year: str = ""
    make: str = ""
    model: str = ""
    trim: str = ""
    price: str = ""
    mileage: str = ""
    vin: str = ""
    stock_number: str = ""
    exterior_color: str = ""
    interior_color: str = ""
    transmission: str = ""
    engine: str = ""
    fuel_type: str = ""
    drivetrain: str = ""
    body_type: str = ""
    image_url: str = ""
    extraction_notes: str = ""


def normalize_url(raw: str) -> str:
    raw = raw.strip()
    if not raw:
        return ""
    if not raw.startswith("http"):
        raw = "https://" + raw
    return raw


def get(url: str):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        if resp.status_code == 200:
            return resp
        return None
    except requests.RequestException:
        return None


def detect_platform_and_links(html: str, base_url: str):
    """Try each known detail-URL pattern in order. Returns
    (platform_name, regex, set_of_absolute_urls) for the first pattern
    that finds at least one match, or (None, None, set()) if none do."""
    soup = BeautifulSoup(html, "html.parser")
    hrefs = [a["href"] for a in soup.find_all("a", href=True)]
    for name, pattern in DETAIL_URL_PATTERNS:
        links = {urljoin(base_url, h) for h in hrefs if pattern.search(h)}
        if links:
            return name, pattern, links
    return None, None, set()


def find_detail_links_with_pattern(html: str, base_url: str, pattern) -> set:
    soup = BeautifulSoup(html, "html.parser")
    links = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if pattern.search(href):
            links.add(urljoin(base_url, href))
    return links


def stock_id_from_url(url: str, pattern) -> str:
    if pattern is None:
        return ""
    m = pattern.search(url)
    return m.group(1) if m else ""


# Words that commonly start the NEXT labelled field on a dealer page. Used
# as a stop boundary so we don't swallow the next label's value into the
# current one.
FIELD_STOP_PREFIXES = (
    "interior", "exterior", "ext", "int", "colo",  # colo(u)r / colo(u)rs
    "vin", "stock", "transmiss", "engine", "mileage", "odomet", "door",
    "passenger", "fuel", "drivetrain", "drive", "body", "trim", "price",
    "km", "cylind", "seat", "gear", "config",
)


def regex_field(text: str, *labels: str, max_words: int = 6) -> str:
    """Grab the value following a label by walking word-by-word until we
    hit a word that looks like the start of the NEXT label."""
    for label in labels:
        m = re.search(re.escape(label) + r"\s*[:#]?\s*", text, re.IGNORECASE)
        if not m:
            continue
        words = text[m.end():].split()
        collected = []
        for w in words[:max_words]:
            bare = w.strip(".,:#").lower()
            if bare.startswith(FIELD_STOP_PREFIXES):
                break
            collected.append(w)
        val = " ".join(collected).strip().rstrip(".,:#-")
        if val:
            return val
    return ""


def regex_token_field(text: str, *labels: str) -> str:
    """Like regex_field, but for values that should never contain a space
    (stock numbers, codes) -- stops at the first whitespace no matter what."""
    for label in labels:
        pattern = re.compile(re.escape(label) + r"\s*[:#]?\s*([A-Za-z0-9\-]{2,20})", re.IGNORECASE)
        m = pattern.search(text)
        if m:
            return m.group(1).strip()
    return ""


def extract_jsonld_vehicle(soup: BeautifulSoup) -> dict:
    for script in soup.find_all("script", type="application/ld+json"):
        raw = script.string or script.get_text() or "{}"
        try:
            data = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            continue
        candidates = data if isinstance(data, list) else [data]
        for item in candidates:
            if not isinstance(item, dict):
                continue
            types = item.get("@type", "")
            types = types if isinstance(types, list) else [types]
            if any(t in ("Vehicle", "Car", "Product") for t in types):
                return item
            if "@graph" in item:
                for sub in item["@graph"]:
                    sub_types = sub.get("@type", "")
                    sub_types = sub_types if isinstance(sub_types, list) else [sub_types]
                    if any(t in ("Vehicle", "Car", "Product") for t in sub_types):
                        return sub
    return {}


# Markers that typically introduce trailing site-name / marketing text in a
# <title> tag.
TITLE_CUT_MARKERS = [
    "|", " – ", " — ", " - Stock", " Stock #", " For Sale",
    " for sale", " à vendre", " A vendre", " $",
]


def clean_title(title: str) -> str:
    if not title:
        return title
    cut_at = len(title)
    for marker in TITLE_CUT_MARKERS:
        idx = title.find(marker)
        if idx > 0:
            cut_at = min(cut_at, idx)
    return title[:cut_at].strip()


def cap_trim(remainder: str, max_chars: int = 25) -> str:
    """Keep only the first few words of whatever's left after year/make/model
    so option/feature text doesn't get pulled into the trim field."""
    words = remainder.strip(" -|").split()
    kept, length = [], 0
    for w in words:
        if kept and length + len(w) + 1 > max_chars:
            break
        kept.append(w)
        length += len(w) + 1
    return " ".join(kept)


def _na(value):
    """SM360's JSON-LD frequently uses the literal string "n/a" instead of
    omitting a field. Treat that the same as missing."""
    if value is None:
        return ""
    s = str(value).strip()
    return "" if s.lower() in ("n/a", "na", "none", "null") else s


def parse_detail_page(html: str, url: str, domain: str, source_url: str,
                       detail_pattern) -> Vehicle:
    soup = BeautifulSoup(html, "html.parser")
    v = Vehicle(source_url=source_url, domain=domain, url=url)
    v.stock_id = stock_id_from_url(url, detail_pattern)
    text = soup.get_text(" ", strip=True)

    ld = extract_jsonld_vehicle(soup)
    if ld:
        v.title = _na(ld.get("name", ""))
        offers = ld.get("offers")
        if isinstance(offers, dict):
            v.price = _na(offers.get("price", ""))
        mfo = ld.get("mileageFromOdometer")
        v.mileage = _na(mfo.get("value")) if isinstance(mfo, dict) else _na(mfo)
        brand = ld.get("brand")
        v.make = _na(brand.get("name") if isinstance(brand, dict) else brand)
        model = ld.get("model")
        v.model = _na(model if isinstance(model, str) else (model or {}).get("name", ""))
        v.body_type = _na(ld.get("bodyType", ""))
        v.fuel_type = _na(ld.get("fuelType", ""))
        v.drivetrain = _na(ld.get("driveWheelConfiguration", "")).replace(
            "https://schema.org/", "").replace("Configuration", "")
        v.transmission = _na(ld.get("vehicleTransmission", ""))
        v.exterior_color = _na(ld.get("color", ""))
        v.interior_color = _na(ld.get("vehicleInteriorColor", ""))
        v.trim = _na(ld.get("vehicleConfiguration", ""))
        v.vin = _na(ld.get("vehicleIdentificationNumber", ""))
        v.year = _na(ld.get("modelDate", "") or ld.get("productionDate", ""))
        img = ld.get("image")
        if isinstance(img, str):
            v.image_url = img
        elif isinstance(img, list) and img:
            v.image_url = img[0] if isinstance(img[0], str) else ""
        if not v.fuel_type:
            engines = ld.get("vehicleEngine")
            if isinstance(engines, list) and engines:
                e0 = engines[0] if isinstance(engines[0], dict) else {}
                v.fuel_type = _na(e0.get("fuelType", ""))
                if not v.engine:
                    v.engine = _na(e0.get("engineType", ""))
        v.extraction_notes = "json-ld"

    if not v.title:
        title_tag = soup.find("title")
        if title_tag:
            v.title = title_tag.get_text(strip=True)
    og_image = soup.find("meta", property="og:image")
    if og_image and not v.image_url:
        v.image_url = og_image.get("content", "")

    if not v.price:
        pm = re.search(r"\$\s?[\d,]{4,7}", text)
        if pm:
            v.price = pm.group(0)
    if not v.mileage:
        mm = re.search(r"([\d,]{3,7})\s?km", text, re.IGNORECASE)
        if mm:
            v.mileage = mm.group(1)
    if not v.vin:
        vinm = re.search(r"\bVIN[:\s]*([A-HJ-NPR-Z0-9]{11,17})\b", text, re.IGNORECASE)
        if vinm:
            v.vin = vinm.group(1)
    if not v.stock_number:
        v.stock_number = regex_token_field(
            text, "Stock is #", "Stock #", "Stock#", "Stock number", "Stock No")
    if not v.transmission:
        v.transmission = regex_field(text, "Transmission")
    if not v.engine:
        v.engine = regex_field(text, "Engine")
    if not v.drivetrain:
        v.drivetrain = regex_field(text, "Drivetrain")
    if not v.body_type:
        v.body_type = regex_field(text, "Bodystyle", "Body Style", "Body type")
    if not v.fuel_type:
        v.fuel_type = regex_field(text, "Fuel", max_words=2)
    if not v.exterior_color:
        v.exterior_color = regex_field(
            text, "Ext. Colors", "Ext. Color", "Exterior Colour", "Exterior Color", "Ext Color")
    if not v.interior_color:
        v.interior_color = regex_field(
            text, "Int. Colors", "Int. color", "Interior Colour", "Interior Color", "Int Color")

    if v.title and not (v.year and v.make and v.model):
        ymm = re.match(r"(\d{4})\s+([A-Za-z\-]+)\s+([A-Za-z0-9\-]+)(.*)", clean_title(v.title))
        if ymm:
            v.year = v.year or ymm.group(1)
            v.make = v.make or ymm.group(2)
            v.model = v.model or ymm.group(3)
            v.trim = v.trim or cap_trim(ymm.group(4))

    if not v.extraction_notes:
        v.extraction_notes = "regex-fallback"
    elif "json-ld" in v.extraction_notes and (not v.stock_number or not v.transmission):
        v.extraction_notes += "+regex-fallback"
    return v


def build_page_url(source_url: str, param: str, page_num: int) -> str:
    """Re-issue the original URL with its existing query params intact,
    overriding/adding only the pagination param. This preserves filters
    like namedSorting=default&limit=12 that some SM360 sites require."""
    parsed = urlparse(source_url)
    query = parse_qs(parsed.query)
    # drop any other page-like params so we don't send page=2&pageNumber=2
    for other in PAGE_PARAM_CANDIDATES:
        query.pop(other, None)
    query[param] = [str(page_num)]
    new_query = urlencode(query, doseq=True)
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}?{new_query}"


def scrape_one_url(source_url: str, max_pages: int, log_fn, progress_fn) -> tuple:
    """Scrape a single listing-page URL (plus its pagination), returning
    (platform_name, list_of_Vehicle)."""
    domain = urlparse(source_url).netloc
    vehicles = []

    log_fn(f"Fetching listing page: {source_url}")
    resp = get(source_url)
    time.sleep(REQUEST_DELAY)
    if resp is None:
        log_fn(f"Could not fetch {source_url} (site unreachable or non-200 response).", is_err=True)
        return None, vehicles

    platform, pattern, links = detect_platform_and_links(resp.text, source_url)
    if not links:
        log_fn("No vehicle links found on this page with any known platform "
               "pattern (SM360 .html, SM360 inventory-path, syncauto numeric-path, "
               "generic -idNNNN). This site may use a platform this tool doesn't "
               "recognize yet -- inspect a vehicle detail-page URL manually and "
               "add a pattern to DETAIL_URL_PATTERNS in app.py.", is_err=True)
        return None, vehicles

    log_fn(f"Detected platform: {platform} — found {len(links)} vehicle link(s) on the first page.")

    # paginate, preserving other query params, trying each page-param name
    for param in PAGE_PARAM_CANDIDATES:
        stagnant = 0
        found_any_for_param = False
        for page_num in range(2, max_pages + 1):
            page_url = build_page_url(source_url, param, page_num)
            resp = get(page_url)
            time.sleep(REQUEST_DELAY)
            if resp is None:
                break
            new_links = find_detail_links_with_pattern(resp.text, page_url, pattern)
            before = len(links)
            links |= new_links
            if len(links) == before:
                stagnant += 1
                if stagnant >= 2:
                    break
            else:
                found_any_for_param = True
                log_fn(f"Page {page_num} ({param}={page_num}): {len(links)} total vehicle links so far")
                stagnant = 0
        if found_any_for_param:
            break

    log_fn(f"Total unique vehicle pages to scrape: {len(links)}")

    links_sorted = sorted(links)
    total = len(links_sorted)
    for i, url in enumerate(links_sorted, 1):
        resp = get(url)
        time.sleep(REQUEST_DELAY)
        if resp is not None:
            v = parse_detail_page(resp.text, url, domain, source_url, pattern)
            vehicles.append(v)
        progress_fn(i, total)

    return platform, vehicles


def run_batch_job(job_id: str, urls: list, max_pages: int):
    q = JOBS[job_id]["queue"]
    url_jobs = JOBS[job_id]["url_jobs"]
    n_urls = len(urls)

    for idx, source_url in enumerate(urls):
        url_jobs[idx]["status"] = "running"
        q.put({"type": "url_start", "url_index": idx, "url": source_url,
               "overall_index": idx + 1, "overall_total": n_urls})

        def log_fn(msg, is_err=False, _idx=idx):
            q.put({"type": "log", "url_index": _idx, "message": msg, "is_err": is_err})

        def progress_fn(current, total, _idx=idx):
            q.put({"type": "progress", "url_index": _idx, "current": current, "total": total})

        try:
            platform, vehicles = scrape_one_url(source_url, max_pages, log_fn, progress_fn)
            url_jobs[idx]["vehicles"] = vehicles
            url_jobs[idx]["platform"] = platform or "unknown"
            url_jobs[idx]["status"] = "done" if vehicles else "error"
            q.put({"type": "url_done", "url_index": idx, "count": len(vehicles),
                   "platform": platform or "unknown"})
        except Exception as e:
            url_jobs[idx]["status"] = "error"
            q.put({"type": "log", "url_index": idx, "message": f"ERROR: {e}", "is_err": True})
            q.put({"type": "url_done", "url_index": idx, "count": 0, "platform": "error"})

    total_count = sum(len(uj["vehicles"]) for uj in url_jobs)
    JOBS[job_id]["status"] = "done"
    q.put({"type": "done", "total_count": total_count})


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/scrape", methods=["POST"])
def api_scrape():
    data = request.get_json(force=True)
    raw_urls = data.get("urls", [])
    max_pages = int(data.get("max_pages", DEFAULT_MAX_PAGES))

    urls = [normalize_url(u) for u in raw_urls if normalize_url(u)]
    seen = set()
    clean_urls = []
    for u in urls:
        if u not in seen:
            seen.add(u)
            clean_urls.append(u)

    if not clean_urls:
        return jsonify({"error": "No valid URLs provided."}), 400

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "queue": queue.Queue(),
        "status": "running",
        "url_jobs": [
            {"url": u, "domain": urlparse(u).netloc, "status": "pending",
             "platform": None, "vehicles": []}
            for u in clean_urls
        ],
    }
    t = threading.Thread(target=run_batch_job, args=(job_id, clean_urls, max_pages), daemon=True)
    t.start()
    return jsonify({"job_id": job_id, "urls": clean_urls})


@app.route("/api/stream/<job_id>")
def api_stream(job_id):
    if job_id not in JOBS:
        return "Unknown job", 404

    def generate():
        q = JOBS[job_id]["queue"]
        while True:
            try:
                item = q.get(timeout=10)
                yield f"data: {json.dumps(item)}\n\n"
                if item.get("type") == "done":
                    break
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'ping'})}\n\n"

    return Response(generate(), mimetype="text/event-stream")


@app.route("/api/results/<job_id>")
def api_results(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job"}), 404

    all_vehicles = []
    summary = []
    for uj in job["url_jobs"]:
        all_vehicles.extend(asdict(v) for v in uj["vehicles"])
        summary.append({
            "url": uj["url"],
            "domain": uj["domain"],
            "status": uj["status"],
            "platform": uj.get("platform") or "-",
            "count": len(uj["vehicles"]),
        })

    return jsonify({
        "status": job["status"],
        "summary": summary,
        "vehicles": all_vehicles,
    })


def _safe_sheet_name(name: str, used: set) -> str:
    clean = re.sub(r"[\[\]\:\*\?/\\]", "-", name)[:31]
    base = clean or "Sheet"
    candidate = base
    i = 2
    while candidate in used:
        suffix = f" ({i})"
        candidate = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(candidate)
    return candidate


def _all_vehicles_for_job(job):
    all_vehicles = []
    for uj in job["url_jobs"]:
        all_vehicles.extend(uj["vehicles"])
    return all_vehicles


@app.route("/api/download/<job_id>")
def api_download(job_id):
    """Excel workbook: combined sheet + summary + one sheet per source URL."""
    job = JOBS.get(job_id)
    if not job:
        return "Unknown job", 404

    url_jobs = job["url_jobs"]
    all_vehicles = _all_vehicles_for_job(job)
    if not all_vehicles:
        return "No data available for this job", 404

    fieldnames = CSV_FIELDNAMES
    header_font = Font(bold=True)
    wb = Workbook()

    ws_all = wb.active
    ws_all.title = "All Vehicles"
    ws_all.append(fieldnames)
    for cell in ws_all[1]:
        cell.font = header_font
    for v in all_vehicles:
        d = asdict(v)
        ws_all.append([d[f] for f in fieldnames])
    for i, col in enumerate(fieldnames, 1):
        ws_all.column_dimensions[get_column_letter(i)].width = max(12, min(40, len(col) + 4))
    ws_all.freeze_panes = "A2"

    ws_sum = wb.create_sheet("Summary")
    ws_sum.append(["Source URL", "Domain", "Platform", "Status", "Vehicles Found"])
    for cell in ws_sum[1]:
        cell.font = header_font
    for uj in url_jobs:
        ws_sum.append([uj["url"], uj["domain"], uj.get("platform") or "-",
                        uj["status"], len(uj["vehicles"])])
    ws_sum.column_dimensions["A"].width = 60
    ws_sum.column_dimensions["B"].width = 30
    ws_sum.column_dimensions["C"].width = 20
    ws_sum.column_dimensions["D"].width = 12
    ws_sum.column_dimensions["E"].width = 16

    used_names = {"All Vehicles", "Summary"}
    for uj in url_jobs:
        if not uj["vehicles"]:
            continue
        label = uj["domain"] or uj["url"]
        sheet_name = _safe_sheet_name(label, used_names)
        ws = wb.create_sheet(sheet_name)
        ws.append(fieldnames)
        for cell in ws[1]:
            cell.font = header_font
        for v in uj["vehicles"]:
            d = asdict(v)
            ws.append([d[f] for f in fieldnames])
        for i, col in enumerate(fieldnames, 1):
            ws.column_dimensions[get_column_letter(i)].width = max(12, min(40, len(col) + 4))
        ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="dealer_inventory_scrape.xlsx",
    )


@app.route("/api/download_csv/<job_id>")
def api_download_csv(job_id):
    """Single CSV, columns matching the garage_tardif.csv reference schema."""
    job = JOBS.get(job_id)
    if not job:
        return "Unknown job", 404

    all_vehicles = _all_vehicles_for_job(job)
    if not all_vehicles:
        return "No data available for this job", 404

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=CSV_FIELDNAMES)
    writer.writeheader()
    for v in all_vehicles:
        writer.writerow(asdict(v))

    mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(
        mem,
        mimetype="text/csv",
        as_attachment=True,
        download_name="dealer_inventory_scrape.csv",
    )


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)
