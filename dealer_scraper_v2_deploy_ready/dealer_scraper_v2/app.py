#!/usr/bin/env python3
"""
Multi-Platform Dealer Inventory Scraper - local web app (v3).

Run:
    pip install -r requirements.txt
    python app.py            (or:  PORT=5002 python app.py)
Then open http://localhost:5000

WHAT WAS WRONG IN v2 (found on valleyfieldhonda.com, an SM360 / 360.Agency site)
--------------------------------------------------------------------------------
1. SM360 listing pages build their vehicle cards with JavaScript. A plain
   `requests.get()` sees ZERO cars, and ?page=2..N return the identical shell.
2. With no real cars in the HTML, the "generic-id-suffix" fallback matched the
   footer's "Honda Vehicles" links (/en/new-catalog/honda/...-id33538). Those
   are model-brochure pages, so the output was 15 NEW 2025-2026 Hondas with
   VIN 00000000000000000 and mileage 1 -- and the same 15 on every page URL.
3. The VIN regex didn't allow "VIN #1C4...", so the placeholder JSON-LD VIN
   was never overridden.
4. Detail-page regexes ran over the whole page, including the "Similar
   Vehicles" carousel, which holds other cars' prices/VINs/stock numbers.

FIXES
-----
* SM360 sites are detected (img.sm360.ca / 360.agency) and the full inventory
  is read from the site's own HTML sitemap (/en/sitemap or /fr/plan-du-site),
  which lists every in-stock vehicle. It is then filtered to what the listing
  URL asked for: used vs new, make/model sub-paths, certified-only,
  hybrid/EV-only.
* Catalog / news / offers / form links are never treated as vehicles.
* Detail pages are parsed from SM360's server-rendered "Specifications" block
  (Stock #, VIN #, Fuel, colours, Drivetrain, Trim, Transmission, Mileage,
  Bodystyle, Doors, Passengers, Cylinders, Engine), with selling price AND
  original (pre-reduction) price. JSON-LD and regex are fallbacks only.
* Placeholder VINs are rejected; "Similar Vehicles" is excluded.
* Pasting the same listing as ?page=1..5 scrapes it once; vehicles are
  de-duplicated across the whole batch.
"""

import csv
import io
import json
import os
import queue
import re
import threading
import time
import uuid
from dataclasses import dataclass, asdict, fields
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urldefrag

import requests
from bs4 import BeautifulSoup
from flask import Flask, Response, jsonify, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

app = Flask(__name__)
VERSION = "v3.2 (SM360 used-inventory fix + retries)"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-CA,en;q=0.9,fr-CA;q=0.8,fr;q=0.7",
}
REQUEST_DELAY = 1.2
REQUEST_TIMEOUT = 20
DEFAULT_MAX_PAGES = 40

# --------------------------------------------------------------------------
# Detail-page URL patterns. First capture group = numeric listing id.
# --------------------------------------------------------------------------
DETAIL_URL_PATTERNS = [
    ("sm360-html",
     re.compile(r"/used/[^\"'>\s]+-id(\d+)\.html", re.IGNORECASE)),

    ("sm360-inventory-path",
     # /en/used-inventory/jeep/wrangler-4xe/2022-jeep-wrangler-4xe-id38899854
     # /fr/inventaire-occasion/jeep/wrangler-4xe/jeep-wrangler-4xe-2022-id38899854
     re.compile(r"/(?:used|new|all|certified)-inventory/[a-z0-9\-]+/[a-z0-9\-]+/"
                r"[a-z0-9\-]*-id(\d+)(?=[/?#\"'>\s]|$)"
                r"|/inventaire-(?:occasion|neuf)/[a-z0-9\-]+/[a-z0-9\-]+/"
                r"[a-z0-9\-]*-id(\d+)(?=[/?#\"'>\s]|$)", re.IGNORECASE)),

    ("syncauto-numeric-path",
     re.compile(r"/(?:pre-owned|used|inventory)/(?:19|20)\d{2}/"
                r"[a-z0-9\-]+/[a-z0-9\-]+/(\d{3,9})(?=[/?#\"'>\s]|$)",
                re.IGNORECASE)),

    ("generic-id-suffix",
     re.compile(r"-id(\d{4,10})(?:\.html)?(?=[/?#\"'>\s]|$)", re.IGNORECASE)),
]

# BUG FIX #1: URLs that end in "-idNNNN" but are NOT vehicles for sale.
# On SM360 sites the footer of EVERY page has a "Honda Vehicles" block linking
# to /en/new-catalog/honda/2026-honda-accord-se-id33538 etc. (model-lineup
# brochure pages with a placeholder VIN 00000000000000000 and mileage 1).
# The old generic-id-suffix fallback happily scraped those as "inventory".
NON_VEHICLE_PATH_RE = re.compile(
    r"/(?:new-catalog|catalogue-neuf|catalog|catalogue|special-offers|"
    r"offres-speciales|promotions|news|nouvelles|form|formulaire|showroom|"
    r"services|blog|careers|carrieres)(?:/|$)", re.IGNORECASE)

PAGE_PARAM_CANDIDATES = ["page", "pageNumber", "p"]

CSV_FIELDNAMES = [
    "source_url", "domain", "url", "stock_id", "condition", "certified",
    "title", "year", "make", "model", "trim", "price",
    "mileage", "vin", "stock_number", "exterior_color", "interior_color",
    "transmission", "engine", "cylinders", "fuel_type", "drivetrain",
    "body_type", "doors", "passengers", "image_url", "extraction_notes",
]

JOBS = {}


@dataclass
class Vehicle:
    source_url: str = ""
    domain: str = ""
    url: str = ""
    stock_id: str = ""
    condition: str = ""
    certified: str = ""
    title: str = ""
    year: str = ""
    make: str = ""
    model: str = ""
    trim: str = ""
    price: str = ""
    original_price: str = ""
    mileage: str = ""
    vin: str = ""
    stock_number: str = ""
    exterior_color: str = ""
    interior_color: str = ""
    transmission: str = ""
    engine: str = ""
    cylinders: str = ""
    fuel_type: str = ""
    drivetrain: str = ""
    body_type: str = ""
    doors: str = ""
    passengers: str = ""
    image_url: str = ""
    extraction_notes: str = ""


# ==========================================================================
# helpers
# ==========================================================================
def normalize_url(raw: str) -> str:
    raw = raw.strip()
    if not raw:
        return ""
    if not raw.startswith("http"):
        raw = "https://" + raw
    return raw


_session = requests.Session()
_session.headers.update(HEADERS)
_session.headers.update({
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Cache-Control": "no-cache",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
})
_last_error = {}


def last_error() -> str:
    return _last_error.get(threading.get_ident(), "")


def get(url: str, retries: int = 3):
    """GET with retries + backoff. On failure returns None and records WHY
    (HTTP status or network error) so the log can show it."""
    err = ""
    for attempt in range(1, retries + 1):
        try:
            resp = _session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            if resp.status_code == 200:
                _last_error.pop(threading.get_ident(), None)
                return resp
            err = f"HTTP {resp.status_code}"
            if resp.status_code in (403, 429, 503):
                err += " (the site is blocking/rate-limiting requests - wait 10-15 min)"
            if resp.status_code == 404:
                break
        except requests.RequestException as e:
            err = f"{type(e).__name__}: {str(e)[:150]}"
        if attempt < retries:
            time.sleep(3 * attempt)
    _last_error[threading.get_ident()] = err
    return None


def _pattern_id(pattern, s: str) -> str:
    m = pattern.search(s) if pattern is not None else None
    if not m:
        return ""
    return next((g for g in m.groups() if g), "")


def stock_id_from_url(url: str, pattern) -> str:
    sid = _pattern_id(pattern, url)
    if not sid:
        m = re.search(r"-id(\d{4,10})", url)
        sid = m.group(1) if m else ""
    return sid


def clean_link(href: str, base_url: str) -> str:
    return urldefrag(urljoin(base_url, href))[0]


def is_vehicle_link(url: str, pattern) -> bool:
    if NON_VEHICLE_PATH_RE.search(urlparse(url).path):
        return False
    return bool(pattern.search(url))


def is_sm360(html: str) -> bool:
    h = html[:400000].lower()
    return "img.sm360.ca" in h or "360.agency" in h or "sm360" in h


def listing_key(url: str) -> str:
    """Listing URL with pagination params removed, used to spot the same
    listing pasted several times as ?page=1, ?page=2, ..."""
    p = urlparse(url)
    q = {k: v for k, v in parse_qs(p.query).items()
         if k not in PAGE_PARAM_CANDIDATES}
    qs = urlencode(sorted(q.items()), doseq=True)
    return f"{p.netloc.lower()}{p.path.rstrip('/').lower()}?{qs}"


def _all_links(html: str, base_url: str):
    soup = BeautifulSoup(html, "html.parser")
    return [clean_link(a["href"], base_url) for a in soup.find_all("a", href=True)]


def detect_platform_and_links(html: str, base_url: str):
    """Pick the detail-URL pattern that matches the MOST links on the page
    (ignoring catalog/news/offer links). Returns (name, regex, set)."""
    hrefs = _all_links(html, base_url)
    best = (None, None, set())
    for name, pattern in DETAIL_URL_PATTERNS:
        links = {h for h in hrefs if is_vehicle_link(h, pattern)}
        if len(links) > len(best[2]):
            best = (name, pattern, links)
    return best


def find_detail_links_with_pattern(html: str, base_url: str, pattern) -> set:
    return {h for h in _all_links(html, base_url) if is_vehicle_link(h, pattern)}


# ---------- listing scope: what did the user actually ask for? ----------
USED_MARKERS = ("used", "occasion", "pre-owned", "preowned", "certified",
                "certifi", "usag", "d-occasion")
NEW_MARKERS = ("new-inventory", "inventaire-neuf", "/new/", "/neuf/",
               "hybrid-electric-inventory", "electric-inventory")


def condition_of_url(url: str) -> str:
    path = urlparse(url).path.lower() + "/"
    if any(m in path for m in USED_MARKERS):
        return "used"
    if any(m in path for m in NEW_MARKERS):
        return "new"
    return ""


def listing_scope(source_url: str) -> dict:
    """Derive filters from an SM360-style listing URL, e.g.
       /en/used-inventory                -> condition=used
       /en/used-inventory/honda          -> + make=honda
       /en/certified-inventory           -> + certified only
       /en/hybrid-electric-used-inventory-> + hybrid/electric only"""
    path = urlparse(source_url).path.lower().strip("/")
    segs = [s for s in path.split("/") if s]
    scope = {"condition": condition_of_url(source_url), "make": "",
             "model": "", "certified": False, "electrified": False}
    for i, s in enumerate(segs):
        if s.endswith("inventory") or s.startswith("inventaire"):
            if "certif" in s:
                scope["certified"] = True
            if "hybrid" in s or "electri" in s:
                scope["electrified"] = True
            rest = segs[i + 1:]
            if rest:
                scope["make"] = rest[0]
            if len(rest) > 1 and not re.search(r"-id\d+$", rest[1]):
                scope["model"] = rest[1]
            break
    return scope


def link_matches_scope(url: str, scope: dict) -> bool:
    if scope["condition"]:
        lc = condition_of_url(url)
        if lc and lc != scope["condition"]:
            return False
    if scope["make"] or scope["model"]:
        segs = [s for s in urlparse(url).path.lower().split("/") if s]
        for i, s in enumerate(segs):
            if s.endswith("inventory") or s.startswith("inventaire"):
                rest = segs[i + 1:]
                if scope["make"] and (not rest or rest[0] != scope["make"]):
                    return False
                if scope["model"] and (len(rest) < 2 or rest[1] != scope["model"]):
                    return False
                break
    return True


def vehicle_matches_scope(v, scope: dict) -> bool:
    if scope["certified"] and v.certified != "yes":
        return False
    if scope["electrified"] and not re.search(
            r"hybrid|hybride|electric|électrique|electrique|phev|plug",
            f"{v.fuel_type} {v.title}", re.IGNORECASE):
        return False
    return True


def sm360_sitemap_urls(source_url: str) -> list:
    p = urlparse(source_url)
    root = f"{p.scheme}://{p.netloc}"
    first = (p.path.strip("/").split("/") or [""])[0].lower()
    en, fr = f"{root}/en/sitemap", f"{root}/fr/plan-du-site"
    return [fr, en] if first == "fr" else [en, fr]


# ==========================================================================
# detail-page parsing
# ==========================================================================
FIELD_STOP_PREFIXES = (
    "interior", "exterior", "ext", "int", "colo", "vin", "stock", "transmiss",
    "engine", "mileage", "odomet", "door", "passenger", "fuel", "drivetrain",
    "drive", "body", "trim", "price", "km", "cylind", "seat", "gear", "config",
)


def regex_field(text: str, *labels: str, max_words: int = 6) -> str:
    for label in labels:
        m = re.search(re.escape(label) + r"\s*[:#]?\s*", text, re.IGNORECASE)
        if not m:
            continue
        collected = []
        for w in text[m.end():].split()[:max_words]:
            if w.strip(".,:#").lower().startswith(FIELD_STOP_PREFIXES):
                break
            collected.append(w)
        val = " ".join(collected).strip().rstrip(".,:#-")
        if val:
            return val
    return ""


def regex_token_field(text: str, *labels: str) -> str:
    for label in labels:
        m = re.search(re.escape(label) + r"\s*[:#]?\s*([A-Za-z0-9\-]{2,20})",
                      text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


def valid_vin(vin: str) -> str:
    """BUG FIX #2: reject placeholder VINs such as 00000000000000000."""
    vin = (vin or "").strip().upper()
    if not re.fullmatch(r"[A-HJ-NPR-Z0-9]{11,17}", vin):
        return ""
    if len(set(vin)) <= 2:          # 000..., 111..., XXXX...
        return ""
    return vin


def digits(s: str) -> str:
    return re.sub(r"[^\d]", "", s or "")


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
            pool = [item] + [s for s in item.get("@graph", []) if isinstance(s, dict)]
            for sub in pool:
                t = sub.get("@type", "")
                t = t if isinstance(t, list) else [t]
                if any(x in ("Vehicle", "Car", "Product") for x in t):
                    return sub
    return {}


TITLE_CUT_MARKERS = ["|", " – ", " — ", " - Stock", " Stock #", " For Sale",
                     " for sale", " à vendre", " A vendre", " $"]


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
    words = remainder.strip(" -|").split()
    kept, length = [], 0
    for w in words:
        if kept and length + len(w) + 1 > max_chars:
            break
        kept.append(w)
        length += len(w) + 1
    return " ".join(kept)


def _na(value):
    if value is None:
        return ""
    s = str(value).strip()
    return "" if s.lower() in ("n/a", "na", "none", "null") else s


# ---- SM360 "Specifications" block -------------------------------------
# Rendered server-side on every SM360 detail page, e.g.
#   Stock # 2463U VIN # 1C4JJXP68NW108455 Fuel Plug In Hybrid (PHEV)
#   Ext. Color Green Drivetrain 4x4 Int. color Black Trim Unlimited Sahara
#   Transmission Automatic Mileage 57841 Bodystyle SUV Doors 4 ...
SM360_SPEC_LABELS = {
    "Stock #": "stock_number", "Stock#": "stock_number",
    "No. d'inventaire": "stock_number", "# d'inventaire": "stock_number",
    "Inventaire #": "stock_number", "No d'inventaire": "stock_number",
    "VIN #": "vin", "NIV #": "vin", "VIN": "vin", "NIV": "vin",
    "Fuel": "fuel_type", "Carburant": "fuel_type",
    "Ext. Color": "exterior_color", "Ext. Colour": "exterior_color",
    "Ext. color": "exterior_color", "Couleur ext.": "exterior_color",
    "Int. color": "interior_color", "Int. Color": "interior_color",
    "Int. Colour": "interior_color", "Couleur int.": "interior_color",
    "Drivetrain": "drivetrain", "Rouage": "drivetrain",
    "Entraînement": "drivetrain",
    "Trim": "trim", "Version": "trim",
    "Transmission": "transmission",
    "Mileage": "mileage", "Kilométrage": "mileage", "Odometer": "mileage",
    "Bodystyle": "body_type", "Body Style": "body_type",
    "Carrosserie": "body_type",
    "Doors": "doors", "Portes": "doors",
    "Passengers": "passengers", "Passagers": "passengers",
    "Cylinders": "cylinders", "Cylindres": "cylinders",
    "Engine": "engine", "Moteur": "engine",
    # labels we don't keep, but must recognise as boundaries
    "Frame": None, "Châssis": None, "Previous Use": None,
    "Previous Owner": None, "Usage antérieur": None,
    "Propriétaire antérieur": None, "Category": None, "Catégorie": None,
    "Condition": None, "Warranty": None, "Garantie": None,
    "Model code": None, "Code de modèle": None,
}
_SPEC_RE = re.compile(
    r"(?<![A-Za-zÀ-ÿ])(" +
    "|".join(re.escape(k) for k in sorted(SM360_SPEC_LABELS, key=len, reverse=True)) +
    r")(?![A-Za-zÀ-ÿ])")
SPEC_START_RE = re.compile(r"\b(Specifications|Spécifications|Caractéristiques)\b")
SPEC_END_RE = re.compile(r"\b(Options|Read more|Lire la suite|Lire plus|"
                         r"Similar Vehicles|Véhicules similaires|Equipment|Équipements)\b")
SIMILAR_RE = re.compile(r"Similar Vehicles|Véhicules similaires|"
                        r"You may also like|Vous aimerez aussi", re.IGNORECASE)


def parse_sm360_specs(main_text: str) -> dict:
    m = SPEC_START_RE.search(main_text)
    if not m:
        return {}
    block = main_text[m.end():]
    e = SPEC_END_RE.search(block)
    if e:
        block = block[:e.start()]
    hits = list(_SPEC_RE.finditer(block))
    out = {}
    for i, h in enumerate(hits):
        field = SM360_SPEC_LABELS[h.group(1)]
        end = hits[i + 1].start() if i + 1 < len(hits) else len(block)
        val = block[h.end():end].strip(" #:\u00a0-").strip()
        if field and val and field not in out:
            out[field] = val
    return out


PRICE_RE = re.compile(
    r"\$\s?\d{1,3}(?:[,\u00a0\u202f ]\d{3})+(?:\.\d{2})?"      # $35,495
    r"|\d{1,3}(?:[,\u00a0\u202f ]\d{3})+(?:,\d{2})?\s?\$")      # 35 495 $


def sm360_prices(soup: BeautifulSoup, main_text: str):
    """Returns (selling_price, original_price) as digit strings.
    SM360 shows 'Purchase Price $36,995 $35,495*' -> the struck-through
    first number is the original, the last one is what the car sells for."""
    price, original = "", ""
    m = re.search(r"(Purchase Price|Selling Price|Sale Price|Prix d'achat|"
                  r"Prix de vente|Prix)\b(.{0,160})", main_text)
    if m:
        chunk = re.split(r"\*|\(GST|\(TPS|\+ ?\(|taxes", m.group(2))[0]
        found = [digits(x) for x in PRICE_RE.findall(chunk)]
        found = [f for f in found if len(f) >= 3]
        if found:
            price = found[-1]
            if len(found) > 1 and found[0] != found[-1]:
                original = found[0]
    if not price:
        for key in ("twitter:data1",):
            tag = soup.find("meta", attrs={"name": key}) or soup.find("meta", attrs={"property": key})
            if tag and PRICE_RE.search(tag.get("content", "")):
                price = digits(PRICE_RE.search(tag["content"]).group(0))
    if not price:
        og = soup.find("meta", property="og:description")
        if og and PRICE_RE.search(og.get("content", "")):
            price = digits(PRICE_RE.search(og["content"]).group(0))
    return price, original


def year_make_model(og_title: str, url: str, pattern):
    """'2022 Jeep Wrangler 4xe' + URL /used-inventory/jeep/wrangler-4xe/...
    -> ('2022', 'Jeep', 'Wrangler 4xe'). Uses the URL's make slug to know how
    many words the make has (Land Rover, Mercedes-Benz, Alfa Romeo...)."""
    segs = [s for s in urlparse(url).path.split("/") if s]
    make_slug = model_slug = ""
    for i, s in enumerate(segs):
        if s.lower().endswith("inventory") or s.lower().startswith("inventaire"):
            if len(segs) > i + 2:
                make_slug, model_slug = segs[i + 1], segs[i + 2]
            break
    year = make = model = ""
    toks = (og_title or "").split()
    ym = next((t for t in toks if re.fullmatch(r"(19|20)\d{2}", t)), "")
    if ym:
        year = ym
        toks.remove(ym)
    if make_slug and toks:
        n = len(make_slug.split("-"))
        # Mercedes-Benz is one token in the title but two in the slug
        if "-" in toks[0] and toks[0].lower() == make_slug.lower():
            n = 1
        make = " ".join(toks[:n])
        model = " ".join(toks[n:])
    if not make and make_slug:
        make = make_slug.replace("-", " ").title()
    if not model and model_slug:
        model = model_slug.replace("-", " ").title()
    if not year:
        m = re.search(r"-((?:19|20)\d{2})-|/((?:19|20)\d{2})-", url)
        year = (m.group(1) or m.group(2)) if m else ""
    return year, make, model


def parse_detail_page(html: str, url: str, domain: str, source_url: str,
                      detail_pattern) -> Vehicle:
    soup = BeautifulSoup(html, "html.parser")
    v = Vehicle(source_url=source_url, domain=domain, url=url)
    v.stock_id = stock_id_from_url(url, detail_pattern)
    v.condition = condition_of_url(url)
    notes = []

    full_text = soup.get_text(" ", strip=True)
    # BUG FIX #3: ignore the "Similar Vehicles" carousel at the bottom of the
    # page -- it contains OTHER cars' VINs, stock #s and prices.
    sim = SIMILAR_RE.search(full_text)
    main_text = full_text[:sim.start()] if sim else full_text
    h1 = soup.find("h1")
    if h1:
        h1_txt = h1.get_text(" ", strip=True)
        idx = main_text.find(h1_txt)
        if idx > 0:
            main_text_from_h1 = main_text[idx:]
        else:
            main_text_from_h1 = main_text
    else:
        main_text_from_h1 = main_text

    # ---- 1. SM360 specifications block (visible, most reliable) ----------
    specs = parse_sm360_specs(main_text_from_h1)
    if specs:
        for k, val in specs.items():
            setattr(v, k, val)
        notes.append("sm360-specs")

    # ---- 2. JSON-LD fills anything still blank ---------------------------
    ld = extract_jsonld_vehicle(soup)
    if ld:
        def fill(attr, val):
            val = _na(val)
            if val and not getattr(v, attr):
                setattr(v, attr, val)
        fill("title", ld.get("name"))
        offers = ld.get("offers")
        if isinstance(offers, list) and offers:
            offers = offers[0]
        if isinstance(offers, dict):
            fill("price", digits(str(offers.get("price", ""))))
        mfo = ld.get("mileageFromOdometer")
        fill("mileage", mfo.get("value") if isinstance(mfo, dict) else mfo)
        brand = ld.get("brand")
        fill("make", brand.get("name") if isinstance(brand, dict) else brand)
        model = ld.get("model")
        fill("model", model if isinstance(model, str) else (model or {}).get("name", ""))
        fill("body_type", ld.get("bodyType"))
        fill("fuel_type", ld.get("fuelType"))
        fill("drivetrain", _na(ld.get("driveWheelConfiguration", "")).replace(
            "https://schema.org/", "").replace("http://schema.org/", "").replace("Configuration", ""))
        fill("transmission", ld.get("vehicleTransmission"))
        fill("exterior_color", ld.get("color"))
        fill("interior_color", ld.get("vehicleInteriorColor"))
        fill("trim", ld.get("vehicleConfiguration"))
        fill("vin", ld.get("vehicleIdentificationNumber"))
        fill("year", ld.get("modelDate") or ld.get("productionDate"))
        img = ld.get("image")
        if isinstance(img, list) and img:
            img = img[0]
        if isinstance(img, dict):
            img = img.get("url", "")
        if isinstance(img, str):
            fill("image_url", img)
        engines = ld.get("vehicleEngine")
        if isinstance(engines, list) and engines and isinstance(engines[0], dict):
            fill("fuel_type", engines[0].get("fuelType"))
            fill("engine", engines[0].get("engineType"))
        notes.append("json-ld")

    # ---- 3. meta tags / title ---------------------------------------------
    title_tag = soup.find("title")
    if title_tag and not v.title:
        v.title = clean_title(title_tag.get_text(strip=True))
    elif v.title:
        v.title = clean_title(v.title)
    og_img = soup.find("meta", property="og:image")
    if og_img and not v.image_url:
        v.image_url = og_img.get("content", "")
    og_title = soup.find("meta", property="og:title")
    og_title = og_title.get("content", "") if og_title else ""

    if is_sm360(html) or specs:
        price, original = sm360_prices(soup, main_text_from_h1)
        if price:
            v.price = price             # visible selling price beats JSON-LD
        v.original_price = original
        y, mk, md = year_make_model(og_title or (h1.get_text(" ", strip=True) if h1 else ""),
                                    url, detail_pattern)
        v.year = v.year or y
        v.make = v.make or mk
        v.model = v.model or md

    # ---- 4. regex fallback on the MAIN text only --------------------------
    used_regex = False
    if not v.price:
        pm = PRICE_RE.search(main_text_from_h1)
        if pm:
            v.price = digits(pm.group(0)); used_regex = True
    if not v.mileage:
        mm = re.search(r"(\d{1,3}(?:[,\s\u00a0]\d{3})+|\d{3,7})\s?km\b", main_text_from_h1, re.IGNORECASE)
        if mm:
            v.mileage = mm.group(1); used_regex = True
    if not valid_vin(v.vin):
        vinm = re.search(r"\b(?:VIN|NIV)\s*[#:]?\s*([A-HJ-NPR-Z0-9]{17})\b", main_text, re.IGNORECASE)
        v.vin = vinm.group(1) if vinm else ""
        used_regex = used_regex or bool(vinm)
    if not v.stock_number:
        v.stock_number = regex_token_field(
            main_text, "Stock is #", "Stock #", "Stock#", "Stock number", "Stock No",
            "No. d'inventaire", "# d'inventaire")
        used_regex = used_regex or bool(v.stock_number)
    for attr, labels in (
        ("transmission", ("Transmission",)),
        ("engine", ("Engine", "Moteur")),
        ("drivetrain", ("Drivetrain", "Rouage")),
        ("body_type", ("Bodystyle", "Body Style", "Body type", "Carrosserie")),
        ("exterior_color", ("Ext. Colors", "Ext. Color", "Exterior Colour",
                            "Exterior Color", "Ext Color", "Couleur ext.")),
        ("interior_color", ("Int. Colors", "Int. color", "Interior Colour",
                            "Interior Color", "Int Color", "Couleur int.")),
    ):
        if not getattr(v, attr):
            val = regex_field(main_text_from_h1, *labels)
            if val:
                setattr(v, attr, val); used_regex = True
    if not v.fuel_type:
        v.fuel_type = regex_field(main_text_from_h1, "Fuel", "Carburant", max_words=3)

    if v.title and not (v.year and v.make and v.model):
        ymm = re.match(r"(\d{4})\s+([A-Za-z\-]+)\s+([A-Za-z0-9\-]+)(.*)", v.title)
        if ymm:
            v.year = v.year or ymm.group(1)
            v.make = v.make or ymm.group(2)
            v.model = v.model or ymm.group(3)
            v.trim = v.trim or cap_trim(ymm.group(4))

    # ---- normalise ----------------------------------------------------------
    v.vin = valid_vin(v.vin)
    v.mileage = digits(v.mileage)
    v.price = digits(v.price)
    v.stock_number = v.stock_number.strip(" #:")

    # certified? (Honda/Toyota/etc. CPO badge, or "Certified/Certifié" in title/trim)
    main_html = html
    sim_html = SIMILAR_RE.search(html)
    h1_pos = html.lower().find("<h1")
    if sim_html:
        main_html = html[max(h1_pos, 0):sim_html.start()]
    if (re.search(r"<img[^>]+certified", main_html, re.IGNORECASE)
            or re.search(r"certifi(?:ed|é|e)\b", f"{v.title} {v.trim}", re.IGNORECASE)):
        v.certified = "yes"

    if used_regex:
        notes.append("regex-fallback")
    v.extraction_notes = "+".join(notes) or "regex-fallback"
    return v


def build_page_url(source_url: str, param: str, page_num: int) -> str:
    parsed = urlparse(source_url)
    query = parse_qs(parsed.query)
    for other in PAGE_PARAM_CANDIDATES:
        query.pop(other, None)
    query[param] = [str(page_num)]
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}?{urlencode(query, doseq=True)}"


# ==========================================================================
# per-URL scrape
# ==========================================================================
def scrape_one_url(source_url: str, max_pages: int, log_fn, progress_fn,
                   seen_vehicle_keys: set) -> tuple:
    domain = urlparse(source_url).netloc
    vehicles = []
    scope = listing_scope(source_url)

    log_fn(f"Scraper {VERSION}")
    log_fn(f"Fetching listing page: {source_url}")
    resp = get(source_url)
    time.sleep(REQUEST_DELAY)
    if resp is None:
        log_fn(f"Could not fetch the listing page: {last_error()}", is_err=True)
        log_fn("Trying the dealer's sitemap directly instead...")
        html = "sm360"          # assume SM360; the sitemap check below confirms it
    else:
        html = resp.text
    platform, pattern, links = detect_platform_and_links(html, source_url)
    links = {l for l in links if link_matches_scope(l, scope)}
    sm360 = is_sm360(html)

    # BUG FIX #4: SM360 listing pages render their vehicle cards with
    # JavaScript, so the raw HTML contains 0 cars and ?page=N returns the
    # same empty shell every time. Every SM360 site publishes an HTML
    # sitemap (/en/sitemap, /fr/plan-du-site) listing EVERY in-stock new
    # and used vehicle with its real detail URL -> use that.
    if sm360:
        platform = "sm360"
        pattern = dict(DETAIL_URL_PATTERNS)["sm360-inventory-path"]
        for sm_url in sm360_sitemap_urls(source_url):
            log_fn(f"SM360 site: listing cards are JavaScript-rendered, reading sitemap {sm_url}")
            r = get(sm_url)
            time.sleep(REQUEST_DELAY)
            if r is None:
                log_fn(f"Sitemap not available: {last_error()}", is_err=True)
                continue
            sm_links = {l for l in find_detail_links_with_pattern(r.text, sm_url, pattern)
                        if link_matches_scope(l, scope)}
            listing_links = {l for l in links if pattern.search(l)}
            links = sm_links | listing_links
            if sm_links:
                break
        # the sitemap lists en+fr for the same car sometimes; keep one per id
        by_id = {}
        for l in sorted(links):
            by_id.setdefault(stock_id_from_url(l, pattern) or l, l)
        links = set(by_id.values())

    if not links:
        log_fn("No vehicle links found with any known platform pattern "
               "(SM360 sitemap, SM360 .html, SM360 inventory-path, syncauto "
               "numeric-path, generic -idNNNN).", is_err=True)
        return platform, vehicles

    desc = [scope["condition"] or "all", scope["make"], scope["model"],
            "certified-only" if scope["certified"] else "",
            "hybrid/EV-only" if scope["electrified"] else ""]
    log_fn(f"Detected platform: {platform} — found {len(links)} vehicle link(s) "
           f"[scope: {' / '.join(d for d in desc if d)}].")

    # pagination only for non-SM360 (SM360 sitemap already has everything)
    if not sm360:
        first_page = set(links)
        for param in PAGE_PARAM_CANDIDATES:
            stagnant, found_any = 0, False
            for page_num in range(2, max_pages + 1):
                page_url = build_page_url(source_url, param, page_num)
                r = get(page_url)
                time.sleep(REQUEST_DELAY)
                if r is None:
                    break
                new_links = {l for l in find_detail_links_with_pattern(r.text, page_url, pattern)
                             if link_matches_scope(l, scope)}
                if new_links == first_page:
                    log_fn(f"'{param}=' is ignored by this server (same cars as page 1).")
                    break
                before = len(links)
                links |= new_links
                if len(links) == before:
                    stagnant += 1
                    if stagnant >= 2:
                        break
                else:
                    found_any, stagnant = True, 0
                    log_fn(f"Page {page_num} ({param}={page_num}): {len(links)} total vehicle links so far")
            if found_any:
                break

    # cross-URL de-duplication within the batch
    todo = []
    for l in sorted(links):
        key = (domain.lower(), stock_id_from_url(l, pattern) or l)
        if key in seen_vehicle_keys:
            continue
        seen_vehicle_keys.add(key)
        todo.append(l)
    skipped = len(links) - len(todo)
    log_fn(f"{len(links)} unique vehicle(s) on this listing; {skipped} already "
           f"collected from an earlier URL in this batch; {len(todo)} to scrape.")

    total = len(todo)
    dropped = 0
    for i, url in enumerate(todo, 1):
        r = get(url)
        time.sleep(REQUEST_DELAY)
        if r is not None:
            v = parse_detail_page(r.text, url, domain, source_url, pattern)
            if vehicle_matches_scope(v, scope):
                vehicles.append(v)
            else:
                dropped += 1
        else:
            log_fn(f"Could not fetch vehicle page {url}: {last_error()}", is_err=True)
        progress_fn(i, total)
    if dropped:
        log_fn(f"{dropped} vehicle(s) dropped because they don't match this "
               f"listing's filter (certified / hybrid-EV).")
    return platform, vehicles


def run_batch_job(job_id: str, urls: list, max_pages: int):
    q = JOBS[job_id]["queue"]
    url_jobs = JOBS[job_id]["url_jobs"]
    n_urls = len(urls)
    seen_vehicle_keys = set()
    seen_listings = {}

    for idx, source_url in enumerate(urls):
        url_jobs[idx]["status"] = "running"
        q.put({"type": "url_start", "url_index": idx, "url": source_url,
               "overall_index": idx + 1, "overall_total": n_urls})

        def log_fn(msg, is_err=False, _idx=idx):
            q.put({"type": "log", "url_index": _idx, "message": msg, "is_err": is_err})

        def progress_fn(current, total, _idx=idx):
            q.put({"type": "progress", "url_index": _idx, "current": current, "total": total})

        # same listing pasted as ?page=1, ?page=2 ... -> scrape it once
        key = listing_key(source_url)
        if key in seen_listings:
            first = seen_listings[key]
            log_fn(f"Same listing as URL {first + 1} (only the page number differs) "
                   f"— already fully covered, skipping.")
            url_jobs[idx]["status"] = "skipped"
            url_jobs[idx]["platform"] = url_jobs[first].get("platform") or "-"
            q.put({"type": "url_done", "url_index": idx, "count": 0,
                   "platform": url_jobs[idx]["platform"], "skipped": True,
                   "duplicate_of": first + 1})
            continue
        seen_listings[key] = idx

        try:
            platform, vehicles = scrape_one_url(source_url, max_pages, log_fn,
                                                progress_fn, seen_vehicle_keys)
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
            item = q.get()
            yield f"data: {json.dumps(item)}\n\n"
            if item.get("type") == "done":
                break

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
    writer = csv.DictWriter(buf, fieldnames=CSV_FIELDNAMES, extrasaction="ignore")
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
    port = int(os.environ.get("PORT", "5003").split()[0])
    print(f"\n  Dealer scraper {VERSION}\n  Open http://127.0.0.1:{port}\n")
    app.run(debug=True, port=port, threaded=True)