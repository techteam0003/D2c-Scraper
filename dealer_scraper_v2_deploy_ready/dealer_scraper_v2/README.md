# Multi-Platform Dealer Inventory Scraper — local web app (v3)

## v3: what was broken, and the fix (valleyfieldhonda.com)

**Platform:** Valleyfield Honda runs on **SM360 / 360.Agency** (footer: "Powered and
developed by 360.Agency", all assets on `img.sm360.ca`, "My Transaction" bar).

**Symptoms in v2:** every `?page=N` URL returned the same 15 "vehicles" — all new
2025–2026 Hondas with VIN `00000000000000000`, mileage `1`, and stock #s like 29471.

**Root cause:**
1. SM360's listing page builds its vehicle cards with JavaScript. `requests` gets an
   empty shell (0 cars), and `?page=2..5` return the identical shell.
2. With no real cars in the HTML, the `generic-id-suffix` fallback matched the
   **footer "Honda Vehicles" links** (`/en/new-catalog/honda/2026-honda-accord-se-id33538`).
   Those are model-lineup brochure pages, not inventory — hence the placeholder VIN,
   mileage 1, and "stock #" = catalog id.
3. The VIN regex rejected the `VIN #1C4…` format, so the placeholder was never replaced.
4. Detail-page regexes also read the "Similar Vehicles" carousel (other cars' data).

**Fix:**
- SM360 is detected automatically, and the complete inventory is read from the
  dealer's own HTML sitemap (`/en/sitemap` or `/fr/plan-du-site`), which lists every
  in-stock vehicle with its real URL. It is then scoped to what your listing URL asked
  for: used vs new, `/used-inventory/{make}[/{model}]`, `certified-inventory`
  (certified only), `hybrid-electric-used-inventory` (hybrid/EV only).
- Catalog / news / special-offer / form links are never treated as vehicles.
- Detail pages are parsed from the server-rendered **Specifications** block:
  Stock #, VIN, Fuel, Ext./Int. colour, Drivetrain, Trim, Transmission, Mileage,
  Bodystyle, Doors, Passengers, Cylinders, Engine — plus the selling price and the
  original (pre-reduction) price. Works on EN and FR pages.
- Placeholder VINs rejected; "Similar Vehicles" section ignored.
- Paste just ONE listing URL. If you paste `?page=1…5`, the duplicates are skipped,
  and vehicles are de-duplicated across the whole batch.

New CSV/Excel columns: `condition`, `certified`, `original_price`, `cylinders`,
`doors`, `passengers`. `price` and `mileage` are now plain numbers.

Run on a different port: `PORT=5002 python app.py` (Windows: `set PORT=5002` first).

---


A Flask app with a browser frontend: paste one or more dealer listing-page
URLs, watch it scrape live (progress bar + log, with the detected platform
shown per URL), then browse results in a table and download CSV or Excel.

This is a v2 of the original `sm360_webapp` scraper, extended after
inspecting **Hamel Chevrolet Buick GMC** (`hamelchevrolet.ca`), which turned
out to be on a *different URL shape* of the same SM360 platform than the
original tool assumed.

## What platform is hamelchevrolet.ca on?

**SM360** — confirmed by:
- Dealer logo/photos served from `img.sm360.ca`
- `schema.org/Car` JSON-LD blocks embedded on every vehicle detail page
  (`bodyType`, `brand`, `vehicleIdentificationNumber`, `mileageFromOdometer`,
  `offers.price`, etc.)

But its URL structure differs from the SM360 pattern the original tool was
built around:

| | Original tool assumed | hamelchevrolet.ca actually uses |
|---|---|---|
| Detail page | `/used/{slug}-id{stock_id}.html` | `/en/used-inventory/{make}/{model}/{year}-{make}-{model}-id{stock_id}` (no `.html`) |
| Listing page | `/en/used-inventory` | `/en/used-inventory` (same), paginated as `?namedSorting=default&limit=12&page=N` |

Because of that mismatch, the original scraper would have found **0 vehicle
links** on hamelchevrolet.ca and reported failure — not because the site
was unscrapable, but because it's a different flavor of the same platform.

A third site referenced in your sample CSV, **garagetardif.com**, is a
*different platform entirely* (WordPress-based, images served from a
`syncauto-*.s3.amazonaws.com` bucket, not SM360), with yet another detail
URL shape: `/en/pre-owned/{year}/{make}/{model}/{numeric_id}`.

## How this version handles multiple platforms

Instead of hard-coding one URL shape, `app.py` tries a list of known
detail-page URL patterns against each listing page and uses whichever one
actually finds vehicle links, logging which "platform" it detected:

1. **`sm360-html`** — `/used/{slug}-id{id}.html` (original pattern)
2. **`sm360-inventory-path`** — `/used-inventory/{make}/{model}/{...}-id{id}` (hamelchevrolet.ca and similar Dilawri-group SM360 sites)
3. **`syncauto-numeric-path`** — `/pre-owned/{year}/{make}/{model}/{id}` (garagetardif.com and similar)
4. **`generic-id-suffix`** — catch-all for any `...-id123456` href, in case a site is SM360 but under a URL prefix not covered above

If none of these match, the tool reports 0 results for that URL rather than
guessing — check the log message, inspect one vehicle detail-page URL by
hand, and add a new regex to `DETAIL_URL_PATTERNS` in `app.py` (each entry
just needs a name and a regex whose first capture group is the numeric ID).

Pagination now preserves the listing page's original query string (e.g.
`namedSorting=default&limit=12`) instead of dropping it, since some SM360
sites need those params to return results at all.

## Setup

```bash
cd dealer_scraper_v2
pip install -r requirements.txt
python app.py
```

Then open **http://localhost:5000**.

## Using it

1. Paste one or more dealer listing-page URLs (one per line).
2. Click **Start Scraping**. You'll see, per URL:
   - The detected platform (once the first listing page is fetched).
   - A log of pagination progress and vehicle links found.
   - A progress bar as each vehicle detail page is visited.
3. When done, results appear in a table.
4. **Download CSV** gives a single CSV with the columns below (same schema
   as `garage_tardif.csv`). **Download Excel** gives a workbook with an
   "All Vehicles" sheet, a "Summary" sheet, and one sheet per source URL.

## Output columns (CSV and Excel)

```
source_url, domain, url, stock_id, condition, certified, title, year, make,
model, trim, price, original_price, mileage, vin, stock_number,
exterior_color, interior_color, transmission, engine, cylinders, fuel_type,
drivetrain, body_type, doors, passengers, image_url, extraction_notes
```

`extraction_notes` says whether a row came from JSON-LD, a regex fallback,
or both (`json-ld+regex-fallback` when JSON-LD covered most fields but a
few — like stock number, which usually isn't in the JSON-LD — needed the
regex pass). Blank fields mean the value genuinely wasn't findable on that
page; nothing is guessed.

## Before running this against real dealer sites

- Check `https://{domain}/robots.txt` and the site's Terms of Use.
- Keep request volume reasonable — `REQUEST_DELAY` in `app.py` (default
  1.2s between requests) exists to avoid hammering a dealer's server or
  triggering bot protection / IP bans. Don't lower it just to go faster.
- **This tool has not been run against live dealer sites from the
  environment it was built in (no outbound internet access there).** The
  detail-URL patterns and JSON-LD field mapping were built and unit-tested
  (`python test_offline.py`) against real HTML/JSON-LD structure retrieved
  via search for hamelchevrolet.ca, garagetardif.com, and the original
  SM360 `.html` pattern — but test it on one domain first and sanity-check
  a handful of rows before treating the output as fully reliable, especially
  for pagination edge cases (a site's actual "last page" behavior can only
  be confirmed by running against it live).

## Running the offline self-tests

```bash
python test_offline.py
```

This checks pattern detection and field extraction against synthetic HTML
snapshots of each supported platform — no network access required.
