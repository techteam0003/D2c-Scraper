"""
Offline sanity checks -- no network access needed.

Builds small synthetic HTML pages that mimic the three platforms we've
identified, and runs them through detect_platform_and_links() and
parse_detail_page() to confirm the pattern matching / extraction logic
behaves as expected before pointing this at a live site.
"""
from app import detect_platform_and_links, parse_detail_page, CSV_FIELDNAMES
from dataclasses import asdict

PASS = 0
FAIL = 0


def check(label, cond):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  OK   {label}")
    else:
        FAIL += 1
        print(f"  FAIL {label}")


# ---------------------------------------------------------------------
# 1) Hamel-Chevrolet-style SM360 listing page (new inventory-path shape)
# ---------------------------------------------------------------------
hamel_listing_html = """
<html><body>
<a href="/en/used-inventory/ford/maverick/2023-ford-maverick-id35313906">2023 Ford Maverick</a>
<a href="/en/used-inventory/gmc/sierra-2500hd/2025-gmc-sierra-2500hd-id37738716">2025 GMC Sierra</a>
<a href="/en/used-inventory/buick/encore-gx/2026-buick-encore-gx-id36490472">2026 Buick Encore GX</a>
<a href="/en/inventory/filters">not a vehicle</a>
</body></html>
"""
print("Test 1: Hamel Chevrolet (SM360, inventory-path) listing page")
platform, pattern, links = detect_platform_and_links(hamel_listing_html, "https://www.hamelchevrolet.ca/en/used-inventory")
check("platform detected as sm360-inventory-path", platform == "sm360-inventory-path")
check("found exactly 3 vehicle links", len(links) == 3)
check("absolute URL built correctly", "https://www.hamelchevrolet.ca/en/used-inventory/ford/maverick/2023-ford-maverick-id35313906" in links)

# ---------------------------------------------------------------------
# 2) Hamel-Chevrolet-style detail page with schema.org/Car JSON-LD
#    (reconstructed from the real JSON-LD fields observed via search)
# ---------------------------------------------------------------------
hamel_detail_html = """
<html><head><title>2025 GMC Sierra 2500HD SLE | Hamel Chevrolet</title>
<script type="application/ld+json">
{
    "@context": "http://schema.org/",
    "@type": "Car",
    "bodyType": "Pickup - Crew Cab",
    "brand": "GMC",
    "color": "White",
    "description": "GMC Sierra SLE 2500HD Crew Cab 4X4",
    "driveWheelConfiguration": "https://schema.org/AllWheelDriveConfiguration",
    "manufacturer": "GMC",
    "mileageFromOdometer": {"unitCode": "KMT", "value": "49492"},
    "model": "Sierra 2500HD",
    "modelDate": "2025",
    "offers": {"availability": "https://schema.org/InStock", "price": "79999", "priceCurrency": "CAD"},
    "vehicleConfiguration": "SLE Crew Cab * 4X4 * 6.6L V8 Duramax Diesel *",
    "vehicleEngine": [{"engineType": "8", "fuelType": "n/a"}],
    "vehicleIdentificationNumber": "1GT1UMEY6SF150323",
    "vehicleInteriorColor": "n/a",
    "vehicleTransmission": "Automatic"
}
</script>
</head>
<body>
Stock # H4870 VIN 1GT1UMEY6SF150323
Mileage: 49492 km
Ext. Color Summit White
Int. color Jet Black
Drivetrain All Wheel Drive
Bodystyle Pickup
</body></html>
"""
print("Test 2: Hamel Chevrolet detail page (JSON-LD schema.org/Car)")
v = parse_detail_page(
    hamel_detail_html,
    "https://www.hamelchevrolet.ca/en/used-inventory/gmc/sierra-2500hd/2025-gmc-sierra-2500hd-id37738716",
    "www.hamelchevrolet.ca",
    "https://www.hamelchevrolet.ca/en/used-inventory",
    pattern,
)
d = asdict(v)
check("stock_id extracted from URL", d["stock_id"] == "37738716")
check("make from JSON-LD", d["make"] == "GMC")
check("model from JSON-LD", d["model"] == "Sierra 2500HD")
check("year from modelDate", d["year"] == "2025")
check("price from JSON-LD offers", d["price"] == "79999")
check("mileage from JSON-LD", d["mileage"] == "49492")
check("vin from JSON-LD", d["vin"] == "1GT1UMEY6SF150323")
check("trim from vehicleConfiguration", "SLE Crew Cab" in d["trim"])
check("drivetrain cleaned of schema.org URL", d["drivetrain"] == "AllWheelDrive")
check("stock_number picked up via regex fallback", d["stock_number"] == "H4870")
check("transmission from JSON-LD", d["transmission"] == "Automatic")
check("interior_color falls back to regex since JSON-LD said n/a", d["interior_color"] == "Jet Black")
check("extraction_notes mentions json-ld", "json-ld" in d["extraction_notes"])

# ---------------------------------------------------------------------
# 3) garagetardif-style (syncauto numeric path) listing page
# ---------------------------------------------------------------------
tardif_listing_html = """
<html><body>
<a href="/en/pre-owned/2004/nissan/350z/2745">2004 Nissan 350Z</a>
<a href="/en/pre-owned/2019/buick/encore-ti-cx/2910">2019 Buick Encore TI CX</a>
<a href="/en/pre-owned/filters">not a vehicle</a>
</body></html>
"""
print("Test 3: garagetardif-style listing page")
platform3, pattern3, links3 = detect_platform_and_links(tardif_listing_html, "https://www.garagetardif.com/en/pre-owned")
check("platform detected as syncauto-numeric-path", platform3 == "syncauto-numeric-path")
check("found exactly 2 vehicle links", len(links3) == 2)

# ---------------------------------------------------------------------
# 4) Old-style sm360 .html listing page (what the original tool targeted)
# ---------------------------------------------------------------------
old_sm360_html = """
<html><body>
<a href="/used/2020-honda-civic-lx-id1234567.html">2020 Honda Civic LX</a>
<a href="/used/2019-toyota-corolla-le-id7654321.html">2019 Toyota Corolla LE</a>
</body></html>
"""
print("Test 4: legacy sm360 .html listing page")
platform4, pattern4, links4 = detect_platform_and_links(old_sm360_html, "https://www.example-dealer.com/used-inventory")
check("platform detected as sm360-html", platform4 == "sm360-html")
check("found exactly 2 vehicle links", len(links4) == 2)

# ---------------------------------------------------------------------
# 5) Unknown platform -> no links found, should not crash
# ---------------------------------------------------------------------
unknown_html = "<html><body><a href='/vehicles/random-page'>nothing here</a></body></html>"
print("Test 5: unknown / unsupported platform")
platform5, pattern5, links5 = detect_platform_and_links(unknown_html, "https://www.unknown-dealer.com/inventory")
check("no platform detected", platform5 is None)
check("no links found", len(links5) == 0)

print()
print(f"CSV schema has {len(CSV_FIELDNAMES)} columns: {CSV_FIELDNAMES}")
print()
print(f"RESULT: {PASS} passed, {FAIL} failed")
if FAIL:
    raise SystemExit(1)
