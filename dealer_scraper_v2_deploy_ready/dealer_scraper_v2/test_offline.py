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


# ---------------------------------------------------------------------
# 6) Valleyfield Honda (SM360 / 360.Agency) -- the v2 failure case.
#    Listing page is a JS shell: no cars, only footer new-catalog links.
# ---------------------------------------------------------------------
import app as appmod

VF = "https://www.valleyfieldhonda.com"
vf_listing = f"""<html><head><meta property="og:image" content="https://img.sm360.ca/x.png"></head><body>
<a href="{VF}/en/used-inventory">Pre-Owned</a>
<a href="{VF}/en/new-catalog/honda/2026-honda-accord-se-id33538">Accord</a>
<a href="{VF}/en/new-catalog/honda/2025-honda-accord-sedan-se-id29470">Accord Sedan</a>
<a href="{VF}/en/special-offers/2026-honda-accord/217221">offer</a>
Powered and developed by 360.Agency</body></html>"""

vf_sitemap = f"""<html><body>img.sm360.ca
<a href="{VF}/en/new-catalog/honda/2026-honda-civic-si-base-id31537">catalog</a>
<a href="{VF}/en/new-inventory/honda/cr-v/2026-honda-cr-v-id38532685">new crv</a>
<a href="{VF}/en/used-inventory/jeep/wrangler-4xe/2022-jeep-wrangler-4xe-id38899854">jeep</a>
<a href="{VF}/en/used-inventory/honda/cr-v/2022-honda-cr-v-id39252329">crv used</a>
<a href="{VF}/en/used-inventory/land-rover/range-rover-sport/2025-land-rover-range-rover-sport-id39040994">rr</a>
<a href="{VF}/en/news/view/excellent-service/129037">news</a>
</body></html>"""

def vf_detail(title, og_title, stock, vin, prices, specs, certified_badge=False):
    badge = '<img src="https://img.sm360.ca/ir/w160h80/images/manufacturer/honda/honda-certified-color-best-en.png">' if certified_badge else ""
    return f"""<html><head><title>{title} | #{stock} | Valleyfield Honda in Salaberry-de-Valleyfield</title>
<meta property="og:title" content="{og_title}">
<meta property="og:image" content="https://img.sm360.ca/images/inventory/valleyfield-honda/{stock}.jpg">
<meta property="og:description" content="{prices[-1]} - 57,841 KM">
<script type="application/ld+json">{{"@type":"Car","vehicleIdentificationNumber":"00000000000000000","mileageFromOdometer":{{"value":"1"}}}}</script>
</head><body>
<nav>Certified Inventory  Honda Vehicles</nav>
<div>VIN #{vin}</div><div>stock # {stock}</div>
<h1>{og_title}</h1>{badge}
<div>Purchase Price</div><div>{' '.join('<span>'+p+'</span>' for p in prices)}*</div>
<div>Dealer Price Reduction Save up to $1,500</div>
<div>(GST/QST), licensing, insurance &amp; registration not included.</div>
<h3>Specifications</h3>
<ul>{''.join(f'<li><span>{k}</span> <span>{v}</span></li>' for k, v in specs)}</ul>
<h3>Options</h3><ul><li>Driver Air Bag</li></ul>
<h2>Similar Vehicles That May Interest You</h2>
<div>VIN 1C4JJXR60NW259724 Stock # 2456U 58602 km Purchase Price $37,995 $37,495*</div>
<img src="https://img.sm360.ca/honda-certified-color-best-en.png">
</body></html>"""

jeep = vf_detail("2022 Jeep Wrangler 4xe Unlimited Sahara 4XE HYBRIDE RECHARGEABLE",
    "2022 Jeep Wrangler 4xe", "2463U", "1C4JJXP68NW108455", ["$36,995", "$35,495"],
    [("Stock #","2463U"),("VIN #","1C4JJXP68NW108455"),("Fuel","Plug In Hybrid (PHEV)"),
     ("Ext. Color","Green"),("Drivetrain","4x4"),("Int. color","Black"),
     ("Trim","Unlimited Sahara 4XE HYBRIDE RECHARGEABLE"),("Transmission","Automatic"),
     ("Mileage","57841"),("Bodystyle","SUV"),("Doors","4"),("Passengers","5"),
     ("Cylinders","4 Cylinder"),("Engine","2.0L - 4 cyl."),("Frame","Sport Utility"),
     ("Previous Use","Passenger car")])
crv = vf_detail("2022 Honda CR-V EX-L IMPECCABLE!", "2022 Honda CR-V", "26173L",
    "2HKRW2H81NH206014", ["$31,995"],
    [("Stock #","26173L"),("VIN #","2HKRW2H81NH206014"),("Fuel","Gasoline"),
     ("Trim","EX-L"),("Transmission","CVT"),("Mileage","34128"),("Drivetrain","All Wheel Drive")],
    certified_badge=True)
rr = vf_detail("2025 Land Rover Range Rover Sport Autobiography", "2025 Land Rover Range Rover Sport",
    "R100U", "SALWA2FE1SA000123", ["$159,995"],
    [("Stock #","R100U"),("Fuel","Plug In Hybrid (PHEV)"),("Mileage","12000")])

PAGES = {
    f"{VF}/en/used-inventory": vf_listing,
    f"{VF}/en/certified-inventory": vf_listing,
    f"{VF}/en/sitemap": vf_sitemap,
    f"{VF}/en/used-inventory/jeep/wrangler-4xe/2022-jeep-wrangler-4xe-id38899854": jeep,
    f"{VF}/en/used-inventory/honda/cr-v/2022-honda-cr-v-id39252329": crv,
    f"{VF}/en/used-inventory/land-rover/range-rover-sport/2025-land-rover-range-rover-sport-id39040994": rr,
}
class _R:
    def __init__(self, t): self.text = t; self.status_code = 200
requested = []
def fake_get(url):
    requested.append(url)
    return _R(PAGES[url]) if url in PAGES else None
appmod.get = fake_get
appmod.REQUEST_DELAY = 0

print("Test 6: Valleyfield Honda (SM360, JS-rendered listing) end-to-end")
logs = []
plat, vs = appmod.scrape_one_url(f"{VF}/en/used-inventory", 5,
                                 lambda m, is_err=False: logs.append(m),
                                 lambda c, t: None, set())
by_stock = {v.stock_number: v for v in vs}
check("platform reported as sm360", plat == "sm360")
check("exactly the 3 used vehicles (no catalog / new / news links)", len(vs) == 3)
check("no new-catalog URLs scraped", not any("new-catalog" in u for u in requested))
check("no ?page= requests wasted on SM360", not any("page=" in u for u in requested))
j = by_stock.get("2463U")
check("Jeep found by stock #", j is not None)
if j:
    check("Jeep VIN from specs, placeholder rejected", j.vin == "1C4JJXP68NW108455")
    check("Jeep selling price 35495", j.price == "35495")
    check("Jeep original price 36995", j.original_price == "36995")
    check("Jeep mileage 57841 (not 1)", j.mileage == "57841")
    check("Jeep year/make/model", (j.year, j.make, j.model) == ("2022", "Jeep", "Wrangler 4xe"))
    check("Jeep trim", j.trim == "Unlimited Sahara 4XE HYBRIDE RECHARGEABLE")
    check("Jeep fuel", j.fuel_type == "Plug In Hybrid (PHEV)")
    check("Jeep colours", (j.exterior_color, j.interior_color) == ("Green", "Black"))
    check("Jeep cylinders / engine", (j.cylinders, j.engine) == ("4 Cylinder", "2.0L - 4 cyl."))
    check("Jeep doors/passengers/body", (j.doors, j.passengers, j.body_type) == ("4", "5", "SUV"))
    check("Jeep condition used", j.condition == "used")
    check("Jeep NOT certified (similar-vehicles badge ignored)", j.certified == "")
    check("Jeep title cleaned", j.title == "2022 Jeep Wrangler 4xe Unlimited Sahara 4XE HYBRIDE RECHARGEABLE")
c = by_stock.get("26173L")
check("CR-V certified detected", c is not None and c.certified == "yes")
check("CR-V single price", c is not None and c.price == "31995" and c.original_price == "")
r = by_stock.get("R100U")
check("Land Rover two-word make", r is not None and (r.make, r.model) == ("Land Rover", "Range Rover Sport"))

print("Test 7: certified-inventory listing keeps only certified cars")
plat, vs = appmod.scrape_one_url(f"{VF}/en/certified-inventory", 5,
                                 lambda m, is_err=False: None, lambda c, t: None, set())
check("only the certified CR-V", [v.stock_number for v in vs] == ["26173L"])

print("Test 8: ?page=1..5 of the same listing collapse into one")
check("listing_key ignores page param",
      appmod.listing_key(f"{VF}/en/used-inventory?page=2") == appmod.listing_key(f"{VF}/en/used-inventory?page=5"))
check("make sub-listing scope", appmod.listing_scope(f"{VF}/en/used-inventory/honda")["make"] == "honda")
check("link scope filters other makes",
      not appmod.link_matches_scope(f"{VF}/en/used-inventory/jeep/wrangler/2022-jeep-wrangler-id1234",
                                    appmod.listing_scope(f"{VF}/en/used-inventory/honda")))

print()
print(f"CSV schema has {len(CSV_FIELDNAMES)} columns: {CSV_FIELDNAMES}")
print()
print(f"RESULT: {PASS} passed, {FAIL} failed")
if FAIL:
    raise SystemExit(1)
