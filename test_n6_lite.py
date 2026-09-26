from main import (
    get_html,
    article_date,
    official_launch_confirmed,
    product_name_from_official,
    official_description,
    extract_variants,
    official_images,
    BRANDS,
)

MODEL = "OnePlus N6 Lite"
BRAND = "ONEPLUS"

DISCOVERY_URL = (
    "https://www.ndtvprofit.com/technology/"
    "oneplus-n6-lite-4g-launched-price-variant-display-batterey-no-cost-emi-and-more-12080366"
)

print("=" * 70)
print("RETAILIX PIPELINE TEST - ONEPLUS N6 LITE")
print("=" * 70)

article_html = get_html(DISCOVERY_URL)

if not article_html:
    raise RuntimeError("Could not open the known launch article.")

print("[1/6] Launch article: OK")

dt = article_date(article_html)
print("[2/6] Article date:", dt.isoformat() if dt else "NOT DETECTED")

domains = BRANDS[BRAND]

official_url, official_html = official_launch_confirmed(
    MODEL,
    BRAND,
    domains,
    article_html,
)

if not official_url or not official_html:
    print("[3/6] Official India page: FAILED")
    raise SystemExit(1)

print("[3/6] Official India page: OK")
print("       ", official_url)

name = product_name_from_official(official_html, MODEL)
print("[4/6] Official product name:", name)

description = official_description(official_html, name)
print("[5/6] Official description:", "OK" if len(description) >= 60 else "FAILED")

variants = extract_variants(official_html, name)
print("[6/6] Exact Color + Storage + RAM variants:")
for color, storage, ram in variants:
    print(f"       {color} | {storage} | {ram}")

images = official_images(official_html)
print("Official images found:", len(images))

print()
print("=" * 70)
print("TEST RESULT")
print("=" * 70)

if variants and images and len(description) >= 60:
    print("SUCCESS - Official product pipeline is working.")
else:
    print("FAILED - The pipeline stopped before complete product data.")
    if not description or len(description) < 60:
        print("Problem: official description")
    if not variants:
        print("Problem: exact Color + Storage + RAM extraction")
    if not images:
        print("Problem: official image extraction")
