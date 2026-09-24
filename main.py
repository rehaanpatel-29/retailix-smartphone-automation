import re
import time
import shutil
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin, quote_plus, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

# ============================================================
# RETAILIX SMARTPHONE AUTOMATION
# NDTV MOBILE DISCOVERY -> OFFICIAL INDIA BRAND VERIFICATION
# ============================================================

XLSX = Path("product_import_template (1).xlsx")
IST = timezone(timedelta(hours=5, minutes=30))

BRANDS = {
    "APPLE": ["apple.com"],
    "INFINIX": ["infinixmobility.com", "infinixmobiles.in"],
    "MOTOROLA": ["motorola.in", "motorola.com"],
    "ONEPLUS": ["oneplus.in", "oneplus.com"],
    "OPPO": ["oppo.com"],
    "POCO": ["mi.com"],
    "REDMI": ["mi.com"],
    "REALME": ["realme.com"],
    "SAMSUNG": ["samsung.com"],
    "VIVO": ["vivo.com"],
}

NDTV_URLS = [
    "https://www.ndtv.com/topic/mobile-phones",
    "https://www.ndtv.com/topic/smartphones",
    "https://www.ndtv.com/topic/smartphone",
    "https://www.ndtv.com/topic/launches",
]

HEADERS = [
    "productName", "brandName", "categoryName", "subcategoryName",
    "unit_name", "taxName", "discountName", "warrantyName",
    "shortDescription", "longDescription", "is_featured", "is_digital",
    "variant_Color", "variant_Storage", "variant_RAM",
    "image1", "image1Primary", "image2", "image3", "image4"
]

SMART_WORDS = (
    "smartphone", "smart phone", "mobile phone", "phone", "iphone",
    "galaxy", "redmi", "poco", "oneplus", "nord", "oppo", "reno",
    "find x", "find n", "vivo", "realme", "narzo", "motorola",
    "moto", "edge", "razr", "infinix", "note"
)

LAUNCH_WORDS = (
    "launched", "launches", "launch", "introduced", "unveiled",
    "debuts", "announced", "available in india", "launched in india",
    "goes on sale", "sale begins", "now available", "arrives in stores"
)

BAD_WORDS = (
    "upcoming", "rumour", "rumor", "leak", "leaked", "expected",
    "may launch", "could launch", "might launch", "tipped", "teased",
    "teaser", "review", "comparison", "discount", "offer", "concept",
    "speculation", "hands-on"
)

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 Chrome/140 Safari/537.36"
})


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def now():
    return datetime.now(IST)


def normalize_url(url, base=None):
    if not url:
        return ""
    return urljoin(base or "", url).split("#")[0]


def same_domain(url, domains):
    host = urlparse(url).netloc.lower().split(":")[0]
    return any(host == d or host.endswith("." + d) for d in domains)


def get_html(url, timeout=25):
    try:
        r = SESSION.get(url, timeout=timeout, allow_redirects=True)
        if r.ok and len(r.text) > 500:
            return r.text
    except Exception:
        pass
    return ""


def parse_date(text):
    text = clean(text)
    patterns = [
        r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})",
        r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})",
        r"([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})",
    ]
    for p in patterns:
        m = re.search(p, text)
        if not m:
            continue
        try:
            if m.lastindex == 3:
                if m.group(2).isdigit():
                    return datetime.strptime(m.group(0), "%d/%m/%Y").replace(tzinfo=IST)
                if m.group(1).isdigit():
                    return datetime.strptime(m.group(0), "%d %B %Y").replace(tzinfo=IST)
                return datetime.strptime(m.group(0), "%B %d, %Y").replace(tzinfo=IST)
        except Exception:
            pass
    return None


def article_date(html):
    soup = BeautifulSoup(html, "html.parser")

    for attrs in (
        {"property": "article:published_time"},
        {"name": "publish-date"},
        {"name": "date"},
        {"itemprop": "datePublished"},
    ):
        m = soup.find("meta", attrs=attrs)
        if m:
            value = m.get("content") or m.get("datetime")
            if value:
                try:
                    dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
                    return dt.astimezone(IST)
                except Exception:
                    d = parse_date(value)
                    if d:
                        return d

    for tag in soup.find_all(["time", "span"]):
        value = tag.get("datetime") or tag.get_text(" ", strip=True)
        d = parse_date(value)
        if d:
            return d

    return None


def ndtv_candidates():
    """
    NDTV is the only discovery layer.
    We inspect its mobile-phone/launch topic pages and collect article links.
    """
    out = {}
    for base in NDTV_URLS:
        html = get_html(base)
        if not html:
            continue

        soup = BeautifulSoup(html, "html.parser")

        for a in soup.find_all("a", href=True):
            url = normalize_url(a["href"], base)
            title = clean(a.get_text(" ", strip=True))

            if not url.startswith("https://www.ndtv.com/"):
                continue
            if len(title) < 15:
                continue

            blob = title.lower()
            if not any(x in blob for x in SMART_WORDS):
                continue

            # Keep only useful launch-looking mobile stories.
            # Article verification later is stricter.
            if not any(x in blob for x in LAUNCH_WORDS):
                continue

            out[url] = title

    return [(u, t) for u, t in out.items()]


def candidate_brand(title):
    low = title.lower()

    aliases = {
        "APPLE": ["iphone", "apple"],
        "INFINIX": ["infinix"],
        "MOTOROLA": ["motorola", "moto"],
        "ONEPLUS": ["oneplus", "one plus"],
        "OPPO": ["oppo"],
        "POCO": ["poco"],
        "REDMI": ["redmi"],
        "REALME": ["realme", "narzo"],
        "SAMSUNG": ["samsung", "galaxy"],
        "VIVO": ["vivo"],
    }

    for brand, words in aliases.items():
        if any(w in low for w in words):
            return brand
    return None


def qualifying_ndtv(title, html, brand, start, end):
    low = (title + " " + BeautifulSoup(html, "html.parser").get_text(" ", strip=True)[:8000]).lower()

    if any(x in low for x in BAD_WORDS):
        return False

    aliases = {
        "APPLE": ("apple", "iphone"),
        "INFINIX": ("infinix",),
        "MOTOROLA": ("motorola", "moto"),
        "ONEPLUS": ("oneplus", "one plus"),
        "OPPO": ("oppo",),
        "POCO": ("poco",),
        "REDMI": ("redmi",),
        "REALME": ("realme", "narzo"),
        "SAMSUNG": ("samsung", "galaxy"),
        "VIVO": ("vivo",),
    }

    if not any(x in low for x in aliases[brand]):
        return False

    if not any(x in low for x in LAUNCH_WORDS):
        return False

    dt = article_date(html)
    if not dt:
        return False

    return start <= dt <= end


def model_from_ndtv(title, brand):
    title = clean(title)
    title = re.sub(
        r"\b(launched|launches|launch|introduced|unveiled|debuts|announced)\b.*$",
        "",
        title,
        flags=re.I,
    )
    title = re.sub(r"\s+in India.*$", "", title, flags=re.I)
    title = re.sub(r"\s+with .*?:.*$", "", title, flags=re.I)
    title = re.sub(r"\s*[-–—:]\s*$", "", title).strip()

    return title


def official_link_from_ndtv(ndtv_html, domains, model):
    """
    Find an official brand URL directly linked from the NDTV article.
    No search engine, marketplace, or third-party product source is used.
    """
    soup = BeautifulSoup(ndtv_html, "html.parser")
    model_words = {
        x.lower() for x in re.findall(r"[A-Za-z0-9]+", model)
        if len(x) >= 2
    }

    candidates = []

    for a in soup.find_all("a", href=True):
        u = normalize_url(a["href"], "https://www.ndtv.com/")
        if not same_domain(u, domains):
            continue

        text = clean(a.get_text(" ", strip=True))
        blob = (text + " " + u).lower()

        score = 0
        if official_page_is_india(u):
            score += 3
        if any(x in blob for x in ("phone", "smartphone", "mobile", "product")):
            score += 2

        for word in model_words:
            if word in blob:
                score += 1

        candidates.append((score, u))

    candidates.sort(reverse=True)

    return candidates[0][1] if candidates else ""


def official_launch_confirmed(model, brand, domains, ndtv_html):
    """
    Verify the launch only on the official brand domain.
    The official page must contain the model and clear launch/availability
    language for India. No Google, marketplace, or other tech site is used.
    """
    product_url = official_link_from_ndtv(ndtv_html, domains, model)

    if not product_url:
        return "", ""

    official_html = get_html(product_url)
    if not official_html:
        return "", ""

    soup = BeautifulSoup(official_html, "html.parser")
    body = clean(soup.get_text(" ", strip=True))
    low = body.lower()
    model_low = model.lower()

    if model_low not in low:
        # Try a slightly looser model match using significant words.
        words = [
            w.lower() for w in re.findall(r"[A-Za-z0-9]+", model)
            if len(w) >= 3
        ]
        if not words or not all(w in low for w in words[:2]):
            return "", ""

    launch_terms = (
        "launched in india",
        "launches in india",
        "introduced in india",
        "available in india",
        "now available in india",
        "goes on sale in india",
        "sale begins in india",
        "available from",
        "available starting",
        "available beginning",
        "launch date",
    )

    if not any(term in low for term in launch_terms):
        # Official product pages often state India availability without the
        # exact phrase above. Accept explicit India + available/launch wording.
        if "india" not in low:
            return "", ""
        if not any(x in low for x in ("launch", "available", "sale")):
            return "", ""

    if not official_page_is_india(product_url):
        # Some official India product pages use a global-looking URL.
        # Require explicit India text on the page in that case.
        if "india" not in low:
            return "", ""

    return product_url, official_html

def official_page_is_india(url):
    low = url.lower()
    return any(x in low for x in ("/in/", "india", "in."))


def official_launch_confirmed(model, brand, domains, ndtv_date):
    """
    Search official domain for the model and require India/launch wording.
    """
    q = quote_plus(f'site:{domains[0]} "{model}" launched India')
    html = get_html("https://www.google.com/search?q=" + q)

    if not html:
        return "", ""

    soup = BeautifulSoup(html, "html.parser")

    for a in soup.find_all("a", href=True):
        u = normalize_url(a["href"])
        if not same_domain(u, domains):
            continue

        text = clean(a.get_text(" ", strip=True))
        blob = (text + " " + u).lower()

        if not official_page_is_india(u):
            continue

        if any(x in blob for x in ("launch", "available", "india", "product", "phone")):
            page_html = get_html(u)
            if not page_html:
                continue

            body = clean(BeautifulSoup(page_html, "html.parser").get_text(" ", strip=True))
            body_low = body.lower()

            if model.lower() not in body_low:
                continue

            if "india" not in body_low and not official_page_is_india(u):
                continue

            return u, page_html

    return "", ""


def product_name_from_official(html, fallback):
    soup = BeautifulSoup(html, "html.parser")

    for selector in ("h1",):
        tag = soup.select_one(selector)
        if tag:
            value = clean(tag.get_text(" ", strip=True))
            if 2 <= len(value) <= 120:
                return value

    for attrs in (
        {"property": "og:title"},
        {"name": "twitter:title"},
    ):
        m = soup.find("meta", attrs=attrs)
        if m and m.get("content"):
            value = clean(m["content"])
            value = re.sub(r"\s*[-|]\s*(Official.*|Buy.*|India.*)$", "", value, flags=re.I)
            if len(value) >= 2:
                return value

    return clean(fallback)


def official_description(html, model):
    soup = BeautifulSoup(html, "html.parser")

    for attrs in (
        {"name": "description"},
        {"property": "og:description"},
        {"name": "twitter:description"},
    ):
        m = soup.find("meta", attrs=attrs)
        if m and m.get("content"):
            text = clean(m["content"])
            if len(text) >= 60:
                return text[:2000]

    paragraphs = []
    for p in soup.find_all(["p", "li"]):
        text = clean(p.get_text(" ", strip=True))
        low = text.lower()

        if len(text) < 60 or len(text) > 1200:
            continue
        if model.lower() not in low and "smartphone" not in low and "phone" not in low:
            continue
        paragraphs.append(text)

    if paragraphs:
        return " ".join(paragraphs[:4])[:2000]

    return ""


def normalize_capacity(value):
    value = clean(value).upper().replace(" ", "")
    return value


def extract_variants(html, model):
    """
    STRICT RULE:
    A row is created only when Color + Storage + RAM can be mapped
    to the same explicit variant. No Cartesian product.
    """
    soup = BeautifulSoup(html, "html.parser")
    text = clean(soup.get_text(" ", strip=True))

    found = []

    # 1. Structured JSON / HTML text containing explicit RAM + storage.
    # Common forms: 8GB RAM / 256GB, 12GB+256GB, 8 GB + 128 GB.
    pattern = re.compile(
        r"(?P<ram>\d{1,2})\s*GB\s*(?:RAM)?\s*(?:\+|/|\||,|with)\s*"
        r"(?P<storage>\d{2,4})\s*(?P<unit>GB|TB)",
        re.I
    )

    pairs = []
    for m in pattern.finditer(text):
        ram = f"{m.group('ram')} GB"
        storage = f"{m.group('storage')} {m.group('unit').upper()}"
        pair = (ram, storage)
        if pair not in pairs:
            pairs.append(pair)

    # 2. Explicit variant blocks. We only accept a color when the same
    # small block contains the RAM/storage pair.
    blocks = soup.find_all(["li", "div", "section", "article", "label", "option"])

    color_re = re.compile(
        r"(?:colour|color|colourway|colorway)\s*[:\-]?\s*"
        r"([A-Za-z][A-Za-z0-9 &'’()./\-]{2,45})",
        re.I
    )

    for block in blocks:
        block_text = clean(block.get_text(" ", strip=True))
        if len(block_text) > 500:
            continue

        pm = pattern.search(block_text)
        cm = color_re.search(block_text)

        if pm and cm:
            ram = f"{pm.group('ram')} GB"
            storage = f"{pm.group('storage')} {pm.group('unit').upper()}"
            color = clean(cm.group(1))

            # Remove obvious trailing labels.
            color = re.split(
                r"\s+(?:ram|storage|memory|price|₹|\|)\b",
                color,
                flags=re.I
            )[0].strip()

            item = (color, storage, ram)
            if item not in found:
                found.append(item)

    # 3. Common product JSON-LD / data attributes.
    for tag in soup.find_all(["script"]):
        raw = tag.string or tag.get_text()
        if not raw or len(raw) > 500000:
            continue

        for pm in pattern.finditer(raw):
            ram = f"{pm.group('ram')} GB"
            storage = f"{pm.group('storage')} {pm.group('unit').upper()}"

            before = raw[max(0, pm.start() - 1000):pm.start()]
            after = raw[pm.end():pm.end() + 1000]
            nearby = before + " " + after

            cm = color_re.search(nearby)
            if cm:
                color = clean(cm.group(1))
                item = (color, storage, ram)
                if item not in found:
                    found.append(item)

    # Deduplicate, normalize, and reject junk colors.
    clean_found = []
    bad_colors = {
        "black", "white", "blue", "green", "red", "silver",
        "grey", "gray", "gold", "pink", "purple"
    }

    for color, storage, ram in found:
        color = clean(color)
        storage = normalize_capacity(storage)
        ram = normalize_capacity(ram)

        if len(color) < 2 or len(color) > 45:
            continue

        # Do not reject generic color names: official pages may legitimately
        # use them. Only reject obvious UI words.
        if color.lower() in {"select", "choose", "available", "variant", "color"}:
            continue

        item = (color, storage, ram)
        if item not in clean_found:
            clean_found.append(item)

    return clean_found


def official_images(html):
    soup = BeautifulSoup(html, "html.parser")
    out = []

    for tag in soup.find_all(["meta", "img", "source"]):
        url = ""

        if tag.name == "meta":
            if tag.get("property") in ("og:image", "twitter:image"):
                url = tag.get("content", "")
        elif tag.name == "img":
            url = tag.get("src") or tag.get("data-src") or tag.get("data-lazy-src") or ""
        else:
            url = tag.get("srcset", "").split(",")[0].strip().split(" ")[0]

        url = normalize_url(url)

        if not url.startswith("http"):
            continue

        low = url.lower()

        if any(x in low for x in (
            "logo", "icon", "favicon", "sprite", "swatch",
            "flag", "avatar", "banner", "tracking"
        )):
            continue

        if url not in out:
            out.append(url)

        if len(out) >= 4:
            break

    return out


def backup_excel():
    stamp = now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(f"{XLSX.stem}_backup_{stamp}{XLSX.suffix}")
    shutil.copy2(XLSX, backup)
    print("Backup:", backup.name)


def main():
    print("=" * 70)
    print("RETAILIX DAILY SMARTPHONE AUTOMATION")
    print("NDTV MOBILE DISCOVERY + OFFICIAL INDIA BRAND VERIFICATION ONLY")
    print("=" * 70)

    if not XLSX.exists():
        raise FileNotFoundError(XLSX)

    wb = load_workbook(XLSX)
    ws = wb["PRODUCT_DATA"]

    headers = [ws.cell(1, c).value for c in range(1, 21)]
    if headers != HEADERS:
        raise RuntimeError("RetailiX template headers do not match expected 20 columns.")

    cols = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}

    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        existing.add(tuple(
            str(row[cols[k] - 1] or "").strip().lower()
            for k in (
                "productName", "brandName",
                "variant_Color", "variant_Storage", "variant_RAM"
            )
        ))

    end = now()
    start = end - timedelta(hours=24)

    print(f"Window: {start.isoformat()} -> {end.isoformat()}")
    print("Existing combinations:", len(existing))

    candidates = ndtv_candidates()
    print("NDTV candidates found:", len(candidates))

    # Keep only the ten requested brands.
    grouped = {brand: [] for brand in BRANDS}

    for url, title in candidates:
        brand = candidate_brand(title)
        if brand and url not in [x[0] for x in grouped[brand]]:
            grouped[brand].append((url, title))

    additions = []

    for brand, urls in grouped.items():
        print("\nChecking:", brand)

        if not urls:
            print("  No NDTV candidate.")
            continue

        found_model = False

        for ndtv_url, ndtv_title in urls[:12]:
            if found_model:
                break

            html = get_html(ndtv_url)
            if not html:
                continue

            if not qualifying_ndtv(ndtv_title, html, brand, start, end):
                continue

            ndtv_dt = article_date(html)
            model_hint = model_from_ndtv(ndtv_title, brand)

            print("  NDTV:", ndtv_title)
            print("  Date:", ndtv_dt.isoformat() if ndtv_dt else "unknown")

            # Official launch confirmation + official product page.
            domains = BRANDS[brand]

            official_url, official_html = official_launch_confirmed(
                model_hint, brand, domains, ndtv_dt
            )

            if not official_url or not official_html:
                print("  SKIPPED: official India confirmation not found.")
                continue

            print("  Official:", official_url)

            name = product_name_from_official(official_html, model_hint)

            # Never use an NDTV article headline as the product name.
            if name.lower() == ndtv_title.lower():
                print("  SKIPPED: official page did not provide an exact product name.")
                continue

            description = official_description(official_html, name)

            if len(description) < 60:
                print("  SKIPPED: official product description not verified.")
                continue

            variants = extract_variants(official_html, name)

            if not variants:
                print("  SKIPPED: exact Color + Storage + RAM mapping not verified.")
                continue

            images = official_images(official_html)

            if not images:
                print("  SKIPPED: official product image not found.")
                continue

            for color, storage, ram in variants:
                key = (
                    name.strip().lower(),
                    brand.lower(),
                    color.strip().lower(),
                    storage.strip().lower(),
                    ram.strip().lower(),
                )

                if key in existing:
                    continue

                additions.append({
                    "name": name,
                    "brand": brand,
                    "short": description[:500],
                    "long": description[:2000],
                    "color": color,
                    "storage": storage,
                    "ram": ram,
                    "images": images,
                })

                existing.add(key)

            print("  Product:", name)
            print("  Exact variants:", len(variants))
            print("  Official images:", len(images))
            found_model = True

        if not found_model:
            print("  No qualifying verified launch.")

    if not additions:
        print("\nNo new verified smartphones to add.")
        print("AUTOMATION COMPLETED")
        return

    backup_excel()

    for p in additions:
        imgs = p["images"]

        ws.append([
            p["name"],
            p["brand"],
            "Electronics",
            "Accessories",
            "piece",
            "NO TAX APPLIED",
            "NODISCOUNT0",
            "1 Year Warranty",
            p["short"],
            p["long"],
            "true",
            "false",
            p["color"],
            p["storage"],
            p["ram"],
            imgs[0] if len(imgs) > 0 else "",
            "true",
            imgs[1] if len(imgs) > 1 else "",
            imgs[2] if len(imgs) > 2 else "",
            imgs[3] if len(imgs) > 3 else "",
        ])

    wb.save(XLSX)

    print("\nRows added:", len(additions))
    print("Rows skipped: verified duplicates / incomplete official data")
    print("AUTOMATION COMPLETED")


if __name__ == "__main__":
    main()
