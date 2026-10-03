from pathlib import Path
import re
import shutil
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin, urlparse, quote_plus

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook


# ============================================================
# RETAILIX SMARTPHONE AUTOMATION
#
# DISCOVERY:
#   1) 91mobiles - India latest/release date
#   2) Gadgets 360 - latest-phone/release-date cross-check
#
# PRODUCT DATA:
#   Official India brand website - RAM / storage / colour /
#   description
#
# IMAGE:
#   91mobiles - ONE primary product image
#
# IMPORTANT:
#   - Only selected brands are processed.
#   - Only phones with an India release date in the previous
#     24 hours are considered.
#   - Upcoming / rumours / reviews / global-only launches are
#     excluded.
#   - Variant rows are never generated as a Cartesian product.
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

BRAND_ALIASES = {
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

HEADERS = [
    "productName", "brandName", "categoryName", "subcategoryName",
    "unit_name", "taxName", "discountName", "warrantyName",
    "shortDescription", "longDescription", "is_featured", "is_digital",
    "variant_Color", "variant_Storage", "variant_RAM",
    "image1", "image1Primary", "image2", "image3", "image4"
]

SELECTED_BRANDS_LOWER = {
    brand: tuple(x.lower() for x in aliases)
    for brand, aliases in BRAND_ALIASES.items()
}

N91_LATEST_URL = "https://www.91mobiles.com/list-of-phones/latest-mobiles-in-india"
G360_LATEST_URL = "https://www.gadgets360.com/mobiles/latest-phones"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-IN,en;q=0.9",
})


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def now():
    return datetime.now(IST)


def normalize_url(url, base=""):
    if not url:
        return ""
    return urljoin(base, url).split("#")[0]


def same_domain(url, domains):
    host = urlparse(url).netloc.lower().split(":")[0]
    return any(host == d or host.endswith("." + d) for d in domains)


def get_html(url, timeout=30):
    try:
        r = SESSION.get(url, timeout=timeout, allow_redirects=True)
        if r.ok and len(r.text) > 300:
            return r.text
    except requests.RequestException:
        pass
    return ""


# ------------------------------------------------------------
# DATE PARSING
# ------------------------------------------------------------

MONTHS = (
    "January February March April May June July August "
    "September October November December"
)


def parse_date(text):
    text = clean(text)
    if not text:
        return None

    # ISO datetime/date
    iso_match = re.search(
        r"\b(20\d{2}-\d{1,2}-\d{1,2}"
        r"(?:[T ]\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?"
        r"(?:Z|[+-]\d{2}:?\d{2})?)?)\b",
        text,
    )
    if iso_match:
        value = iso_match.group(1)
        try:
            if "T" in value or " " in value:
                dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=IST)
                return dt.astimezone(IST)
            return datetime.strptime(value, "%Y-%m-%d").replace(tzinfo=IST)
        except ValueError:
            pass

    patterns = [
        (r"\b(\d{1,2})\s+([A-Za-z]+)\s+(20\d{2})\b", "%d %B %Y"),
        (r"\b([A-Za-z]+)\s+(\d{1,2}),\s*(20\d{2})\b", "%B %d, %Y"),
        (r"\b(\d{1,2})[/-](\d{1,2})[/-](20\d{2})\b", "%d/%m/%Y"),
        (r"\b(\d{1,2})\s+([A-Za-z]{3})\s+(20\d{2})\b", "%d %b %Y"),
    ]

    for pattern, fmt in patterns:
        m = re.search(pattern, text, re.I)
        if not m:
            continue
        try:
            return datetime.strptime(m.group(0), fmt).replace(tzinfo=IST)
        except ValueError:
            continue

    # Common site wording: "14 Sep, 2026"
    m = re.search(r"\b(\d{1,2})\s+([A-Za-z]{3,9}),\s*(20\d{2})\b", text, re.I)
    if m:
        try:
            return datetime.strptime(m.group(0), "%d %b, %Y").replace(tzinfo=IST)
        except ValueError:
            try:
                return datetime.strptime(m.group(0), "%d %B, %Y").replace(tzinfo=IST)
            except ValueError:
                pass

    return None


def date_only(dt):
    return dt.date() if dt else None


# ------------------------------------------------------------
# BRAND / MODEL HELPERS
# ------------------------------------------------------------

def detect_brand(text):
    low = clean(text).lower()
    for brand, aliases in SELECTED_BRANDS_LOWER.items():
        if any(re.search(r"\b" + re.escape(alias) + r"\b", low) for alias in aliases):
            return brand
    return None


def normalize_model_name(name, brand=None):
    name = clean(name)
    name = re.sub(
        r"\s+(?:full specs?|specifications?|price|review|photos?|"
        r"details?|features?)\s*$",
        "",
        name,
        flags=re.I,
    )
    name = re.sub(r"\s*\|\s*.*$", "", name)
    name = re.sub(r"\s+-\s+(?:price|specs?|review).*$", "", name, flags=re.I)
    name = re.sub(r"\s+", " ", name).strip()

    if brand == "REDMI":
        name = re.sub(r"^Xiaomi\s+", "", name, flags=re.I)
    elif brand == "POCO":
        name = re.sub(r"^Xiaomi\s+", "", name, flags=re.I)

    return name


def model_key(name):
    low = clean(name).lower()
    low = re.sub(r"[^a-z0-9]+", " ", low)
    low = re.sub(r"\b(5g|4g|india|phone|smartphone|mobile)\b", " ", low)
    return re.sub(r"\s+", " ", low).strip()


# ------------------------------------------------------------
# 91MOBILES DISCOVERY
# ------------------------------------------------------------

def parse_91mobiles_candidates():
    """
    Scrape the 91mobiles "Latest Mobiles in India" page.

    91mobiles exposes entries with:
      - product name
      - release date
      - product link

    We intentionally use the India page rather than an upcoming page.
    """
    html = get_html(N91_LATEST_URL)
    if not html:
        print("91mobiles: page unavailable.")
        return []

    soup = BeautifulSoup(html, "html.parser")
    candidates = {}

    # Primary method: locate links that point to phone detail pages.
    for a in soup.find_all("a", href=True):
        href = normalize_url(a["href"], N91_LATEST_URL)
        text = clean(a.get_text(" ", strip=True))

        if "91mobiles.com" not in urlparse(href).netloc.lower():
            continue
        if "/mobile/" not in href.lower():
            continue
        if len(text) < 2:
            continue

        # Inspect the closest useful container for release-date text.
        node = a
        container_text = ""
        for _ in range(5):
            node = node.parent
            if not node:
                break
            candidate_text = clean(node.get_text(" ", strip=True))
            if len(candidate_text) <= 3000:
                container_text = candidate_text
                if re.search(r"(?:new\s+)?release\s+date", candidate_text, re.I):
                    break

        release = None

        # Prefer "New Release Date" / "Release Date" immediately around the card.
        date_match = re.search(
            r"(?:new\s+)?release\s+date\s*[:\-]?\s*"
            r"(\d{1,2}\s+[A-Za-z]{3,9},?\s+20\d{2})",
            container_text,
            re.I,
        )
        if date_match:
            release = parse_date(date_match.group(1))

        if not release:
            date_match = re.search(
                r"(?:new\s+)?release\s+date\s*[:\-]?\s*"
                r"(\d{1,2}[/-]\d{1,2}[/-]20\d{2})",
                container_text,
                re.I,
            )
            if date_match:
                release = parse_date(date_match.group(1))

        if not release:
            continue

        brand = detect_brand(text)
        if not brand:
            brand = detect_brand(container_text)

        if not brand:
            continue

        name = normalize_model_name(text, brand)
        if not name or len(name) < 3:
            continue

        key = model_key(name)
        if not key:
            continue

        # Ignore obviously upcoming cards if the page marks them.
        low_container = container_text.lower()
        if "(upcoming)" in low_container or "upcoming" in low_container:
            continue

        candidates[key] = {
            "brand": brand,
            "name": name,
            "release": release,
            "url": href,
            "source": "91mobiles",
        }

    return list(candidates.values())


# ------------------------------------------------------------
# GADGETS 360 DISCOVERY
# ------------------------------------------------------------

def parse_gadgets360_candidates():
    """
    Scrape the Gadgets 360 latest-phone database.

    The page exposes product names and release dates.  We use it as
    the second independent discovery/cross-check source.
    """
    html = get_html(G360_LATEST_URL)
    if not html:
        print("Gadgets 360: page unavailable.")
        return []

    soup = BeautifulSoup(html, "html.parser")
    candidates = {}

    for a in soup.find_all("a", href=True):
        href = normalize_url(a["href"], G360_LATEST_URL)
        text = clean(a.get_text(" ", strip=True))

        if "gadgets360.com" not in urlparse(href).netloc.lower():
            continue
        if "/mobiles/" not in href.lower():
            continue
        if len(text) < 2:
            continue

        node = a
        container_text = ""
        for _ in range(6):
            node = node.parent
            if not node:
                break
            candidate_text = clean(node.get_text(" ", strip=True))
            if len(candidate_text) <= 3500:
                container_text = candidate_text
                if re.search(r"release\s+date", candidate_text, re.I):
                    break

        release = None
        m = re.search(
            r"release\s+date\s*[:\-]?\s*"
            r"(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+20\d{2})",
            container_text,
            re.I,
        )
        if m:
            release = parse_date(re.sub(r"(\d)(st|nd|rd|th)", r"\1", m.group(1)))

        if not release:
            m = re.search(
                r"release\s+date\s*[:\-]?\s*"
                r"(\d{1,2}\s+[A-Za-z]{3,9},?\s+20\d{2})",
                container_text,
                re.I,
            )
            if m:
                release = parse_date(m.group(1))

        if not release:
            continue

        brand = detect_brand(text) or detect_brand(container_text)
        if not brand:
            continue

        name = normalize_model_name(text, brand)
        if not name or len(name) < 3:
            continue

        key = model_key(name)
        if not key:
            continue

        low_container = container_text.lower()
        if "upcoming" in low_container:
            continue

        candidates[key] = {
            "brand": brand,
            "name": name,
            "release": release,
            "url": href,
            "source": "gadgets360",
        }

    return list(candidates.values())


def discover_launches(start, end):
    """
    A phone qualifies for discovery when:
      - it is in the selected brands;
      - 91mobiles has an India release date in the previous 24h;
      - Gadgets 360 also has the same model with a release date in the
        previous 24h.

    We use model/date matching rather than requiring either site to link
    to the other.
    """
    n91 = parse_91mobiles_candidates()
    g360 = parse_gadgets360_candidates()

    print("91mobiles candidates:", len(n91))
    print("Gadgets 360 candidates:", len(g360))

    g360_map = {}
    for item in g360:
        g360_map[(item["brand"], model_key(item["name"]))] = item

    qualified = []

    for item in n91:
        release = item["release"]
        if not (start <= release <= end):
            continue

        key = (item["brand"], model_key(item["name"]))
        other = g360_map.get(key)

        if not other:
            print(
                f"  CROSS-CHECK WAIT: {item['brand']} {item['name']} "
                f"({release.date()}) not present on Gadgets 360."
            )
            continue

        # Both sources must agree on the calendar release date.
        if date_only(other["release"]) != date_only(release):
            print(
                f"  DATE MISMATCH: {item['name']} | "
                f"91mobiles={release.date()} | "
                f"Gadgets360={other['release'].date()}"
            )
            continue

        item["g360_url"] = other["url"]
        item["g360_release"] = other["release"]
        item["discovery_date"] = release
        qualified.append(item)

        print(
            f"  QUALIFIED: {item['brand']} {item['name']} | "
            f"India release: {release.date()}"
        )

    return qualified


# ------------------------------------------------------------
# OFFICIAL INDIA BRAND SEARCH
# ------------------------------------------------------------

def official_page_is_india(url):
    low = url.lower()
    host = urlparse(url).netloc.lower()

    if "in." in host or host.endswith(".in"):
        return True

    return any(token in low for token in (
        "/in/", "/india/", "india.", "/india",
        "?country=in", "country=in", "locale=in",
    ))


def official_search_urls(model, domains):
    """
    Build search-engine URLs only to locate the official domain.

    The product data itself is read only from the resulting official
    India page. Search results are not accepted as product data.
    """
    q = quote_plus(
        f'site:{domains[0]} "{model}" smartphone India'
    )
    urls = [
        f"https://www.google.com/search?q={q}&num=10",
        f"https://www.bing.com/search?q={q}",
    ]

    # Add each official domain because some brands have multiple domains.
    for domain in domains[1:]:
        q2 = quote_plus(f'site:{domain} "{model}" India phone')
        urls.append(f"https://www.google.com/search?q={q2}&num=10")

    return urls


def extract_search_result_links(html, allowed_domains):
    soup = BeautifulSoup(html, "html.parser")
    results = []

    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        text = clean(a.get_text(" ", strip=True))

        # Google/Bing can wrap target URLs.
        if href.startswith("/url?q="):
            href = href.split("/url?q=", 1)[1].split("&", 1)[0]
        elif "uddg=" in href:
            href = href.split("uddg=", 1)[1].split("&", 1)[0]

        href = normalize_url(href)
        if not href:
            continue

        if same_domain(href, allowed_domains):
            results.append((href, text))

    # Preserve order, remove duplicates.
    seen = set()
    out = []
    for href, text in results:
        if href not in seen:
            seen.add(href)
            out.append((href, text))
    return out


def find_official_product_page(model, brand):
    domains = BRANDS[brand]
    candidates = []

    for search_url in official_search_urls(model, domains):
        html = get_html(search_url, timeout=20)
        if not html:
            continue

        for href, text in extract_search_result_links(html, domains):
            blob = (href + " " + text).lower()
            score = 0

            if official_page_is_india(href):
                score += 8
            if any(w in blob for w in ("phone", "smartphone", "mobile", "product")):
                score += 3

            for word in re.findall(r"[a-z0-9]+", model.lower()):
                if len(word) >= 3 and word in blob:
                    score += 1

            # Avoid support, manuals, careers, and generic landing pages.
            if any(x in href.lower() for x in (
                "/support", "/help", "/careers", "/about",
                "/privacy", "/terms", "/store-locator",
            )):
                score -= 10

            candidates.append((score, href))

        if candidates and max(x[0] for x in candidates) >= 8:
            break

    candidates.sort(reverse=True)

    for score, url in candidates:
        if score < 5:
            continue
        html = get_html(url, timeout=30)
        if not html:
            continue

        soup = BeautifulSoup(html, "html.parser")
        body = clean(soup.get_text(" ", strip=True))
        low = body.lower()

        model_words = [
            w.lower()
            for w in re.findall(r"[A-Za-z0-9]+", model)
            if len(w) >= 3
        ]

        if model_words and not all(w in low for w in model_words[:2]):
            continue

        # India confirmation. URL can be global-looking, but page must
        # contain India when the URL itself doesn't establish India.
        if not official_page_is_india(url) and "india" not in low:
            continue

        return url, html

    return "", ""


# ------------------------------------------------------------
# OFFICIAL PRODUCT DATA
# ------------------------------------------------------------

def product_name_from_official(html, fallback):
    soup = BeautifulSoup(html, "html.parser")

    tag = soup.find("h1")
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
            value = re.sub(
                r"\s*[-|]\s*(Official.*|Buy.*|India.*)$",
                "",
                value,
                flags=re.I,
            )
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
        if (
            model.lower() not in low
            and "smartphone" not in low
            and "phone" not in low
            and "mobile" not in low
        ):
            continue

        paragraphs.append(text)

    if paragraphs:
        return " ".join(paragraphs[:4])[:2000]

    return ""


def normalize_capacity(value):
    value = clean(value).upper()
    value = re.sub(r"\s+", " ", value)
    return value


def extract_variants(html, model):
    """
    Extract explicit RAM + storage + colour combinations.

    IMPORTANT:
    Never create a Cartesian product. A row is accepted only when
    RAM + storage + colour occur together in the same explicit block.
    """
    soup = BeautifulSoup(html, "html.parser")
    found = []

    ram_storage = re.compile(
        r"(?P<ram>\d{1,2})\s*GB\s*(?:RAM)?\s*"
        r"(?:\+|/|\||,|with)\s*"
        r"(?P<storage>\d{2,4})\s*(?P<unit>GB|TB)",
        re.I,
    )

    # Also accept common "8GB RAM, 256GB Storage" wording.
    ram_storage_2 = re.compile(
        r"(?P<ram>\d{1,2})\s*GB\s*RAM"
        r".{0,100}?"
        r"(?P<storage>\d{2,4})\s*(?P<unit>GB|TB)\s*(?:storage|memory)?",
        re.I,
    )

    color_re = re.compile(
        r"(?:colour|color|colourway|colorway)\s*[:\-]?\s*"
        r"([A-Za-z][A-Za-z0-9 &'’()./\-]{2,45})",
        re.I,
    )

    blocks = soup.find_all(
        ["li", "div", "section", "article", "label", "option", "tr", "td"]
    )

    for block in blocks:
        block_text = clean(block.get_text(" ", strip=True))
        if not block_text or len(block_text) > 700:
            continue

        pm = ram_storage.search(block_text) or ram_storage_2.search(block_text)
        cm = color_re.search(block_text)

        if pm and cm:
            ram = f"{pm.group('ram')} GB"
            storage = f"{pm.group('storage')} {pm.group('unit').upper()}"
            color = clean(cm.group(1))

            color = re.split(
                r"\s+(?:ram|storage|memory|price|₹|\|)\b",
                color,
                flags=re.I,
            )[0].strip()

            if (
                len(color) >= 2
                and len(color) <= 45
                and color.lower() not in {
                    "select", "choose", "available", "variant", "color"
                }
            ):
                item = (
                    color,
                    normalize_capacity(storage),
                    normalize_capacity(ram),
                )
                if item not in found:
                    found.append(item)

    # JSON/script blocks can contain explicit variant objects.
    for tag in soup.find_all("script"):
        raw = tag.string or tag.get_text()
        if not raw or len(raw) > 700000:
            continue

        for pm in ram_storage.finditer(raw):
            ram = f"{pm.group('ram')} GB"
            storage = f"{pm.group('storage')} {pm.group('unit').upper()}"

            nearby = (
                raw[max(0, pm.start() - 1500):pm.start()]
                + " "
                + raw[pm.end():pm.end() + 1500]
            )
            cm = color_re.search(nearby)
            if cm:
                color = clean(cm.group(1))
                color = re.split(
                    r"\s+(?:ram|storage|memory|price|₹|\|)\b",
                    color,
                    flags=re.I,
                )[0].strip()

                if len(color) >= 2 and len(color) <= 45:
                    item = (
                        color,
                        normalize_capacity(storage),
                        normalize_capacity(ram),
                    )
                    if item not in found:
                        found.append(item)

    # Deduplicate.
    output = []
    for item in found:
        if item not in output:
            output.append(item)

    return output


# ------------------------------------------------------------
# 91MOBILES PRIMARY IMAGE
# ------------------------------------------------------------

def score_91_image(url, alt="", width=0, height=0):
    if not url or not url.startswith("http"):
        return -999

    u = url.lower()
    a = alt.lower()

    bad = (
        "logo", "icon", "favicon", "sprite", "banner", "advert",
        "ad-", "offer", "amazon", "flipkart", "review", "screenshot",
        "avatar", "placeholder", "video", "gallery-thumb",
    )
    if any(x in u or x in a for x in bad):
        return -999

    score = 0

    if any(x in u for x in ("phone", "mobile", "product", "device")):
        score += 5

    if width >= 700 and height >= 700:
        score += 5
    elif width >= 400 and height >= 400:
        score += 3

    # Prefer product-like portrait images.
    if width and height and height >= width * 0.75:
        score += 2

    # Alt text matching the model is a strong signal.
    if any(word in a for word in re.findall(r"[a-z0-9]+", alt.lower()) if len(word) >= 3):
        score += 1

    return score


def extract_91_primary_image(page_url, model):
    html = get_html(page_url, timeout=30)
    if not html:
        return ""

    soup = BeautifulSoup(html, "html.parser")
    candidates = []

    # og:image is usually the strongest single primary-image candidate.
    for attrs in (
        {"property": "og:image"},
        {"name": "twitter:image"},
    ):
        m = soup.find("meta", attrs=attrs)
        if m and m.get("content"):
            u = normalize_url(m.get("content"), page_url)
            sc = score_91_image(u, model)
            if sc >= 0:
                candidates.append((sc + 4, u))

    for tag in soup.find_all(["img", "source"]):
        alt = tag.get("alt", "") or tag.get("title", "")

        urls = []
        for attr in (
            "src", "data-src", "data-original",
            "data-lazy-src", "data-image", "data-zoom-image",
        ):
            if tag.get(attr):
                urls.append(tag.get(attr))

        for part in (tag.get("srcset") or "").split(","):
            piece = part.strip().split(" ")[0]
            if piece:
                urls.append(piece)

        width = int(tag.get("width") or 0) if str(tag.get("width") or "").isdigit() else 0
        height = int(tag.get("height") or 0) if str(tag.get("height") or "").isdigit() else 0

        for raw in urls:
            u = normalize_url(raw, page_url)
            sc = score_91_image(u, alt, width, height)
            if sc >= 0:
                candidates.append((sc, u))

    if not candidates:
        return ""

    # Deduplicate while retaining the highest score.
    best = {}
    for score, url in candidates:
        best[url] = max(score, best.get(url, -999))

    ranked = sorted(
        ((score, url) for url, score in best.items()),
        reverse=True,
    )

    return ranked[0][1] if ranked else ""


# ------------------------------------------------------------
# EXCEL
# ------------------------------------------------------------

def backup_excel():
    stamp = now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(
        f"{XLSX.stem}_backup_{stamp}{XLSX.suffix}"
    )
    shutil.copy2(XLSX, backup)
    print("Backup:", backup.name)


def load_existing(ws):
    cols = {
        ws.cell(1, c).value: c
        for c in range(1, ws.max_column + 1)
    }

    required = {
        "productName", "brandName",
        "variant_Color", "variant_Storage", "variant_RAM"
    }
    if not required.issubset(cols):
        raise RuntimeError("Required Excel columns are missing.")

    existing = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        key = tuple(
            str(row[cols[k] - 1] or "").strip().lower()
            for k in (
                "productName",
                "brandName",
                "variant_Color",
                "variant_Storage",
                "variant_RAM",
            )
        )
        existing.add(key)

    return existing


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():
    print("=" * 78)
    print("RETAILIX DAILY SMARTPHONE AUTOMATION - FINAL")
    print("DISCOVERY: 91MOBILES + GADGETS 360")
    print("DATA: OFFICIAL INDIA BRAND WEBSITE")
    print("IMAGE: 91MOBILES - ONE PRIMARY IMAGE")
    print("=" * 78)

    if not XLSX.exists():
        raise FileNotFoundError(
            f"Excel file not found: {XLSX.resolve()}"
        )

    wb = load_workbook(XLSX)
    if "PRODUCT_DATA" not in wb.sheetnames:
        raise RuntimeError("Worksheet 'PRODUCT_DATA' was not found.")

    ws = wb["PRODUCT_DATA"]

    headers = [ws.cell(1, c).value for c in range(1, 21)]
    if headers != HEADERS:
        raise RuntimeError(
            "RetailiX template headers do not match the expected 20 columns."
        )

    existing = load_existing(ws)

    end = now()
    start = end - timedelta(hours=24)

    print(
        f"24-HOUR INDIA RELEASE WINDOW: "
        f"{start.isoformat()} -> {end.isoformat()}"
    )
    print("Existing combinations:", len(existing))

    # --------------------------------------------------------
    # 1. DISCOVERY
    # --------------------------------------------------------
    launches = discover_launches(start, end)

    if not launches:
        print("\nNo selected-brand phone was independently confirmed")
        print("by BOTH 91mobiles and Gadgets 360 in the last 24 hours.")
        print("AUTOMATION COMPLETED - NO EXCEL CHANGES.")
        return

    print(f"\nQualified launches: {len(launches)}")

    additions = []

    # --------------------------------------------------------
    # 2. OFFICIAL INDIA DATA
    # --------------------------------------------------------
    for item in launches:
        brand = item["brand"]
        model = item["name"]

        print("\n" + "-" * 70)
        print(f"PROCESSING: {brand} | {model}")
        print(f"India release date: {item['discovery_date'].date()}")
        print(f"91mobiles: {item['url']}")
        print(f"Gadgets 360: {item.get('g360_url', '')}")

        official_url, official_html = find_official_product_page(
            model, brand
        )

        if not official_url or not official_html:
            print("SKIPPED: official India product page not found.")
            continue

        print("Official India page:", official_url)

        name = product_name_from_official(official_html, model)
        description = official_description(official_html, name)
        variants = extract_variants(official_html, name)

        if len(description) < 60:
            print("SKIPPED: official product description not verified.")
            continue

        if not variants:
            print(
                "SKIPPED: official RAM + Storage + Colour combinations "
                "could not be verified."
            )
            continue

        # ----------------------------------------------------
        # 3. ONE PRIMARY IMAGE FROM 91MOBILES
        # ----------------------------------------------------
        image = extract_91_primary_image(item["url"], model)

        if not image:
            print("SKIPPED: primary 91mobiles product image not found.")
            continue

        print("Primary image:", image)
        print("Official variants:", len(variants))

        # ----------------------------------------------------
        # 4. DUPLICATE CHECK + ADDITIONS
        # ----------------------------------------------------
        added_for_phone = 0

        for color, storage, ram in variants:
            key = (
                name.strip().lower(),
                brand.lower(),
                color.strip().lower(),
                storage.strip().lower(),
                ram.strip().lower(),
            )

            if key in existing:
                print(
                    f"  DUPLICATE: {color} | {storage} | {ram}"
                )
                continue

            additions.append({
                "name": name,
                "brand": brand,
                "short": description[:500],
                "long": description[:2000],
                "color": color,
                "storage": storage,
                "ram": ram,
                "image": image,
            })

            existing.add(key)
            added_for_phone += 1

        print(f"New variants to add: {added_for_phone}")

    # --------------------------------------------------------
    # 5. WRITE EXCEL
    # --------------------------------------------------------
    if not additions:
        print("\nNo new variants to add.")
        print("AUTOMATION COMPLETED - NO EXCEL CHANGES.")
        return

    backup_excel()

    for p in additions:
        ws.append([
            p["name"],
            p["brand"],
            "Electronics",
            "Mobiles",
            "piece",
            "NO TAX APPLIED",
            "",
            "NO WARRANTY APPLIED",
            p["short"],
            p["long"],
            "true",
            "false",
            p["color"],
            p["storage"],
            p["ram"],
            p["image"],   # ONE primary image only
            "true",
            "",           # image2
            "",           # image3
            "",           # image4
        ])

    wb.save(XLSX)

    print("\n" + "=" * 78)
    print(f"ROWS ADDED: {len(additions)}")
    print(f"EXCEL SAVED: {XLSX}")
    print("image1 = 91mobiles primary product image")
    print("image1Primary = true")
    print("image2/image3/image4 = blank")
    print("AUTOMATION COMPLETED SUCCESSFULLY")
    print("=" * 78)


if __name__ == "__main__":
    main()
