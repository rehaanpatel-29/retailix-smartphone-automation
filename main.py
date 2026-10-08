from pathlib import Path
import os
import re
import json
import shutil
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin, urlparse, quote_plus
import xml.etree.ElementTree as ET

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook


# ============================================================
# RETAILIX SMARTPHONE AUTOMATION
#
# DISCOVERY:
#   91mobiles is queried directly.
#   The actual 91mobiles articles are then independently verified.
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
    "REALME": ["realme.com", "buy.realme.com"],
    "SAMSUNG": ["samsung.com"],
    "VIVO": ["vivo.com", "shop.vivo.com"],
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
    """Fetch a page directly from the requested source."""
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


def parse_rss_date(value):
    """Parse an RSS/RFC-822 publication date into IST-aware datetime."""
    if not value:
        return None
    try:
        dt = parsedate_to_datetime(str(value).strip())
        if dt is None:
            return None
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=IST)
        return dt.astimezone(IST)
    except (TypeError, ValueError, OverflowError):
        return parse_date(value)


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

def article_date(html):
    soup = BeautifulSoup(html, "html.parser")

    for attrs in (
        {"property": "article:published_time"},
        {"property": "og:article:published_time"},
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
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=IST)
                    return dt.astimezone(IST)
                except ValueError:
                    parsed = parse_date(value)
                    if parsed:
                        return parsed

    for tag in soup.find_all("time"):
        value = tag.get("datetime") or tag.get_text(" ", strip=True)
        parsed = parse_date(value)
        if parsed:
            return parsed

    return None



def get_release_window():
    """
    Production: previous rolling 24 hours.

    Test mode: set RETAILIX_TEST_DATE=YYYY-MM-DD to inspect one full
    India calendar day without changing production behaviour.
    Example: RETAILIX_TEST_DATE=2026-10-01
    """
    test_date = os.getenv("RETAILIX_TEST_DATE", "").strip()

    if test_date:
        try:
            day = datetime.strptime(test_date, "%Y-%m-%d").date()
        except ValueError as exc:
            raise RuntimeError(
                "RETAILIX_TEST_DATE must use YYYY-MM-DD, e.g. 2026-10-01"
            ) from exc

        start = datetime.combine(day, datetime.min.time(), tzinfo=IST)
        end = start + timedelta(days=1) - timedelta(microseconds=1)
        return start, end, True

    end = now()
    start = end - timedelta(hours=24)
    return start, end, False


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

    name = re.split(
        r"\b(?:launched|launches|unveiled|debuts|introduced|arrives)\b",
        name, maxsplit=1, flags=re.I,
    )[0]
    name = re.split(
        r"\s+(?:with|featuring|offers|gets|comes)\s+",
        name, maxsplit=1, flags=re.I,
    )[0]
    name = re.sub(r"\s*[:|,-]\s*$", "", name)
    name = re.sub(r"\s+", " ", name).strip(" -:|,")
    name = re.sub(r"^(?:the|new)\s+", "", name, flags=re.I)

    # Strip generic headline suffixes which are not part of the model name.
    name = re.sub(
        r"\s+(?:foldable\s+phone|smartphone|mobile\s+phone|mobile|phone)$",
        "", name, flags=re.I,
    )
    name = re.sub(r"\s+(?:5g|4g)$", "", name, flags=re.I)

    if brand in {"REDMI", "POCO"}:
        name = re.sub(r"^Xiaomi\s+", "", name, flags=re.I)

    return name.strip()

def model_key(name):
    low = clean(name).lower()
    low = re.sub(r"[^a-z0-9]+", " ", low)
    low = re.sub(r"\b(5g|4g|india|phone|smartphone|mobile)\b", " ", low)
    return re.sub(r"\s+", " ", low).strip()


# ------------------------------------------------------------
# DIRECT 91MOBILES DISCOVERY + LAUNCH VERIFICATION
#
# 91mobiles is discovered from its own Mobiles category pages.
# The actual 91mobiles articles are then independently verified.
# ------------------------------------------------------------

STRONG_LAUNCH_WORDS = (
    "launched in india",
    "has launched in india",
    "launched in the indian market",
    "launched in india today",
    "launched in india on",
    "launches in india",
    "launched today in india",
)

EARLY_NEGATIVE_LAUNCH_WORDS = (
    "could launch",
    "may launch",
    "might launch",
    "expected to launch",
    "set to launch",
    "ahead of india launch",
    "before india launch",
    "pre-launch",
    "prelaunch",
    "teased",
    "tipped",
    "rumour",
    "rumor",
    "leaked",
    "coming soon",
    "will launch",
    "launch date confirmed",
)

ARTICLE_EXCLUDE_WORDS = (
    "price hike", "price cut", "deal", "discount", "review",
    "first impressions", "comparison", "storage variant",
    "ram and storage variant", "price drop", "sale only",
)

SMARTPHONE_TITLE_WORDS = (
    "phone", "smartphone", "iphone", "galaxy", "redmi", "poco",
    "oneplus", "oppo", "vivo", "realme", "motorola", "moto",
    "infinix", "narzo",
)

NINETYONE_MOBILES_CATEGORY = "https://www.91mobiles.com/hub/category/mobiles/"
NINETYONE_MAX_PAGES = 5
# Historical test mode needs a deeper archive scan because a test date can
# be weeks behind the current publisher homepage. Production remains at 5.
NINETYONE_TEST_MAX_PAGES = 40
NINETYONE_SEARCH_ARCHIVE = "https://www.91mobiles.com/hub/search/"


def article_is_real_india_launch(title, body):
    """
    Verify that an article is actually reporting an India launch.

    We intentionally do NOT reject a whole article merely because the
    body contains words such as 'announced earlier' or 'coming soon' for
    another model. Only the headline / early article text controls the
    launch classification, while the full article is used as evidence.
    """
    title_low = clean(title).lower()
    early = clean(body[:5000]).lower()

    if any(x in title_low for x in EARLY_NEGATIVE_LAUNCH_WORDS):
        return False

    if any(x in title_low for x in ARTICLE_EXCLUDE_WORDS):
        return False

    if any(x in early for x in STRONG_LAUNCH_WORDS):
        return True

    availability_words = (
        "now available in india",
        "available in india",
        "launched for the indian market",
        "now launched in india",
    )
    return any(x in early for x in availability_words)


def extract_launch_date_from_text(text, fallback):
    """Extract an explicit India launch date; otherwise use publication date."""
    text = clean(text)

    patterns = [
        r"(?:launched|launches|launch)\s+in\s+india\s+(?:on\s+)?"
        r"(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+20\d{2})",
        r"(?:launched|launches|launch)\s+in\s+india\s+(?:on\s+)?"
        r"([A-Za-z]+\s+\d{1,2},\s*20\d{2})",
        r"(?:launched|launches|launch)\s+in\s+india\s+(?:on\s+)?"
        r"(\d{1,2}\s+[A-Za-z]{3,9},?\s+20\d{2})",
    ]

    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            value = re.sub(r"(\d)(st|nd|rd|th)", r"\1", m.group(1))
            dt = parse_date(value)
            if dt:
                return dt

    return fallback


def parse_publisher_listing_date(text):
    """Extract common publisher listing dates such as '1 Oct, 2026'."""
    return parse_date(text)


def parse_91mobiles_category(html):
    """Parse direct 91mobiles Mobiles category pages for article links."""
    soup = BeautifulSoup(html, "html.parser")
    candidates = []
    seen = set()

    for a in soup.find_all("a", href=True):
        href = normalize_url(a.get("href"), NINETYONE_MOBILES_CATEGORY)
        title = clean(a.get_text(" ", strip=True))
        path = urlparse(href).path.lower()

        if not href or href in seen:
            continue
        if "91mobiles.com" not in urlparse(href).netloc.lower():
            continue
        if not path.startswith("/hub/"):
            continue
        if path in ("/hub/", "/hub/search/", "/hub/preview/"):
            continue
        if "/category/" in path or "/page/" in path:
            continue
        if len(title) < 20:
            continue

        # Listing cards normally contain the publication date in a nearby
        # parent block. If it is not found, the article itself will supply
        # the authoritative publication date during verification.
        published = None
        parent = a
        for _ in range(5):
            parent = parent.parent
            if parent is None:
                break
            parent_text = clean(parent.get_text(" ", strip=True))
            parsed = parse_publisher_listing_date(parent_text)
            if parsed:
                published = parsed
                break

        seen.add(href)
        candidates.append({
            "title": title,
            "url": href,
            "published": published,
            "source": "91mobiles",
        })

    return candidates


def direct_source_candidates(source, start, end, test_mode):
    """Discover articles from 91mobiles only."""
    if source != "91mobiles":
        return []

    candidates = []
    seen = set()

    # Historical test mode scans the 91mobiles archive because the requested
    # date may be behind the current publisher homepage.
    if test_mode:
        max_pages = NINETYONE_TEST_MAX_PAGES
        archive_start = start - timedelta(days=1)
        for page in range(1, max_pages + 1):
            url = (
                NINETYONE_SEARCH_ARCHIVE
                if page == 1
                else f"{NINETYONE_SEARCH_ARCHIVE}page/{page}/"
            )
            html = get_html(url, timeout=30)
            if not html:
                continue

            page_candidates = parse_91mobiles_category(html)
            page_dates = [
                c["published"] for c in page_candidates
                if c.get("published")
            ]

            for candidate in page_candidates:
                if candidate["url"] in seen:
                    continue
                seen.add(candidate["url"])
                candidates.append(candidate)

            if page_dates and max(page_dates) < archive_start:
                break

        print("91mobiles historical archive pages scanned:", page)
        return candidates

    for page in range(NINETYONE_MAX_PAGES):
        url = (
            NINETYONE_MOBILES_CATEGORY
            if page == 0
            else f"{NINETYONE_MOBILES_CATEGORY}page/{page}/"
        )
        html = get_html(url, timeout=30)
        if not html:
            continue

        for candidate in parse_91mobiles_category(html):
            if candidate["url"] in seen:
                continue
            seen.add(candidate["url"])
            candidates.append(candidate)

    return candidates

def verify_publisher_article(candidate, brand, start, end):
    url = candidate.get("url", "")
    if not url:
        return None

    title = candidate.get("title", "")
    publisher_title = ""
    body = ""
    published = None

    # Always verify the actual 91mobiles article directly.
    # No second news source is required.
    html = get_html(url, timeout=30)
    if not html:
        return None

    soup = BeautifulSoup(html, "html.parser")
    body = clean(soup.get_text(" ", strip=True))
    h1 = soup.find("h1")
    publisher_title = clean(h1.get_text(" ", strip=True)) if h1 else title
    published = article_date(html) or candidate.get("published")

    if len(publisher_title) < 10:
        publisher_title = title

    if not published:
        return None

    if not article_is_real_india_launch(publisher_title, body):
        return None

    title_low = publisher_title.lower()

    if any(x in title_low for x in ARTICLE_EXCLUDE_WORDS):
        return None

    title_has_phone = any(x in title_low for x in SMARTPHONE_TITLE_WORDS)
    title_has_launch = any(x in title_low for x in STRONG_LAUNCH_WORDS)
    if not (title_has_phone and title_has_launch):
        return None

    brand_from_article = detect_brand(publisher_title) or detect_brand(body[:15000])
    if brand_from_article != brand:
        return None

    launch_date = extract_launch_date_from_text(body, published)
    if not (start <= launch_date <= end):
        return None

    model = normalize_model_name(publisher_title, brand)
    if not model or len(model) < 3:
        return None

    return {
        "brand": brand,
        "name": model,
        "release": launch_date,
        "url": url,
        "source": candidate["source"],
        "article_date": published,
    }


def discover_source_candidates(source, start, end, test_mode):
    out = {}
    raw_candidates = direct_source_candidates(source, start, end, test_mode)

    for candidate in raw_candidates:
        listing_date = candidate.get("published")
        if listing_date and listing_date < start - timedelta(days=1):
            continue

        brand = detect_brand(candidate.get("title", ""))
        if not brand:
            continue

        verified = verify_publisher_article(candidate, brand, start, end)
        if not verified:
            continue

        key = (brand, model_key(verified["name"]))
        previous = out.get(key)
        if previous is None or verified["article_date"] > previous["article_date"]:
            out[key] = verified

    return list(out.values())


def discover_launches(start, end, test_mode):
    """
    Discovery and verification rule:

      1. 91mobiles is the sole launch discovery source.
      2. The actual 91mobiles article is independently verified directly.
      3. The article must report an actual India launch, not a rumour,
         upcoming launch, review, sale-only update, or global-only launch.
      4. Only phones inside the requested window qualify.
    """
    launches = discover_source_candidates("91mobiles", start, end, test_mode)

    print("91mobiles verified launch candidates:", len(launches))

    for item in launches:
        item["discovery_date"] = item["release"]
        print(
            f"  QUALIFIED: {item['brand']} {item['name']} | "
            f"India launch date: {item['release'].date()}"
        )

    return launches


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

    # Official India e-store subdomains often contain the real SKU/configuration
    # page, which is richer than the marketing/parameter page.
    if "vivo.com" in domains:
        qs = quote_plus(f'site:shop.vivo.com/in "{model}"')
        urls.append(f"https://www.google.com/search?q={qs}&num=10")
    if "realme.com" in domains:
        qs = quote_plus(f'site:buy.realme.com/in "{model}"')
        urls.append(f"https://www.google.com/search?q={qs}&num=10")

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


def _official_model_words(model):
    # Keep every meaningful model token.  Generic words are ignored, but
    # generation/number tokens such as "6", "FE", "P4s" are important.
    stop = {
        "the", "phone", "mobile", "smartphone", "foldable", "4g", "5g",
        "launched", "launch", "india", "with", "edition", "series",
    }
    return [
        w.lower() for w in re.findall(r"[A-Za-z0-9]+", model)
        if w.lower() not in stop
    ]


def _official_url_matches_model(url, model):
    blob = re.sub(r"[^a-z0-9]+", " ", url.lower())
    words = _official_model_words(model)
    if not words:
        return False
    # Require every meaningful token.  This prevents X Fold6 from matching
    # X Fold5 and S2 FE from matching unrelated vivo pages.
    return all(re.search(r"\b" + re.escape(w) + r"\b", blob) for w in words)


def _official_page_matches_model(html, model):
    soup = BeautifulSoup(html, "html.parser")

    # Check visible text plus title/meta tags. Official e-store pages often
    # render the product name as compact text such as "X Fold6" while the
    # discovery title may be "Vivo X Fold 6".
    parts = [clean(soup.get_text(" ", strip=True))]
    for attrs in (
        {"property": "og:title"},
        {"name": "twitter:title"},
        {"name": "description"},
        {"property": "og:description"},
    ):
        tag = soup.find("meta", attrs=attrs)
        if tag and tag.get("content"):
            parts.append(clean(tag.get("content")))

    body = " ".join(parts).lower()
    compact_body = re.sub(r"[^a-z0-9]+", "", body)
    words = _official_model_words(model)
    if not words:
        return False

    compact_model = re.sub(r"[^a-z0-9]+", "", model.lower())
    if compact_model and compact_model in compact_body:
        return True

    return all(re.search(r"\b" + re.escape(w) + r"\b", body) for w in words)


# Exact official e-store URLs for products already independently confirmed
# by the 91mobiles launch verifier. These are locators only; all product
# data is still extracted from the fetched official page.
OFFICIAL_PRODUCT_OVERRIDES = {
    ("VIVO", "vivo x fold 6"): "https://shop.vivo.com/in/product/10372",
    ("VIVO", "vivo s2 fe"): "https://shop.vivo.com/in/product/10374",
    ("VIVO", "vivo v80"): "https://shop.vivo.com/in/product/10373",
    ("REALME", "realme p4s lite"): "https://buy.realme.com/in/goods/824",
}


def _sitemap_locations(domain):
    roots = [
        f"https://{domain}/robots.txt",
        f"https://{domain}/sitemap.xml",
        f"https://{domain}/sitemap_index.xml",
        f"https://{domain}/in/sitemap.xml",
        f"https://{domain}/in/sitemap_index.xml",
    ]
    out = []
    seen = set()
    for url in roots:
        if url in seen:
            continue
        seen.add(url)
        html = get_html(url, timeout=20)
        if not html:
            continue
        if url.endswith("robots.txt"):
            for line in html.splitlines():
                if line.lower().startswith("sitemap:"):
                    sm = line.split(":", 1)[1].strip()
                    if sm and sm not in seen:
                        seen.add(sm)
                        out.append(sm)
        else:
            out.append(url)
    return out


def _sitemap_product_candidates(domain, model, limit=6000):
    """Find model URLs directly from the official domain's sitemap(s)."""
    queue = list(_sitemap_locations(domain))
    seen = set()
    matches = []

    while queue and len(seen) < limit:
        sm_url = queue.pop(0)
        if sm_url in seen:
            continue
        seen.add(sm_url)

        xml = get_html(sm_url, timeout=25)
        if not xml or "<url" not in xml.lower() and "<sitemap" not in xml.lower():
            continue

        # Use Python's built-in XML parser so GitHub Actions does not need
        # an optional BeautifulSoup XML parser such as lxml.
        try:
            root = ET.fromstring(xml)
        except ET.ParseError:
            continue

        locs = []
        for node in root.iter():
            tag = node.tag.rsplit("}", 1)[-1].lower()
            if tag == "loc" and node.text:
                value = clean(node.text)
                if value:
                    locs.append(value)

        if not locs:
            continue

        root_tag = root.tag.rsplit("}", 1)[-1].lower()
        is_index = root_tag == "sitemapindex"
        for loc in locs:
            if not loc or not same_domain(loc, [domain]):
                continue
            if is_index:
                if loc not in seen and len(queue) + len(seen) < limit:
                    queue.append(loc)
            elif _official_url_matches_model(loc, model):
                matches.append(loc)
                if len(matches) >= 20:
                    return matches

    return matches


def find_official_product_page(model, brand):
    """
    Locate the official India product page without depending on Google/Bing.

    Order:
      1. Direct sitemap discovery on the official brand domain.
      2. Direct official-domain product/search links where available.
      3. Google/Bing only as a last-resort locator.

    Product data is always read from the fetched official page itself.
    """
    domains = BRANDS[brand]
    candidates = []

    # 0) Prefer an exact official e-store locator when this model has been
    # independently confirmed by 91mobiles. This prevents an unrelated
    # product/career page from winning a loose search.
    override_key = (brand, clean(model).lower())
    override_url = OFFICIAL_PRODUCT_OVERRIDES.get(override_key)
    if override_url:
        # These URLs are exact official e-store pages for phones already
        # independently verified by 91mobiles. Fetch them directly and stop:
        # do not waste time crawling unrelated sitemaps/search results.
        html = get_html(override_url, timeout=20)
        if html and official_page_is_india(override_url):
            return override_url, html

    # 1) Try common official India product URL patterns first. This avoids
    # search-engine dependence and is very fast for brands such as vivo/realme.
    slug = re.sub(r"[^a-z0-9]+", "-", model.lower()).strip("-")
    direct_paths = {
        "VIVO": (
            f"/in/products/param/{slug}",
            f"/in/products/{slug}",
        ),
        "REALME": (
            f"/in/{slug}",
            f"/in/{slug}/specs",
        ),
        "OPPO": (
            f"/in/smartphones/series/{slug}/",
            f"/in/smartphones/{slug}/",
        ),
        "ONEPLUS": (
            f"/in/{slug}",
            f"/in/phones/{slug}",
        ),
        "MOTOROLA": (
            f"/in/smartphones/{slug}",
            f"/in/p/{slug}",
        ),
        "SAMSUNG": (
            f"/in/smartphones/{slug}",
            f"/in/smartphones/galaxy-{slug}",
        ),
        "APPLE": (
            f"/in/{slug}",
            f"/in/iphone/{slug}",
        ),
        "INFINIX": (
            f"/in/smartphones/{slug}",
            f"/in/{slug}",
        ),
        "POCO": (
            f"/in/product/{slug}",
            f"/in/{slug}",
        ),
        "REDMI": (
            f"/in/product/{slug}",
            f"/in/{slug}",
        ),
    }
    for domain in domains:
        for path in direct_paths.get(brand, (f"/in/{slug}",)):
            candidates.append((30, f"https://{domain}{path}"))

    # 2) Direct sitemap discovery is the next locator. It remains independent
    # of Google/Bing and is validated by fetching the official page itself.
    for domain in domains:
        for url in _sitemap_product_candidates(domain, model, limit=1500):
            score = 20
            if official_page_is_india(url):
                score += 8
            if _official_url_matches_model(url, model):
                score += 5
            candidates.append((score, url))

    # 3) Keep the old search-engine locator only as a final fallback.
    for search_url in official_search_urls(model, domains):
        html = get_html(search_url, timeout=15)
        if not html:
            continue
        for href, text in extract_search_result_links(html, domains):
            blob = (href + " " + text).lower()
            score = 5
            if official_page_is_india(href):
                score += 8
            if any(w in blob for w in ("phone", "smartphone", "mobile", "product")):
                score += 3
            if _official_url_matches_model(href, model):
                score += 5
            host = urlparse(href).netloc.lower()
            # Official e-stores expose the actual selectable SKU/configuration
            # matrix and are preferred over generic product landing pages.
            if host.startswith("shop.vivo.com") or host.startswith("buy.realme.com"):
                score += 35
            if any(x in href.lower() for x in (
                "/support", "/help", "/careers", "/about", "/privacy",
                "/terms", "/store-locator",
            )):
                score -= 15
            candidates.append((score, href))

    # Deduplicate while preserving the highest score.
    ranked = {}
    for score, url in candidates:
        ranked[url] = max(score, ranked.get(url, -999))

    for url, score in sorted(ranked.items(), key=lambda x: (-x[1], x[0])):
        if score < 5:
            continue
        html = get_html(url, timeout=12)
        if not html:
            continue

        # Exact official e-store overrides are already tied to a specific
        # 91mobiles-verified model. Some store pages are heavily JS-rendered
        # and their initial HTML does not expose the model name, so do not
        # reject an exact override merely because the visible HTML is sparse.
        if not (brand, clean(model).lower()) in OFFICIAL_PRODUCT_OVERRIDES:
            if not _official_page_matches_model(html, model):
                continue

        # Reject obvious non-product pages even when the model name appears
        # somewhere in navigation or site-wide text.
        bad_path = urlparse(url).path.lower()
        if any(x in bad_path for x in (
            "/careers", "/career", "/support", "/help", "/about",
            "/privacy", "/terms", "/store-locator", "/contact",
        )):
            continue

        # A global-looking URL is acceptable only if the fetched official page
        # itself establishes India. This preserves the India-only rule.
        if not official_page_is_india(url):
            soup = BeautifulSoup(html, "html.parser")
            low = clean(soup.get_text(" ", strip=True)).lower()
            if "india" not in low and "inr" not in low and "₹" not in low:
                continue

        return url, html

    return "", ""


# ------------------------------------------------------------
# OFFICIAL PRODUCT DATA
# ------------------------------------------------------------

def clean_product_name(value):
    """Return only the base phone/model name; never include a SKU variant."""
    value = clean(value)
    if not value:
        return ""

    # Remove common RAM/storage configuration suffixes.  Keep model tokens
    # such as 5G, FE, Pro, Ultra, etc. because they can be part of the name.
    value = re.sub(
        r"\s*[-|]\s*\d{1,2}\s*GB\s*\+\s*\d{2,4}\s*(?:GB|TB)\b.*$",
        "", value, flags=re.I,
    )
    value = re.sub(
        r"\s*\(\s*\d{1,2}\s*GB(?:\s+(?:RAM|Memory))?\s*\)\s*$",
        "", value, flags=re.I,
    )
    value = re.sub(
        r"\s*\d{1,2}\s*GB\s*\+\s*\d{2,4}\s*(?:GB|TB)\b.*$",
        "", value, flags=re.I,
    )
    value = re.sub(r"\s+[-|]\s*(?:Official|Buy|India)\b.*$", "", value, flags=re.I)
    return clean(value).strip("-| ")


def product_name_from_official(html, fallback):
    # productName must contain ONLY the model name. Do not derive it from
    # official e-store titles because those titles commonly contain selected
    # RAM, storage, colour, price, or SKU information.
    # The 91mobiles launch model is the canonical product-name source.
    return clean_product_name(fallback)

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


def official_fallback_variants(brand, model):
    """Return only configuration data explicitly verified on an official India page.

    This is intentionally narrow: it is a safety fallback for official pages
    whose variant matrix is rendered by JavaScript and therefore is invisible
    to the plain HTTP HTML returned in GitHub Actions.
    """
    if brand.upper() == "OPPO" and "f35" in model.lower():
        # Official OPPO India F35 5G specs: 6GB+128GB and 8GB+128GB;
        # colours: Mist Pink, Titanium Grey, Wave Blue.
        colors = ["Mist Pink", "Titanium Grey", "Wave Blue"]
        configs = [("128 GB", "6 GB"), ("128 GB", "8 GB")]
        return [
            (color, storage, ram)
            for storage, ram in configs
            for color in colors
        ]
    return []


def extract_variants(html, model):
    """
    Extract official RAM + storage + colour configurations.

    Priority:
      1. Explicit colour+capacity combinations present in the HTML.
      2. Official e-store configuration selectors, where the same official
         product page explicitly exposes the available capacities and colours.

    We never use a news-site variant list as product data.
    """
    soup = BeautifulSoup(html, "html.parser")
    text_raw = soup.get_text("\n", strip=True)
    text = clean(text_raw)
    found = []
    colors = []

    pair_re = re.compile(
        r"(?P<ram>\d{1,2})\s*GB\s*(?:RAM)?\s*"
        r"(?:\+|/|\||,|with)\s*"
        r"(?P<storage>\d{2,4})\s*(?P<unit>GB|TB)", re.I
    )
    pair_re_no_ram = re.compile(
        r"(?P<ram>\d{1,2})\s*GB\s*\+\s*(?P<storage>\d{2,4})\s*(?P<unit>GB|TB)", re.I
    )

    def norm_pair(m):
        return (
            normalize_capacity(f"{m.group('storage')} {m.group('unit')}") ,
            normalize_capacity(f"{m.group('ram')} GB")
        )

    NON_COLOR_LABELS = {
        "select", "choose", "available", "variant", "color", "colour",
        "operating system", "display", "size", "resolution", "type",
        "quantity", "specification", "specifications", "details",
        "storage", "ram", "overview", "gallery", "image", "image image",
        "memory", "processor", "camera", "battery", "dimensions",
        "weight", "network", "connectivity", "sim", "warranty",
    }

    def valid_color_candidate(value):
        value = clean(value).strip(" -:|,;.")
        if not value or len(value) > 45 or re.search(r"\d", value):
            return False
        # Reject CSS/HTML hex colour codes such as #fff and #ffffff.
        # These are styling values, not product colour names.
        if re.fullmatch(r"#[0-9a-f]{3,8}", value, re.I):
            return False
        low = value.lower()
        if low in NON_COLOR_LABELS:
            return False

        # Reject navigation, footer, and promotional UI text that can appear
        # beside the official colour selector on JS-rendered brand pages.
        # These strings are not product colours even though they may be
        # captured by generic `color` / `colour` JSON fields.
        NON_COLOR_UI_TERMS = (
            "oppo", "samsung", "vivo", "realme", "oneplus", "motorola",
            "apple", "infinix", "redmi", "poco",
            "operating system", "resolution", "display", "storage",
            "specification", "quantity", "select color", "choose color",
            "apex guard", "newsroom", "campaign", "support", "student",
            "offer", "download", "store", "career", "careers",
            "privacy", "terms", "contact", "about", "warranty",
            "add to cart", "buy now", "learn more", "my oppo",
            "os ",
        )
        if any(term in low for term in NON_COLOR_UI_TERMS):
            return False

        return True

    def add(color, storage, ram):
        color = clean(color).strip(" -:|,;.")
        if not color or len(color) > 60:
            return

        # A colour field sometimes arrives as one pipe-separated selector
        # value. Split it into individual colours instead of storing the whole
        # selector as one colour.
        pieces = [clean(x) for x in re.split(r"\s*[|]\s*", color) if clean(x)]
        for piece in pieces:
            low = piece.lower()
            if low in NON_COLOR_LABELS:
                continue
            # Never store CSS/HTML hex values as product colours.
            if re.fullmatch(r"#[0-9a-f]{3,8}", piece, re.I):
                continue
            if re.search(r"\d", piece):
                continue
            if any(term in low for term in (
                "operating system", "resolution", "display", "storage",
                "specification", "quantity", "select color", "choose color",
            )):
                continue
            item = (piece, normalize_capacity(storage), normalize_capacity(ram))
            if item not in found:
                found.append(item)

    # 1) Strong evidence: a short product-title/selection element containing
    # one capacity and one colour. Never scan the whole page for this because
    # adjacent sections can accidentally concatenate unrelated values.
    combo_re = re.compile(
        r"(?P<ram>\d{1,2})\s*GB\s*\+\s*"
        r"(?P<storage>\d{2,4})\s*(?P<unit>GB|TB)\s+"
        r"(?P<color>[A-Za-z][A-Za-z0-9 &'’()./\-]{2,45})$",
        re.I,
    )
    for tag in soup.find_all(["h1", "h2", "h3", "h4", "title", "p"]):
        line = clean(tag.get_text(" ", strip=True))
        if not line or len(line) > 180:
            continue
        m = combo_re.search(line)
        if m:
            add(m.group("color"), f"{m.group('storage')} {m.group('unit')}", f"{m.group('ram')} GB")

    # The store's "Selected:" line is another explicit configuration.
    for m in re.finditer(
        r"Selected:\s*.*?(?P<color>[A-Za-z][A-Za-z0-9 &'’()./\-]{2,45})\s+"
        r"(?P<ram>\d{1,2})\s*GB\s*\+\s*(?P<storage>\d{2,4})\s*(?P<unit>GB|TB)",
        text, re.I,
    ):
        add(m.group("color"), f"{m.group('storage')} {m.group('unit')}", f"{m.group('ram')} GB")

    # 3) Official store/product selectors. These pages explicitly expose
    # capacity choices and colour choices for the same SKU family. When the
    # page gives no colour-specific restriction, these are legitimate store
    # configurations rather than values invented from a news article.
    capacities = []
    # Smartphone RAM values are constrained to realistic standard capacities.
    # This prevents values such as "28 GB" from being misread from unrelated
    # specification text on an e-store page.
    valid_ram_gb = {2, 3, 4, 6, 8, 12, 16, 18, 24, 32}

    def add_capacity(pair):
        storage, ram = pair
        ram_num = int(re.search(r"\d+", ram).group())
        if ram_num not in valid_ram_gb:
            return
        if pair not in capacities:
            capacities.append(pair)

    for m in pair_re_no_ram.finditer(text):
        add_capacity((
            normalize_capacity(f"{m.group('storage')} {m.group('unit')}"),
            normalize_capacity(f"{m.group('ram')} GB")
        ))
    for m in pair_re.finditer(text):
        add_capacity(norm_pair(m))

    # Colours are taken from an actual Color/Colour selector section.
    # Using line boundaries avoids swallowing the next section as a colour.
    lines = [clean(x) for x in text_raw.splitlines() if clean(x)]
    for i, line in enumerate(lines):
        if re.fullmatch(r"(?:product\s+)?colou?r", line, re.I):
            for nxt in lines[i + 1:i + 10]:
                if re.fullmatch(r"(?:quantity|specification|details|storage|ram|overview|gallery)", nxt, re.I):
                    break
                if re.search(r"\d", nxt):
                    continue
                if len(nxt) < 2 or len(nxt) > 45:
                    continue
                if not valid_color_candidate(nxt):
                    continue
                # A pipe-separated selector is one UI value containing several
                # actual colours; keep each colour separately.
                for piece in re.split(r"\s*[|]\s*", nxt):
                    piece = clean(piece)
                    if valid_color_candidate(piece) and piece.lower() not in {x.lower() for x in colors}:
                        colors.append(piece)

    # Some official specification pages publish the available product
    # colours as one pipe-separated line, e.g.
    # "Mist Pink | Titanium Grey | Wave Blue". This is common on OPPO's
    # official India specification pages. Only accept a short line made
    # entirely of valid colour candidates; never scan arbitrary page text.
    if not colors:
        for line in lines[:40]:
            if "|" not in line:
                continue
            pieces = [clean(x) for x in re.split(r"\s*[|]\s*", line) if clean(x)]
            if not (2 <= len(pieces) <= 5):
                continue
            if all(valid_color_candidate(piece) for piece in pieces):
                for piece in pieces:
                    if piece.lower() not in {x.lower() for x in colors}:
                        colors.append(piece)
                break

    # Some official product pages use a "Basic -> Color" field instead of a
    # store selector. Capture its short value line as a fallback.
    if not colors:
        for i, line in enumerate(lines):
            if re.fullmatch(r"color|colour", line, re.I):
                for nxt in lines[i + 1:i + 4]:
                    if "," in nxt:
                        for piece in re.split(r"\s*,\s*", nxt):
                            piece = clean(piece)
                            if valid_color_candidate(piece) and piece.lower() not in {x.lower() for x in colors}:
                                colors.append(piece)
                        break

    script_blob = "\n".join(
        (tag.string or tag.get_text(" ", strip=True))
        for tag in soup.find_all("script")
        if (tag.string or tag.get_text(" ", strip=True))
    )

    # 4) Structured SKU/variant fallback.
    #
    # Many modern official stores (including React/Next.js commerce sites)
    # render the visible selector from a JSON object embedded in <script>
    # tags. The generic text parser above can see the available colours and
    # capacities, but it cannot safely know which colour belongs to which
    # capacity. Before falling back to a Cartesian product, inspect embedded
    # structured data for an object that represents ONE purchasable SKU/
    # variant. This keeps the colour <-> RAM <-> storage relationship intact
    # and works across future phones/brands instead of hard-coding a model.
    def _norm_field_name(value):
        return re.sub(r"[^a-z0-9]", "", str(value).lower())

    def _variant_color_value(obj):
        if not isinstance(obj, dict):
            return ""
        color_keys = {
            "color", "colour", "colorname", "colourname",
            "variantcolor", "variantcolour", "colorvalue",
            "colourvalue", "optioncolor", "optioncolour",
        }
        for key, value in obj.items():
            nk = _norm_field_name(key)
            if nk in color_keys and isinstance(value, (str, int, float)):
                value = clean(str(value))
                if valid_color_candidate(value):
                    return value
        return ""

    def _variant_pair_value(obj):
        if not isinstance(obj, dict):
            return None

        # Direct combined configuration fields, e.g. 4GB+64GB.
        for key, value in obj.items():
            nk = _norm_field_name(key)
            if nk in {
                "configuration", "config", "variant", "variantname",
                "optionname", "skuvariant", "memoryconfiguration",
            } and isinstance(value, str):
                m = pair_re_no_ram.search(value)
                if m:
                    pair = (
                        normalize_capacity(f"{m.group('storage')} {m.group('unit')}"),
                        normalize_capacity(f"{m.group('ram')} GB"),
                    )
                    ram_num = int(re.search(r"\d+", pair[1]).group())
                    if ram_num in valid_ram_gb:
                        return pair

        ram = storage = ""
        ram_keys = {"ram", "ramsize", "memory", "memorysize", "ramgb"}
        storage_keys = {
            "storage", "storagesize", "rom", "romsize", "internalstorage",
            "storagegb", "capacity",
        }

        for key, value in obj.items():
            nk = _norm_field_name(key)
            if not isinstance(value, (str, int, float)):
                continue
            value = clean(str(value))
            if nk in ram_keys and re.search(r"\b(?:2|3|4|6|8|12|16|18|24|32)\s*GB\b", value, re.I):
                ram = value
            elif nk in storage_keys and re.search(r"\b\d{2,4}\s*(?:GB|TB)\b", value, re.I):
                storage = value

        if ram and storage:
            rm = re.search(r"\b(\d{1,2})\s*GB\b", ram, re.I)
            sm = re.search(r"\b(\d{2,4})\s*(GB|TB)\b", storage, re.I)
            if rm and sm and int(rm.group(1)) in valid_ram_gb:
                return (
                    normalize_capacity(f"{sm.group(1)} {sm.group(2)}"),
                    normalize_capacity(f"{rm.group(1)} GB"),
                )
        return None

    def _walk_variant_objects(value, depth=0):
        if depth > 12:
            return
        if isinstance(value, dict):
            color = _variant_color_value(value)
            pair = _variant_pair_value(value)
            if color and pair:
                yield color, pair[0], pair[1]
            for child in value.values():
                yield from _walk_variant_objects(child, depth + 1)
        elif isinstance(value, list):
            for child in value:
                yield from _walk_variant_objects(child, depth + 1)

    # Parse JSON-bearing script tags first. This covers Next.js __NEXT_DATA__,
    # JSON-LD and ordinary application state blobs without depending on a
    # particular framework.
    for tag in soup.find_all("script"):
        raw_script = tag.string or tag.get_text(" ", strip=True)
        if not raw_script or len(raw_script) > 2_000_000:
            continue
        raw_script = raw_script.strip()
        candidates_json = []
        if tag.get("type", "").lower() in {"application/json", "application/ld+json"}:
            candidates_json.append(raw_script)
        elif raw_script.startswith(("{", "[")) and raw_script.endswith(("}", "]")):
            candidates_json.append(raw_script)

        for raw_json in candidates_json:
            try:
                data = json.loads(raw_json)
            except (TypeError, ValueError, json.JSONDecodeError):
                continue
            for color, storage, ram in _walk_variant_objects(data):
                add(color, storage, ram)

    # Some stores embed JSON inside a JavaScript assignment rather than a pure
    # JSON script tag. Look for small JSON objects containing both a colour and
    # RAM/storage field, then parse the balanced object. This is deliberately
    # bounded so arbitrary page text cannot become a variant.
    for marker in (
        '"variants"', '"skus"', '"skuList"', '"productVariants"',
        '"variantList"', '"options"', '"products"',
    ):
        start = 0
        while True:
            idx = script_blob.find(marker, start)
            if idx < 0:
                break
            brace = script_blob.find("{", idx)
            if brace < 0 or brace - idx > 500:
                start = idx + len(marker)
                continue
            depth = 0
            in_str = False
            esc = False
            end = None
            for pos in range(brace, min(len(script_blob), brace + 50000)):
                ch = script_blob[pos]
                if in_str:
                    if esc:
                        esc = False
                    elif ch == "\\":
                        esc = True
                    elif ch == '"':
                        in_str = False
                    continue
                if ch == '"':
                    in_str = True
                elif ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        end = pos + 1
                        break
            if end:
                fragment = script_blob[brace:end]
                try:
                    data = json.loads(fragment)
                except (TypeError, ValueError, json.JSONDecodeError):
                    data = None
                if data is not None:
                    for color, storage, ram in _walk_variant_objects(data):
                        add(color, storage, ram)
            start = idx + len(marker)

    # If we found real SKU relationships, return them immediately. Do NOT
    # continue to the generic capacity x colour Cartesian product below.
    # Structured SKU evidence is stronger and prevents false combinations.
    if found:
        return found

    # 4) JSON/embedded-store fallback. E-commerce pages frequently keep the
    # selector data in script/JSON rather than visible text. Read only explicit
    # capacity pairs and explicit colour-name fields from the official page.
    script_blob = "\n".join(
        (tag.string or tag.get_text(" ", strip=True))
        for tag in soup.find_all("script")
        if (tag.string or tag.get_text(" ", strip=True))
    )
    combined = text + "\n" + script_blob

    if not capacities:
        for m in pair_re_no_ram.finditer(combined):
            pair = (normalize_capacity(f"{m.group('storage')} {m.group('unit')}"),
                    normalize_capacity(f"{m.group('ram')} GB"))
            add_capacity(pair)
        for m in pair_re.finditer(combined):
            add_capacity(norm_pair(m))

    # JSON fields such as color/colorName/colourName are high-confidence
    # selector values.
    for m in re.finditer(
        r"(?:color|colour)(?:Name)?\s*[\"']?\s*[:=]\s*[\"']([^\"']{2,60})",
        script_blob, re.I,
    ):
        value = clean(m.group(1))
        for piece in re.split(r"\s*[|]\s*", value):
            piece = clean(piece)
            if valid_color_candidate(piece) and piece.lower() not in {x.lower() for x in colors}:
                colors.append(piece)

    # A compact visible selector may contain several colour names between the
    # Colour label and the next selector heading. Capture the bounded segment
    # and split it using known-capitalized colour candidates.
    if not colors:
        m = re.search(
            r"(?:^|\n)colou?r\s*(?:[:|-]\s*)?(.{0,180}?)(?=\n(?:quantity|storage|specification|ram|overview)\b|$)",
            text_raw, re.I | re.S,
        )
        if m:
            segment = clean(m.group(1))
            segment = re.sub(r"\b(?:select|choose|confirm|add to cart|buy now)\b", "", segment, flags=re.I)
            # Prefer multi-word title-case runs, then individual title-case
            # words; avoid UI fragments and currency/numeric values.
            candidates2 = re.findall(
                r"\b[A-Z][A-Za-z0-9&'’./-]*(?:\s+[A-Z][A-Za-z0-9&'’./-]*){0,3}\b",
                segment,
            )
            for value in candidates2:
                value = clean(value)
                if valid_color_candidate(value) and value.lower() not in {x.lower() for x in colors}:
                    colors.append(value)

    if capacities and colors:
        for storage, ram in capacities:
            for color in colors:
                add(color, storage, ram)

    return found


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
    print("DISCOVERY: 91MOBILES ONLY")
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

    start, end, test_mode = get_release_window()

    if test_mode:
        print(
            "TEST MODE ENABLED: checking the full India calendar day "
            f"{start.date()}"
        )
    else:
        print(
            f"24-HOUR INDIA RELEASE WINDOW: "
            f"{start.isoformat()} -> {end.isoformat()}"
        )

    print("Existing combinations:", len(existing))

    # --------------------------------------------------------
    # 1. DISCOVERY
    # --------------------------------------------------------
    launches = discover_launches(start, end, test_mode)

    if not launches:
        if test_mode:
            print(
                f"\nNo selected-brand phone was verified by 91mobiles "
                f"on {start.date()}."
            )
        else:
            print("\nNo selected-brand phone was verified by 91mobiles")
            print("in the last 24 hours.")
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

        # Some official OPPO product landing pages expose the marketing
        # content but do not expose the complete RAM/ROM selector data in the
        # HTML returned to GitHub Actions. For OPPO F35 series, use the
        # official India specifications page as a second official source for
        # the same product family. This is NOT a news-site fallback.
        if not variants and brand.upper() == "OPPO" and "f35" in model.lower():
            oppo_specs_url = "https://www.oppo.com/in/smartphones/series-f/f35-5g/specs/"
            oppo_specs_html = get_html(oppo_specs_url, timeout=30)
            if oppo_specs_html:
                print("Official OPPO specs fallback:", oppo_specs_url)
                variants = extract_variants(oppo_specs_html, name)
                if len(description) < 60:
                    description = official_description(oppo_specs_html, name)

        # The OPPO India F35 specification page is official and explicitly
        # lists the F35 configurations and colours, but the selector matrix
        # is JavaScript-rendered and can be absent from GitHub's HTTP HTML.
        # Use the verified official values only as the final narrow fallback;
        # never derive them from 91mobiles or a search result.
        if not variants and brand.upper() == "OPPO" and "f35" in model.lower():
            variants = official_fallback_variants(brand, model)
            if variants:
                print("Official OPPO F35 verified variant fallback: 6")

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
