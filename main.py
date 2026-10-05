from pathlib import Path
import os
import re
import shutil
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin, urlparse, quote_plus, parse_qs

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook


# ============================================================
# RETAILIX SMARTPHONE AUTOMATION
#
# DISCOVERY:
#   91mobiles + Gadgets 360 are queried directly.
#   The actual publisher articles are then independently verified.
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

    # News headlines often put the useful model name first and then add
    # chipset/specification/price text. Keep the model portion only.
    name = re.split(
        r"\b(?:launched|launches|unveiled|debuts|introduced|arrives)\b",
        name,
        maxsplit=1,
        flags=re.I,
    )[0]
    name = re.split(
        r"\s+(?:with|featuring|offers|gets|comes)\s+",
        name,
        maxsplit=1,
        flags=re.I,
    )[0]
    name = re.split(
        r"\s+(?:price in india|price,?\s*availability|availability|specifications|features|review|first impressions)\b",
        name,
        maxsplit=1,
        flags=re.I,
    )[0]
    name = re.sub(r"\s*[:|,-]\s*$", "", name)
    name = re.sub(r"\s+", " ", name).strip(" -:|,")
    name = re.sub(r"^(?:the|new)\s+", "", name, flags=re.I)

    if brand in {"REDMI", "POCO"}:
        name = re.sub(r"^Xiaomi\s+", "", name, flags=re.I)

    return name


def model_key(name):
    low = clean(name).lower()
    low = re.sub(r"[^a-z0-9]+", " ", low)
    low = re.sub(r"\b(5g|4g|india|phone|smartphone|mobile)\b", " ", low)
    return re.sub(r"\s+", " ", low).strip()


# ------------------------------------------------------------
# DIRECT 91MOBILES + GADGETS 360 DISCOVERY + TWO-SOURCE LAUNCH VERIFICATION
#
# 91mobiles is discovered from its own Mobiles category pages.
# Gadgets 360 is discovered from its own Mobiles RSS feed.
# The actual publisher articles are then independently verified.
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
    "goes global",
    "going global",
    "global launch",
    "global debut",
    "launched globally",
    "launches globally",
    "global market",
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

GADGETS360_MOBILES_RSS = "https://www.gadgets360.com/rss/mobiles/feeds"
GADGETS360_MOBILES_NEWS = "https://www.gadgets360.com/mobiles/news"
GADGETS360_TURBO_LIST = "https://turbo.gadgets360.com/list"
GADGETS360_HINDI_NEWS = "https://hindi.gadgets360.com/mobiles/news"
GADGETS360_OFFICIAL_HOSTS = {
    "www.gadgets360.com",
    "gadgets360.com",
    "turbo.gadgets360.com",
    "hindi.gadgets360.com",
}
JINA_READER_BASE = "https://r.jina.ai/"
JINA_SEARCH_BASE = "https://s.jina.ai/"
SEARCH_ENGINES = (
    "https://www.google.com/search?q={query}&num=10&hl=en-IN",
    "https://www.bing.com/search?q={query}&count=10",
    "https://html.duckduckgo.com/html/?q={query}",
)
NINETYONE_MOBILES_CATEGORY = "https://www.91mobiles.com/hub/category/mobiles/"
NINETYONE_MAX_PAGES = 5


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

    # Allow natural article wording such as:
    # "Samsung has launched the Galaxy A08 in India".
    if re.search(r"\b(?:has|have|was|were|is|are)?\s*launched\b.{0,120}\bin\s+india\b", early, re.I):
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


def extract_published_date_from_text(text):
    """Extract a publisher Published/Updated date from Jina's text output."""
    text = clean(text)
    patterns = [
        r"(?:published|updated)(?:\s+on|\s*:)?\s+(\d{1,2}\s+[A-Za-z]{3,9},?\s+20\d{2})",
        r"(?:published|updated)(?:\s+on|\s*:)?\s+([A-Za-z]+\s+\d{1,2},\s*20\d{2})",
        r"(?:published|updated)(?:\s+on|\s*:)?\s+(\d{1,2}[/-]\d{1,2}[/-]20\d{2})",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            dt = parse_date(m.group(1))
            if dt:
                return dt
    return None


def parse_gadgets360_rss(xml):
    """Parse the official Gadgets 360 Mobiles RSS feed."""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []

    candidates = []
    for item in root.findall(".//item"):
        title = clean(item.findtext("title"))
        link = normalize_url(clean(item.findtext("link")))
        pub_text = clean(item.findtext("pubDate"))
        published = parse_rss_date(pub_text)

        if not title or not link or not published:
            continue

        if "gadgets360.com" not in urlparse(link).netloc.lower():
            continue
        if not urlparse(link).path.lower().startswith("/mobiles/news/"):
            continue

        candidates.append({
            "title": title,
            "url": link,
            "published": published,
            "source": "gadgets360",
        })

    return candidates


def jina_read_url(url, timeout=45):
    """Read a URL through Jina Reader when the GitHub runner cannot fetch it directly."""
    if not url:
        return None

    reader_url = JINA_READER_BASE + url
    try:
        response = SESSION.get(
            reader_url,
            timeout=timeout,
            headers={
                "Accept": "application/json",
                "X-Retain-Links": "true",
                "X-With-Links-Summary": "true",
            },
        )
        if not response.ok or len(response.text) < 100:
            return None

        try:
            payload = response.json()
        except ValueError:
            return {
                "url": url,
                "title": "",
                "content": response.text,
            }

        data = payload.get("data", payload) if isinstance(payload, dict) else payload
        if not isinstance(data, dict):
            return None

        content = clean(data.get("content", ""))
        title = clean(data.get("title", ""))
        final_url = clean(data.get("url", "")) or url

        if not content:
            return None

        return {
            "url": final_url,
            "title": title,
            "content": content,
        }
    except requests.RequestException:
        return None


def jina_search(query, timeout=45):
    """Search the web through Jina and return normalized result dictionaries."""
    if not query:
        return []

    search_url = JINA_SEARCH_BASE + quote_plus(query)
    try:
        response = SESSION.get(
            search_url,
            timeout=timeout,
            headers={
                "Accept": "application/json",
            },
        )
        if not response.ok or len(response.text) < 50:
            return []

        results = []
        try:
            payload = response.json()
        except ValueError:
            # Jina can return readable Markdown. Extract the source URLs
            # directly from Markdown links as a fallback.
            for match in re.finditer(
                r"\[[^\]]+\]\((https?://[^)]+)\)", response.text
            ):
                results.append({
                    "title": "",
                    "url": normalize_url(match.group(1)),
                    "published": None,
                })
            return results

        data = payload.get("data", payload) if isinstance(payload, dict) else payload
        if isinstance(data, dict):
            data = data.get("results", data.get("items", []))
        if not isinstance(data, list):
            return []

        for item in data:
            if not isinstance(item, dict):
                continue
            url = normalize_url(item.get("url", ""))
            if not url:
                continue
            results.append({
                "title": clean(item.get("title", "")),
                "url": url,
                "published": parse_date(item.get("published", "")),
            })

        return results
    except requests.RequestException:
        return []


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


def get_gadgets360_news_links():
    """Get article links from the official G360 Mobiles News page."""
    html = get_html(GADGETS360_MOBILES_NEWS, timeout=30)
    if not html:
        return []

    soup = BeautifulSoup(html, "html.parser")
    candidates = []
    seen = set()

    for a in soup.find_all("a", href=True):
        href = normalize_url(a.get("href"), GADGETS360_MOBILES_NEWS)
        title = clean(a.get_text(" ", strip=True))
        if not href or href in seen:
            continue

        parsed = urlparse(href)
        if "gadgets360.com" not in parsed.netloc.lower():
            continue
        if not parsed.path.lower().startswith("/mobiles/news/"):
            continue
        if len(title) < 20:
            continue

        seen.add(href)
        candidates.append({
            "title": title,
            "url": href,
            "published": None,
            "source": "gadgets360",
        })

    return candidates


def _is_gadgets360_article_url(url):
    """Return True only for official Gadgets 360 mobile-news article URLs."""
    if not url:
        return False
    parsed = urlparse(url)
    host = parsed.netloc.lower().split(":")[0]
    path = parsed.path.lower()
    if host not in GADGETS360_OFFICIAL_HOSTS:
        return False
    return (
        (host in {"www.gadgets360.com", "gadgets360.com"} and path.startswith("/mobiles/news/"))
        or (host == "turbo.gadgets360.com" and "/mobiles/" in path)
        or (host == "hindi.gadgets360.com" and path.startswith("/mobiles/"))
    )


def _official_g360_candidate(title, url, published=None, snippet=""):
    url = normalize_url(url)
    if not _is_gadgets360_article_url(url):
        return None
    return {
        "title": clean(title),
        "url": url,
        "published": published,
        "snippet": clean(snippet),
        "source": "gadgets360",
    }


def get_gadgets360_official_mirror_links():
    """Discover G360 articles from official Turbo/Hindi listing pages."""
    candidates = []
    seen = set()
    for listing_url in (GADGETS360_TURBO_LIST, GADGETS360_HINDI_NEWS):
        html = get_html(listing_url, timeout=30)
        if not html:
            continue
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = normalize_url(a.get("href"), listing_url)
            title = clean(a.get_text(" ", strip=True))
            if not href or href in seen or len(title) < 12:
                continue
            if not _is_gadgets360_article_url(href):
                continue
            seen.add(href)
            candidates.append(_official_g360_candidate(title, href))
    return [x for x in candidates if x]


def search_engine_gadgets360_candidates(model, brand):
    """Use public search engines only to locate official G360 article URLs.

    Search engines are discovery/transport only. The resulting URL is still
    verified against the official G360 article by verify_publisher_article().
    """
    queries = [
        f'site:gadgets360.com/mobiles/news "{model}"',
        f'site:gadgets360.com/mobiles/news "{brand}" "{model}"',
        f'site:gadgets360.com/mobiles/news "{model}" "launched in India"',
    ]
    found = []
    seen = set()

    for query_text in queries:
        encoded = quote_plus(query_text)
        for template in SEARCH_ENGINES:
            search_url = template.format(query=encoded)
            html = get_html(search_url, timeout=20)
            if not html:
                continue

            soup = BeautifulSoup(html, "html.parser")
            anchors = soup.find_all("a", href=True)
            for a in anchors:
                href = a.get("href", "")
                text = clean(a.get_text(" ", strip=True))
                parent_text = clean(a.parent.get_text(" ", strip=True)) if a.parent else text

                # Google / Bing / DDG wrappers.
                if href.startswith("/url?q="):
                    href = href.split("/url?q=", 1)[1].split("&", 1)[0]
                elif href.startswith("https://www.google.com/url?") or href.startswith("http://www.google.com/url?"):
                    href = parse_qs(urlparse(href).query).get("q", [""])[0]
                elif "uddg=" in href:
                    href = parse_qs(urlparse(href).query).get("uddg", [href.split("uddg=", 1)[1].split("&", 1)[0]])[0]

                href = normalize_url(href)
                if not _is_gadgets360_article_url(href) or href in seen:
                    continue

                # Search-result title/snippet is only a locator. It is not
                # accepted as final launch evidence.
                listing_date = parse_date(parent_text)
                seen.add(href)
                found.append(_official_g360_candidate(text, href, listing_date, parent_text))

            # Some search engines expose the result URLs only in raw HTML.
            for match in re.finditer(
                r'https?://(?:www\.|turbo\.|hindi\.)?gadgets360\.com/[^\s"<>]+',
                html,
                re.I,
            ):
                href = normalize_url(match.group(0).rstrip(".,;')\""))
                if not _is_gadgets360_article_url(href) or href in seen:
                    continue
                seen.add(href)
                found.append(_official_g360_candidate("", href))

    return [x for x in found if x]


def gadgets360_search_fallback_candidates(n91_candidates):
    """Locate official G360 articles when G360 blocks the GitHub runner."""
    found = []
    seen = set()
    for item in n91_candidates:
        model = item.get("name", "").strip()
        brand = item.get("brand", "").strip()
        if not model or not brand:
            continue
        for candidate in search_engine_gadgets360_candidates(model, brand):
            if candidate["url"] in seen:
                continue
            seen.add(candidate["url"])
            found.append(candidate)

    print("Gadgets 360 search-engine article candidates:", len(found))
    return found


def direct_source_candidates(source, start, end, test_mode):
    """Discover articles from the two approved publisher sources."""
    if source == "gadgets360":
        # RSS is useful when accessible, but GitHub Actions can receive 403.
        xml = get_html(GADGETS360_MOBILES_RSS, timeout=30)
        if xml:
            rss_candidates = parse_gadgets360_rss(xml)
            if rss_candidates:
                return rss_candidates
        else:
            print("Gadgets 360 RSS unavailable; using Mobiles News page fallback.")

        page_candidates = get_gadgets360_news_links()
        if page_candidates:
            print("Gadgets 360 Mobiles News page candidates:", len(page_candidates))
            return page_candidates

        mirror_candidates = get_gadgets360_official_mirror_links()
        if mirror_candidates:
            print("Gadgets 360 official mirror candidates:", len(mirror_candidates))
            return mirror_candidates

        # Last direct-source fallback: read the official G360 Mobiles News
        # page through Jina and extract only G360 /mobiles/news/ links.
        jina = jina_read_url(GADGETS360_MOBILES_NEWS)
        if jina:
            text = jina.get("content", "")
            candidates = []
            seen = set()

            # Jina normally preserves publisher links as Markdown. Capture
            # both the article title and the exact G360 URL so the normal
            # brand/title pre-filter can work before article verification.
            for match in re.finditer(
                r"\[([^\]]{10,200})\]\((https?://(?:www\.)?gadgets360\.com/mobiles/news/[^)]+)\)",
                text,
                re.I,
            ):
                title = clean(match.group(1))
                href = normalize_url(match.group(2).rstrip(".,;"))
                if not href or href in seen:
                    continue
                if len(title) < 20:
                    continue
                seen.add(href)
                candidates.append({
                    "title": title,
                    "url": href,
                    "published": None,
                    "source": "gadgets360",
                })

            # If the response contains bare URLs instead of Markdown links,
            # retain them as a secondary locator.
            if not candidates:
                for match in re.finditer(
                    r"https?://(?:www\.)?gadgets360\.com/mobiles/news/[^\s)\]]+",
                    text,
                    re.I,
                ):
                    href = normalize_url(match.group(0).rstrip(".,;"))
                    if not href or href in seen:
                        continue
                    seen.add(href)
                    candidates.append({
                        "title": "",
                        "url": href,
                        "published": None,
                        "source": "gadgets360",
                    })

            if candidates:
                print("Gadgets 360 Mobiles News page via Jina:", len(candidates))
                return candidates

        print("Gadgets 360 Mobiles News page unavailable.")
        return []

    candidates = []
    seen = set()

    for page in range(NINETYONE_MAX_PAGES):
        if page == 0:
            url = NINETYONE_MOBILES_CATEGORY
        else:
            url = f"{NINETYONE_MOBILES_CATEGORY}page/{page}/"

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

    # G360 is frequently blocked from GitHub-hosted runners. Try the real
    # publisher URL first; if that fails, use Jina Reader only as a transport
    # layer to fetch the SAME G360 article. The source remains G360.
    html = get_html(url, timeout=30)
    if html:
        soup = BeautifulSoup(html, "html.parser")
        body = clean(soup.get_text(" ", strip=True))
        h1 = soup.find("h1")
        publisher_title = clean(h1.get_text(" ", strip=True)) if h1 else title
        published = article_date(html) or candidate.get("published")
    elif candidate.get("source") == "gadgets360":
        jina = jina_read_url(url)
        if jina:
            body = clean(jina.get("content", ""))
            publisher_title = clean(jina.get("title", "")) or title
            published = (
                parse_date(jina.get("published", ""))
                or candidate.get("published")
                or extract_published_date_from_text(body[:12000])
            )
            print("  G360 article fetched via Jina:", url)
        else:
            return None
    else:
        return None

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
    if not title_has_phone:
        return None

    # Do not require the exact phrase "launched in India" in the headline.
    # G360 headlines can use forms such as "Price in India, Launch..." while
    # the article body explicitly confirms that the phone was launched in India.
    early_low = clean(body[:5000]).lower()
    if not (
        any(x in early_low for x in STRONG_LAUNCH_WORDS)
        or any(x in early_low for x in (
            "has launched the",
            "has launched",
            "was launched in india",
            "launched in india",
            "launched for the indian market",
            "now available in india",
            "available in india",
        ))
    ):
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

    # Use the publisher listing/feed date as a cheap pre-filter when present.
    # If it is missing, still verify the article because its own metadata is
    # authoritative and may contain the date.
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


def gadgets360_fallback_candidates(n91_candidates, start, end):
    """Locate G360 articles through Jina Search and public search engines.

    Every returned URL must still be an official Gadgets 360 article and is
    independently fetched/verified before it can qualify.
    """
    found = []
    seen = set()

    # 1) Public search engines are the primary recovery route when GitHub
    # cannot fetch G360 directly.
    for candidate in gadgets360_search_fallback_candidates(n91_candidates):
        if candidate["url"] in seen:
            continue
        seen.add(candidate["url"])
        found.append(candidate)

    # 2) Jina Search remains a secondary locator.
    for item in n91_candidates:
        model = item.get("name", "").strip()
        brand = item.get("brand", "").strip()
        if not model or not brand:
            continue

        queries = [
            f'site:gadgets360.com/mobiles/news "{model}" "launched in India"',
            f'site:gadgets360.com/mobiles/news "{brand}" "{model}"',
        ]
        for query_text in queries:
            for result in jina_search(query_text):
                href = normalize_url(result.get("url", ""))
                if not _is_gadgets360_article_url(href) or href in seen:
                    continue
                seen.add(href)
                found.append({
                    "title": clean(result.get("title", "")),
                    "url": href,
                    "published": result.get("published"),
                    "snippet": clean(result.get("snippet", "")),
                    "source": "gadgets360",
                })

    print("Gadgets 360 total fallback article candidates:", len(found))
    return found


def discover_launches(start, end, test_mode):
    """
    Discovery and verification rule:

      1. 91mobiles and Gadgets 360 are queried directly.
      2. The actual publisher articles are independently verified.
      3. Both publishers must report the same selected-brand phone as
         launched in India.
      4. Their India launch dates must agree.
      5. Only phones inside the requested window qualify.
    """
    n91 = discover_source_candidates("91mobiles", start, end, test_mode)
    g360 = discover_source_candidates("gadgets360", start, end, test_mode)

    # GitHub Actions can receive HTTP 403 from the Gadgets 360 RSS endpoint.
    # In that case, locate G360 articles through a search engine and then run
    # the SAME publisher-page verification. This does not weaken the
    # two-source rule: the final evidence still comes from Gadgets 360.
    if not g360 and n91:
        print("Gadgets 360 direct discovery returned 0; starting all recovery methods.")
        fallback = gadgets360_fallback_candidates(n91, start, end)
        fallback_verified = {}

        for candidate in fallback:
            brand = detect_brand(candidate.get("title", ""))
            if not brand:
                # Search snippets can omit the brand; the article verifier
                # checks the publisher page itself, so use the 91mobiles
                # model's brand when the model was the search target.
                candidate_key = model_key(candidate.get("title", ""))
                for item in n91:
                    target_key = model_key(item["name"])
                    if target_key and (target_key in candidate_key or candidate_key in target_key):
                        brand = item["brand"]
                        break
            if not brand:
                continue

            verified = verify_publisher_article(
                candidate, brand, start, end
            )
            if not verified:
                continue

            key = (brand, model_key(verified["name"]))
            previous = fallback_verified.get(key)
            if previous is None or verified["article_date"] > previous["article_date"]:
                fallback_verified[key] = verified

        g360 = list(fallback_verified.values())

    print("91mobiles verified launch candidates:", len(n91))
    print("Gadgets 360 verified launch candidates:", len(g360))

    g360_map = {
        (item["brand"], model_key(item["name"])): item
        for item in g360
    }

    qualified = []

    for item in n91:
        key = (item["brand"], model_key(item["name"]))
        other = g360_map.get(key)

        if not other:
            print(
                f"  CROSS-CHECK WAIT: {item['brand']} {item['name']} "
                f"({item['release'].date()}) not confirmed by Gadgets 360."
            )
            continue

        if date_only(other["release"]) != date_only(item["release"]):
            print(
                f"  DATE MISMATCH: {item['name']} | "
                f"91mobiles={item['release'].date()} | "
                f"Gadgets360={other['release'].date()}"
            )
            continue

        item["g360_url"] = other["url"]
        item["g360_release"] = other["release"]
        item["discovery_date"] = item["release"]
        qualified.append(item)

        print(
            f"  QUALIFIED: {item['brand']} {item['name']} | "
            f"India launch date: {item['release'].date()}"
        )

    return qualified


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
    print("DISCOVERY: 91MOBILES + GADGETS 360 DIRECT")
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
                f"\nNo selected-brand phone was independently confirmed "
                f"by BOTH 91mobiles and Gadgets 360 on {start.date()}."
            )
        else:
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
