from pathlib import Path
import re
import urllib.parse
import time
import shutil
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
import feedparser

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


BRAND_ALIASES = {
    "APPLE": ["Apple", "iPhone"],
    "INFINIX": ["Infinix"],
    "MOTOROLA": ["Motorola", "Moto"],
    "ONEPLUS": ["OnePlus"],
    "OPPO": ["OPPO"],
    "POCO": ["POCO"],
    "REDMI": ["Redmi", "Xiaomi"],
    "REALME": ["realme"],
    "SAMSUNG": ["Samsung", "Galaxy"],
    "VIVO": ["vivo"],
}


BRAND_HOMES = {
    "APPLE": "https://www.apple.com/in/",
    "INFINIX": "https://infinixmobiles.in/",
    "MOTOROLA": "https://www.motorola.in/",
    "ONEPLUS": "https://www.oneplus.in/",
    "OPPO": "https://www.oppo.com/in/",
    "POCO": "https://www.mi.com/in/",
    "REDMI": "https://www.mi.com/in/",
    "REALME": "https://www.realme.com/in/",
    "SAMSUNG": "https://www.samsung.com/in/",
    "VIVO": "https://www.vivo.com/in/",
}

NDTV_URLS = [
    "https://www.ndtv.com/topic/smartphones",
    "https://www.ndtv.com/topic/mobile-phones",
    "https://www.ndtv.com/topic/smartphone",
    "https://www.ndtv.com/topic/phones",
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
    "speculation", "hands-on", "launch date", "set for", "to launch",
    "will launch", "launching on", "coming soon", "ahead of launch"
)

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 Chrome/140 Safari/537.36"
})

GOOGLE_DATES = {}


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


def _collect_ndtv_links(html, base):
    """Extract article links from an NDTV topic page."""
    out = {}
    soup = BeautifulSoup(html, "html.parser")

    for a in soup.find_all("a", href=True):
        url = normalize_url(a["href"], base)
        title = clean(a.get_text(" ", strip=True))

        if not url.startswith("https://www.ndtv.com/"):
            continue
        if len(title) < 15:
            continue

        path = urlparse(url).path.lower()
        if any(path.startswith(x) for x in (
            "/topic/", "/video/", "/photos/", "/photo/",
            "/webstories/", "/web-story/", "/live-updates/"
        )):
            continue

        blob = title.lower()
        if not any(x in blob for x in SMART_WORDS):
            continue

        # Keep discovery broad. Launch/review/leak filtering happens after
        # opening the actual NDTV article.
        if any(x in blob for x in (
            "review", " vs ", "versus", "comparison", "price drop",
            "discount", "deal", "offers", "best phones", "top phones",
            "how to", "tips", "wallpaper", "case", "cover"
        )):
            continue

        out[url] = title

    return out




def resolve_google_news_url(url):
    """Turn a Google News RSS article URL into the original publisher URL."""
    if not url or "news.google.com" not in url:
        return url

    try:
        r = requests.get(
            url,
            headers={"User-Agent": SESSION.headers.get("User-Agent", "Mozilla/5.0")},
            timeout=20,
            allow_redirects=True,
        )
        final_url = r.url
        if final_url and "news.google.com" not in final_url:
            return final_url
    except Exception:
        pass

    return url


def google_news_candidates():
    """Fast discovery: Google News RSS only. No publisher article fetching."""
    out = {}
    search_terms = [
        '"launched" smartphone India',
        '"launched in India" smartphone',
        '"India launch" smartphone',
        '"launched" phone India',
    ]

    for brand in BRANDS:
        brand_terms = [brand] + BRAND_ALIASES.get(brand, [])
        brand_query = "(" + " OR ".join(brand_terms) + ")"
        for term in search_terms:
            query = f"{brand_query} {term} when:1d"
            url = (
                "https://news.google.com/rss/search?"
                + urllib.parse.urlencode({
                    "q": query, "hl": "en-IN", "gl": "IN", "ceid": "IN:en"
                })
            )
            try:
                feed = feedparser.parse(url)
                for entry in feed.entries[:20]:
                    title = clean(getattr(entry, "title", ""))
                    link = getattr(entry, "link", "")
                    if not title or not link:
                        continue
                    low = title.lower()
                    if not any(a.lower() in low for a in brand_terms):
                        continue
                    if not any(w in low for w in SMART_WORDS):
                        continue
                    if not any(w in low for w in LAUNCH_WORDS):
                        continue

                    dt = None
                    q = getattr(entry, "published_parsed", None) or getattr(entry, "updated_parsed", None)
                    if q:
                        try:
                            dt = datetime(*q[:6], tzinfo=timezone.utc).astimezone(IST)
                        except Exception:
                            pass

                    # Keep the RSS URL only as an identifier. We do not open
                    # news articles; launch/data proof comes from official sites.
                    out[link] = title
                    if dt:
                        GOOGLE_DATES[link] = dt
            except Exception:
                continue

    return list(out.items())

def ndtv_candidates():
    """
    NDTV is the only discovery layer.

    First try normal HTTP. If NDTV serves the topic page in a form that does
    not expose article links to requests (which can happen with bot/CDN
    handling), fall back to Playwright Chromium. This is still the same NDTV
    source; no Google, Amazon, Flipkart, or other news site is used.
    """
    out = {}

    for base in NDTV_URLS:
        html = get_html(base)
        if html:
            out.update(_collect_ndtv_links(html, base))

    if out:
        print(f"NDTV HTTP discovery: {len(out)} article candidates")
        return list(out.items())

    print("NDTV HTTP discovery returned 0 candidates; trying browser discovery...")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(
                user_agent=(
                    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                    "Chrome/140.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1365, "height": 900},
            )

            for base in NDTV_URLS:
                try:
                    page.goto(base, wait_until="domcontentloaded", timeout=30000)
                    page.wait_for_timeout(2500)
                    html = page.content()
                    out.update(_collect_ndtv_links(html, base))
                except Exception as e:
                    print(f"  Browser discovery failed for {base}: {e}")

            browser.close()
    except Exception as e:
        print(f"NDTV browser discovery unavailable: {e}")

    print(f"NDTV browser discovery: {len(out)} article candidates")
    return list(out.items())

def candidate_brand(title):
    """Assign the brand from the earliest strong brand mention in the headline."""
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
    hits = []
    for brand, words in aliases.items():
        for word in words:
            m = re.search(r"(?<![a-z0-9])" + re.escape(word.lower()) + r"(?![a-z0-9])", low)
            if m:
                hits.append((m.start(), brand))
    if not hits:
        return None
    hits.sort(key=lambda x: (x[0], x[1]))
    return hits[0][1]


def qualifying_ndtv(title, html, brand, start, end):
    soup = BeautifulSoup(html, "html.parser")
    body = clean(soup.get_text(" ", strip=True))
    title_low = title.lower()
    body_low = body[:12000].lower()

    # Reject obvious non-launch stories based primarily on the headline.
    if any(x in title_low for x in BAD_WORDS):
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

    if not any(x in (title_low + " " + body_low) for x in aliases[brand]):
        return False

    # Launch language may appear in the article body rather than the title.
    if not any(x in (title_low + " " + body_low) for x in LAUNCH_WORDS):
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



def official_direct_product_urls(model, brand):
    """Return conservative official India URL candidates for known site patterns."""
    slug = re.sub(r"[^a-z0-9]+", "-", model.lower()).strip("-")
    compact = re.sub(r"[^a-z0-9]+", "", model.lower())

    patterns = {
        "ONEPLUS": [
            f"https://www.oneplus.in/{slug}",
            f"https://www.oneplus.in/{compact}",
        ],
        "REALME": [
            f"https://www.realme.com/in/realme-{slug}",
            f"https://www.realme.com/in/{slug}",
        ],
        "MOTOROLA": [
            f"https://www.motorola.in/smartphones-{slug}",
            f"https://www.motorola.in/{slug}",
        ],
        "SAMSUNG": [
            f"https://www.samsung.com/in/smartphones/{slug}/",
        ],
        "OPPO": [
            f"https://www.oppo.com/in/smartphones/series-{slug}/",
            f"https://www.oppo.com/in/smartphones/{slug}/",
        ],
        "VIVO": [
            f"https://www.vivo.com/in/products/{slug}",
        ],
        "INFINIX": [
            f"https://infinixmobility.com/{slug}",
        ],
        "APPLE": [
            f"https://www.apple.com/in/{slug}/",
        ],
        "POCO": [
            f"https://www.mi.com/in/product/{slug}/",
        ],
        "REDMI": [
            f"https://www.mi.com/in/product/{slug}/",
        ],
    }
    return patterns.get(brand, [])


def official_product_url_from_home(model, brand, domains):
    """Find a model page by crawling only the official India brand site."""
    home = BRAND_HOMES.get(brand, "")
    if not home:
        return ""

    words = [
        w.lower() for w in re.findall(r"[A-Za-z0-9]+", model)
        if len(w) >= 2
    ]
    if not words:
        return ""

    def score_link(url, text=""):
        if not same_domain(url, domains):
            return -1
        blob = (url + " " + text).lower()
        score = 0
        for w in words:
            if w in blob:
                score += 3
        if any(x in blob for x in ("phone", "smartphone", "mobile", "product")):
            score += 2
        if official_page_is_india(url):
            score += 2
        return score

    # Fast HTTP crawl first.
    html = get_html(home)
    if html:
        soup = BeautifulSoup(html, "html.parser")
        ranked = []
        for a in soup.find_all("a", href=True):
            u = normalize_url(a["href"], home)
            score = score_link(u, clean(a.get_text(" ", strip=True)))
            if score >= 5:
                ranked.append((score, u))
        if ranked:
            ranked.sort(reverse=True)
            return ranked[0][1]

    # JS-heavy official sites: use the already-installed Playwright browser.
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(
                user_agent=(
                    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                    "Chrome/140.0.0.0 Safari/537.36"
                )
            )
            page.goto(home, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2500)
            soup = BeautifulSoup(page.content(), "html.parser")
            ranked = []
            for a in soup.find_all("a", href=True):
                u = normalize_url(a["href"], home)
                score = score_link(u, clean(a.get_text(" ", strip=True)))
                if score >= 5:
                    ranked.append((score, u))
            browser.close()
            if ranked:
                ranked.sort(reverse=True)
                return ranked[0][1]
    except Exception:
        pass

    return ""


def official_launch_confirmed(model, brand, domains, ndtv_html):
    """
    Verify the launch only on the official India brand domain.
    First use an official link inside the NDTV article; if NDTV does not
    provide one, locate the model page by crawling the official India site.
    """
    product_url = official_link_from_ndtv(ndtv_html, domains, model)

    if not product_url:
        product_url = official_product_url_from_home(model, brand, domains)

    # Last resort: try only deterministic URLs on the official India domain.
    if not product_url:
        for direct_url in official_direct_product_urls(model, brand):
            try:
                test_html = get_html(direct_url)
                if test_html:
                    soup = BeautifulSoup(test_html, "html.parser")
                    body = clean(soup.get_text(" ", strip=True)).lower()
                    words = [
                        w.lower() for w in re.findall(r"[A-Za-z0-9]+", model)
                        if len(w) >= 3
                    ]
                    if words and all(w in body for w in words[:2]):
                        product_url = direct_url
                        break
            except Exception:
                continue

    if not product_url:
        return "", ""

    official_html = get_html(product_url)
    if not official_html:
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page()
                page.goto(product_url, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(2000)
                official_html = page.content()
                browser.close()
        except Exception:
            official_html = ""

    if not official_html:
        return "", ""

    soup = BeautifulSoup(official_html, "html.parser")
    body = clean(soup.get_text(" ", strip=True))
    low = body.lower()
    model_low = model.lower()

    if model_low not in low:
        words = [
            w.lower() for w in re.findall(r"[A-Za-z0-9]+", model)
            if len(w) >= 3
        ]
        if not words or not all(w in low for w in words[:2]):
            return "", ""

    # Official product pages commonly prove India availability through
    # India-specific URLs plus buy/availability language rather than a
    # literal "launched in India" sentence.
    launch_terms = (
        "launched in india", "launches in india", "introduced in india",
        "available in india", "now available in india",
        "goes on sale in india", "sale begins in india", "available from",
        "available starting", "available beginning", "buy now", "add to cart",
        "out of stock", "notify me"
    )

    has_india = "india" in low or official_page_is_india(product_url)
    has_launch_signal = any(term in low for term in launch_terms)
    if not (has_india and has_launch_signal):
        return "", ""

    if not official_page_is_india(product_url) and "india" not in low:
        return "", ""

    return product_url, official_html

def official_page_is_india(url):
    low = url.lower()
    return any(x in low for x in ("/in/", "india", "in."))


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
    """Extract exact variants without blind RAM/storage/color Cartesian products."""
    soup = BeautifulSoup(html, "html.parser")
    text = clean(soup.get_text(" ", strip=True))
    found = []

    # Explicit combined configuration: 8GB + 256GB / 8GB RAM 256GB.
    combined = re.compile(
        r"(?P<ram>\d{1,2})\s*GB\s*(?:RAM)?\s*(?:\+|/|\||,|with)\s*"
        r"(?P<storage>\d{2,4})\s*(?P<unit>GB|TB)", re.I
    )
    pairs = []
    for m in combined.finditer(text):
        pair=(f"{m.group('ram')} GB", f"{m.group('storage')} {m.group('unit').upper()}")
        if pair not in pairs: pairs.append(pair)

    color_re = re.compile(r"(?:colour|color|colourway|colorway)\s*[:\-]?\s*([A-Za-z][A-Za-z0-9 &'’()./\-]{2,45})", re.I)
    blocks = soup.find_all(["li","div","section","article","label","option"])
    for block in blocks:
        t=clean(block.get_text(" ", strip=True))
        if len(t)>500: continue
        pm=combined.search(t); cm=color_re.search(t)
        if pm and cm:
            c=clean(cm.group(1))
            item=(c, f"{pm.group('storage')} {pm.group('unit').upper()}", f"{pm.group('ram')} GB")
            if item not in found: found.append(item)

    # Separate official declarations, e.g. RAM: 4 GB and Storage: 64 GB.
    ram_vals=[]; storage_vals=[]; colors=[]
    for m in re.finditer(r"(?:ram|memory)\s*[:\-]?\s*(\d{1,2})\s*GB", text, re.I):
        v=f"{m.group(1)} GB"
        if v not in ram_vals: ram_vals.append(v)
    for m in re.finditer(r"(?:storage|internal storage)\s*[:\-]?\s*(\d{2,4})\s*(GB|TB)", text, re.I):
        v=f"{m.group(1)} {m.group(2).upper()}"
        if v not in storage_vals: storage_vals.append(v)
    for m in color_re.finditer(text):
        c=clean(m.group(1))
        c=re.split(r"\s+(?:ram|storage|memory|price|available|select)\b",c,flags=re.I)[0].strip()
        if 2<=len(c)<=45 and c.lower() not in {x.lower() for x in colors}: colors.append(c)

    # Safe single-configuration rule: one explicit RAM + one explicit storage
    # can be paired with each explicitly listed product color. This is not a
    # Cartesian product because there is only one configuration.
    if not found and len(ram_vals)==1 and len(storage_vals)==1 and colors:
        for c in colors:
            item=(c, storage_vals[0], ram_vals[0])
            if item not in found: found.append(item)

    # JSON/data attributes can contain a single explicit configuration.
    if not found:
        raw_all=" ".join((tag.string or tag.get_text()) for tag in soup.find_all("script") if (tag.string or tag.get_text()))
        for pm in combined.finditer(raw_all):
            ram=f"{pm.group('ram')} GB"; storage=f"{pm.group('storage')} {pm.group('unit').upper()}"
            nearby=raw_all[max(0,pm.start()-1500):pm.end()+1500]
            cm=color_re.search(nearby)
            if cm:
                item=(clean(cm.group(1)),storage,ram)
                if item not in found: found.append(item)

    clean_found=[]
    for color,storage,ram in found:
        color=clean(color); storage=normalize_capacity(storage); ram=normalize_capacity(ram)
        if 2<=len(color)<=45 and color.lower() not in {"select","choose","available","variant","color"}:
            item=(color,storage,ram)
            if item not in clean_found: clean_found.append(item)
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
    print("RETAILIX DAILY SMARTPHONE AUTOMATION - FAST CLOUD VERSION")
    print("GOOGLE NEWS DISCOVERY + OFFICIAL INDIA BRAND VERIFICATION")
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
            for k in ("productName", "brandName", "variant_Color", "variant_Storage", "variant_RAM")
        ))

    end = now()
    start = end - timedelta(hours=24)
    print(f"Window: {start.isoformat()} -> {end.isoformat()}")
    print("Existing combinations:", len(existing))

    candidates = google_news_candidates()
    print("Google News candidates found:", len(candidates))

    grouped = {brand: [] for brand in BRANDS}
    for url, title in candidates:
        brand = candidate_brand(title)
        if not brand:
            continue
        dt = GOOGLE_DATES.get(url)
        if not dt or not (start <= dt <= end):
            continue
        if not any(w in title.lower() for w in LAUNCH_WORDS):
            continue
        if len(grouped[brand]) < 4:
            grouped[brand].append((url, title, dt))

    additions = []

    for brand, rows in grouped.items():
        print("\nChecking:", brand)
        if not rows:
            print("  No qualifying Google News candidate.")
            continue

        found_model = False
        for news_url, title, published in rows:
            model_hint = model_from_ndtv(title, brand)
            if len(model_hint) < 3:
                continue

            print("  Discovery:", title)
            print("  RSS date:", published.isoformat())

            # Verify launch/product data only on the official India brand site.
            domains = BRANDS[brand]
            official_url, official_html = official_launch_confirmed(
                model_hint, brand, domains, ""
            )

            if not official_url or not official_html:
                print("  SKIPPED: official India confirmation not found.")
                continue

            name = product_name_from_official(official_html, model_hint)
            if not name or len(name) < 3:
                print("  SKIPPED: official product name not found.")
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

            added_here = 0
            for color, storage, ram in variants:
                key = (
                    name.strip().lower(), brand.lower(), color.strip().lower(),
                    storage.strip().lower(), ram.strip().lower()
                )
                if key in existing:
                    continue

                additions.append({
                    "name": name, "brand": brand,
                    "short": description[:500], "long": description[:2000],
                    "color": color, "storage": storage, "ram": ram, "images": images
                })
                existing.add(key)
                added_here += 1

            print("  VERIFIED OFFICIAL:", name)
            print("  Official:", official_url)
            print("  Exact variants:", len(variants))
            print("  New rows:", added_here)
            found_model = True
            break

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
            p["name"], p["brand"], "Electronics", "Accessories", "piece",
            "NO TAX APPLIED", "NODISCOUNT0", "1 Year Warranty",
            p["short"], p["long"], "true", "false",
            p["color"], p["storage"], p["ram"],
            imgs[0] if len(imgs) > 0 else "", "true",
            imgs[1] if len(imgs) > 1 else "",
            imgs[2] if len(imgs) > 2 else "",
            imgs[3] if len(imgs) > 3 else "",
        ])

    wb.save(XLSX)
    print("\nRows added:", len(additions))
    print("AUTOMATION COMPLETED")


if __name__ == "__main__":
    main()
