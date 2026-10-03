from pathlib import Path
import os
import re
import shutil
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook


# ============================================================
# RETAILIX DAILY SMARTPHONE AUTOMATION
#
# DISCOVERY:
#   91MOBILES + GADGETS 360 DIRECTLY
#
# PRODUCT DATA:
#   OFFICIAL INDIA BRAND WEBSITE
#
# IMAGE:
#   91MOBILES - ONE PRIMARY PRODUCT IMAGE
#
# RULE:
#   SAME PHONE + SAME INDIA LAUNCH DATE MUST BE CONFIRMED
#   BY BOTH 91MOBILES AND GADGETS 360
#
# GOOGLE NEWS:
#   COMPLETELY REMOVED
# ============================================================

XLSX = Path("product_import_template (1).xlsx")

IST = timezone(timedelta(hours=5, minutes=30))


# ------------------------------------------------------------
# SELECTED BRANDS
# ------------------------------------------------------------

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
    "productName",
    "brandName",
    "categoryName",
    "subcategoryName",
    "unit_name",
    "taxName",
    "discountName",
    "warrantyName",
    "shortDescription",
    "longDescription",
    "is_featured",
    "is_digital",
    "variant_Color",
    "variant_Storage",
    "variant_RAM",
    "image1",
    "image1Primary",
    "image2",
    "image3",
    "image4",
]


SELECTED_BRANDS_LOWER = {
    brand: tuple(x.lower() for x in aliases)
    for brand, aliases in BRAND_ALIASES.items()
}


# ------------------------------------------------------------
# HTTP SESSION
# ------------------------------------------------------------

SESSION = requests.Session()

SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-IN,en;q=0.9",
})


# ------------------------------------------------------------
# BASIC HELPERS
# ------------------------------------------------------------

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

    return any(
        host == d or host.endswith("." + d)
        for d in domains
    )


def get_html(url, timeout=30):
    try:
        response = SESSION.get(
            url,
            timeout=timeout,
            allow_redirects=True,
        )

        if response.ok and len(response.text) > 300:
            return response.text

    except requests.RequestException:
        pass

    return ""


# ------------------------------------------------------------
# DATE PARSING
# ------------------------------------------------------------

def parse_date(text):
    text = clean(text)

    if not text:
        return None

    # ISO datetime/date
    iso_match = re.search(
        r"\b(20\d{2}-\d{1,2}-\d{1,2}"
        r"(?:[T ]\d{1,2}:\d{2}(?::\d{2})?"
        r"(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?)\b",
        text,
    )

    if iso_match:
        value = iso_match.group(1)

        try:
            if "T" in value or " " in value:
                dt = datetime.fromisoformat(
                    value.replace("Z", "+00:00")
                )

                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=IST)

                return dt.astimezone(IST)

            return datetime.strptime(
                value,
                "%Y-%m-%d",
            ).replace(tzinfo=IST)

        except ValueError:
            pass

    patterns = [
        (
            r"\b(\d{1,2})\s+([A-Za-z]+)\s+(20\d{2})\b",
            "%d %B %Y",
        ),
        (
            r"\b([A-Za-z]+)\s+(\d{1,2}),\s*(20\d{2})\b",
            "%B %d, %Y",
        ),
        (
            r"\b(\d{1,2})[/-](\d{1,2})[/-](20\d{2})\b",
            "%d/%m/%Y",
        ),
        (
            r"\b(\d{1,2})\s+([A-Za-z]{3})\s+(20\d{2})\b",
            "%d %b %Y",
        ),
    ]

    for pattern, fmt in patterns:

        match = re.search(
            pattern,
            text,
            re.I,
        )

        if not match:
            continue

        try:
            return datetime.strptime(
                match.group(0),
                fmt,
            ).replace(tzinfo=IST)

        except ValueError:
            continue

    return None


def article_date(html):

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    for attrs in (
        {"property": "article:published_time"},
        {"property": "og:article:published_time"},
        {"name": "publish-date"},
        {"name": "date"},
        {"itemprop": "datePublished"},
    ):

        meta = soup.find(
            "meta",
            attrs=attrs,
        )

        if meta:

            value = (
                meta.get("content")
                or meta.get("datetime")
            )

            if value:

                try:
                    dt = datetime.fromisoformat(
                        value.replace("Z", "+00:00")
                    )

                    if dt.tzinfo is None:
                        dt = dt.replace(
                            tzinfo=IST
                        )

                    return dt.astimezone(IST)

                except ValueError:

                    parsed = parse_date(value)

                    if parsed:
                        return parsed

    for tag in soup.find_all("time"):

        value = (
            tag.get("datetime")
            or tag.get_text(
                " ",
                strip=True,
            )
        )

        parsed = parse_date(value)

        if parsed:
            return parsed

    return None


def get_release_window():

    test_date = os.getenv(
        "RETAILIX_TEST_DATE",
        "",
    ).strip()

    # --------------------------------------------------------
    # TEST MODE
    # --------------------------------------------------------

    if test_date:

        try:
            day = datetime.strptime(
                test_date,
                "%Y-%m-%d",
            ).date()

        except ValueError as exc:

            raise RuntimeError(
                "RETAILIX_TEST_DATE must use "
                "YYYY-MM-DD, e.g. 2026-10-01"
            ) from exc

        start = datetime.combine(
            day,
            datetime.min.time(),
            tzinfo=IST,
        )

        end = (
            start
            + timedelta(days=1)
            - timedelta(microseconds=1)
        )

        return start, end, True

    # --------------------------------------------------------
    # PRODUCTION MODE
    # --------------------------------------------------------

    end = now()

    start = end - timedelta(hours=24)

    return start, end, False


# ------------------------------------------------------------
# BRAND / MODEL HELPERS
# ------------------------------------------------------------

def detect_brand(text):

    low = clean(text).lower()

    for brand, aliases in SELECTED_BRANDS_LOWER.items():

        for alias in aliases:

            if re.search(
                r"\b" + re.escape(alias) + r"\b",
                low,
            ):
                return brand

    return None


def normalize_model_name(
    name,
    brand=None,
):

    name = clean(name)

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

    name = re.sub(
        r"\s*[:|,-]\s*$",
        "",
        name,
    )

    name = re.sub(
        r"\s+",
        " ",
        name,
    ).strip(
        " -:|,"
    )

    name = re.sub(
        r"^(?:the|new)\s+",
        "",
        name,
        flags=re.I,
    )

    if brand in {"REDMI", "POCO"}:

        name = re.sub(
            r"^Xiaomi\s+",
            "",
            name,
            flags=re.I,
        )

    return name


def model_key(name):

    low = clean(name).lower()

    low = re.sub(
        r"[^a-z0-9]+",
        " ",
        low,
    )

    low = re.sub(
        r"\b(5g|4g|india|phone|smartphone|mobile)\b",
        " ",
        low,
    )

    return re.sub(
        r"\s+",
        " ",
        low,
    ).strip()


# ------------------------------------------------------------
# LAUNCH VERIFICATION
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


NEGATIVE_LAUNCH_WORDS = (
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
)


ARTICLE_EXCLUDE_WORDS = (
    "price hike",
    "price cut",
    "deal",
    "discount",
    "review",
    "first impressions",
    "comparison",
    "storage variant",
    "ram and storage variant",
    "price drop",
    "sale only",
)


SMARTPHONE_TITLE_WORDS = (
    "phone",
    "smartphone",
    "iphone",
    "galaxy",
    "redmi",
    "poco",
    "oneplus",
    "oppo",
    "vivo",
    "realme",
    "motorola",
    "moto",
    "infinix",
    "narzo",
)


def article_is_real_india_launch(
    title,
    body,
):

    title_low = clean(title).lower()

    early = clean(
        body[:5000]
    ).lower()

    if any(
        x in title_low
        for x in NEGATIVE_LAUNCH_WORDS
    ):
        return False

    if any(
        x in title_low
        for x in ARTICLE_EXCLUDE_WORDS
    ):
        return False

    if any(
        x in early
        for x in STRONG_LAUNCH_WORDS
    ):
        return True

    availability_words = (
        "now available in india",
        "available in india",
        "launched for the indian market",
        "now launched in india",
    )

    return any(
        x in early
        for x in availability_words
    )


def extract_launch_date_from_text(
    text,
    fallback,
):

    text = clean(text)

    patterns = [

        (
            r"(?:launched|launches|launch)"
            r"\s+in\s+india\s+(?:on\s+)?"
            r"(\d{1,2}(?:st|nd|rd|th)?"
            r"\s+[A-Za-z]+\s+20\d{2})"
        ),

        (
            r"(?:launched|launches|launch)"
            r"\s+in\s+india\s+(?:on\s+)?"
            r"([A-Za-z]+\s+\d{1,2},\s*20\d{2})"
        ),

        (
            r"(?:launched|launches|launch)"
            r"\s+in\s+india\s+(?:on\s+)?"
            r"(\d{1,2}\s+[A-Za-z]{3,9},?\s+20\d{2})"
        ),
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.I,
        )

        if match:

            value = re.sub(
                r"(\d)(st|nd|rd|th)",
                r"\1",
                match.group(1),
            )

            parsed = parse_date(value)

            if parsed:
                return parsed

    return fallback


# ------------------------------------------------------------
# DIRECT 91MOBILES DISCOVERY
# ------------------------------------------------------------

NINETY_ONE_MOBILES_URL = (
    "https://www.91mobiles.com/"
)


def get_91mobiles_article_links():

    urls = [

        "https://www.91mobiles.com/"
        "new-phones",

        "https://www.91mobiles.com/"
        "hub/",

        "https://www.91mobiles.com/"
        "hub/category/news/",
    ]

    links = []

    for page_url in urls:

        html = get_html(
            page_url,
            timeout=30,
        )

        if not html:
            continue

        soup = BeautifulSoup(
            html,
            "html.parser",
        )

        for a in soup.find_all(
            "a",
            href=True,
        ):

            href = normalize_url(
                a.get("href"),
                page_url,
            )

            if not href:
                continue

            parsed = urlparse(href)

            if "91mobiles.com" not in parsed.netloc.lower():
                continue

            path = parsed.path.lower()

            if not path.startswith("/hub/"):
                continue

            if path in (
                "/hub/",
                "/hub/search/",
            ):
                continue

            text = clean(
                a.get_text(
                    " ",
                    strip=True,
                )
            )

            links.append(
                (href, text)
            )

    # Remove duplicates
    seen = set()
    result = []

    for href, text in links:

        if href in seen:
            continue

        seen.add(href)

        result.append(
            (href, text)
        )

    return result


# ------------------------------------------------------------
# DIRECT GADGETS 360 DISCOVERY
# ------------------------------------------------------------

GADGETS360_URLS = [

    "https://www.gadgets360.com/"
    "mobiles/news",

    "https://www.gadgets360.com/"
    "mobiles",

]


def get_gadgets360_article_links():

    links = []

    for page_url in GADGETS360_URLS:

        html = get_html(
            page_url,
            timeout=30,
        )

        if not html:
            continue

        soup = BeautifulSoup(
            html,
            "html.parser",
        )

        for a in soup.find_all(
            "a",
            href=True,
        ):

            href = normalize_url(
                a.get("href"),
                page_url,
            )

            if not href:
                continue

            parsed = urlparse(href)

            if (
                "gadgets360.com"
                not in parsed.netloc.lower()
            ):
                continue

            path = parsed.path.lower()

            if not path.startswith(
                "/mobiles/news/"
            ):
                continue

            text = clean(
                a.get_text(
                    " ",
                    strip=True,
                )
            )

            links.append(
                (href, text)
            )

    seen = set()
    result = []

    for href, text in links:

        if href in seen:
            continue

        seen.add(href)

        result.append(
            (href, text)
        )

    return result


# ------------------------------------------------------------
# VERIFY A PUBLISHER ARTICLE
# ------------------------------------------------------------

def verify_publisher_article(
    candidate,
    start,
    end,
):

    url = candidate.get(
        "url",
        "",
    )

    if not url:
        return None

    html = get_html(
        url,
        timeout=30,
    )

    if not html:
        return None

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    body = clean(
        soup.get_text(
            " ",
            strip=True,
        )
    )

    h1 = soup.find("h1")

    publisher_title = (
        clean(
            h1.get_text(
                " ",
                strip=True,
            )
        )
        if h1
        else candidate.get(
            "title",
            "",
        )
    )

    if len(publisher_title) < 10:
        publisher_title = candidate.get(
            "title",
            "",
        )

    published = (
        article_date(html)
        or candidate.get(
            "published"
        )
    )

    if not published:
        return None

    title_low = publisher_title.lower()

    if any(
        x in title_low
        for x in ARTICLE_EXCLUDE_WORDS
    ):
        return None

    title_has_phone = any(
        x in title_low
        for x in SMARTPHONE_TITLE_WORDS
    )

    title_has_launch = any(
        x in title_low
        for x in STRONG_LAUNCH_WORDS
    )

    if not (
        title_has_phone
        and title_has_launch
    ):
        return None

    if not article_is_real_india_launch(
        publisher_title,
        body,
    ):
        return None

    brand = detect_brand(
        publisher_title
    )

    if not brand:

        brand = detect_brand(
            body[:15000]
        )

    if not brand:
        return None

    launch_date = extract_launch_date_from_text(
        body,
        published,
    )

    if not (
        start
        <= launch_date
        <= end
    ):
        return None

    model = normalize_model_name(
        publisher_title,
        brand,
    )

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


# ------------------------------------------------------------
# SOURCE DISCOVERY
# ------------------------------------------------------------

def discover_source_candidates(
    source,
    start,
    end,
):

    if source == "91mobiles":

        raw_links = (
            get_91mobiles_article_links()
        )

    else:

        raw_links = (
            get_gadgets360_article_links()
        )

    out = {}

    for url, title in raw_links:

        candidate = {
            "url": url,
            "title": title,
            "source": source,
        }

        verified = verify_publisher_article(
            candidate,
            start,
            end,
        )

        if not verified:
            continue

        key = (
            verified["brand"],
            model_key(
                verified["name"]
            ),
        )

        previous = out.get(key)

        if (
            previous is None
            or verified["article_date"]
            > previous["article_date"]
        ):

            out[key] = verified

    return list(
        out.values()
    )


# ------------------------------------------------------------
# TWO-SOURCE CROSS-CHECK
# ------------------------------------------------------------

def discover_launches(
    start,
    end,
):

    n91 = discover_source_candidates(
        "91mobiles",
        start,
        end,
    )

    g360 = discover_source_candidates(
        "gadgets360",
        start,
        end,
    )

    print(
        "91mobiles verified launch candidates:",
        len(n91),
    )

    print(
        "Gadgets 360 verified launch candidates:",
        len(g360),
    )

    g360_map = {
        (
            item["brand"],
            model_key(
                item["name"]
            ),
        ): item
        for item in g360
    }

    qualified = []

    for item in n91:

        key = (
            item["brand"],
            model_key(
                item["name"]
            ),
        )

        other = g360_map.get(key)

        if not other:

            print(
                f"  CROSS-CHECK WAIT: "
                f"{item['brand']} "
                f"{item['name']} "
                f"({item['release'].date()}) "
                f"not confirmed by Gadgets 360."
            )

            continue

        if (
            other["release"].date()
            != item["release"].date()
        ):

            print(
                f"  DATE MISMATCH: "
                f"{item['name']} | "
                f"91mobiles="
                f"{item['release'].date()} | "
                f"Gadgets360="
                f"{other['release'].date()}"
            )

            continue

        item["g360_url"] = (
            other["url"]
        )

        item["g360_release"] = (
            other["release"]
        )

        item["discovery_date"] = (
            item["release"]
        )

        qualified.append(
            item
        )

        print(
            f"  QUALIFIED: "
            f"{item['brand']} "
            f"{item['name']} | "
            f"India launch date: "
            f"{item['release'].date()}"
        )

    return qualified


# ------------------------------------------------------------
# OFFICIAL INDIA BRAND SEARCH
# ------------------------------------------------------------

def official_page_is_india(url):

    low = url.lower()

    host = urlparse(
        url
    ).netloc.lower()

    if (
        "in." in host
        or host.endswith(".in")
    ):
        return True

    return any(
        token in low
        for token in (
            "/in/",
            "/india/",
            "india.",
            "/india",
            "?country=in",
            "country=in",
            "locale=in",
        )
    )


def official_search_urls(
    model,
    domains,
):

    urls = []

    for domain in domains:

        query = (
            f'site:{domain} '
            f'"{model}" '
            f'smartphone India'
        )

        encoded = requests.utils.quote(
            query
        )

        urls.append(
            "https://www.google.com/search"
            f"?q={encoded}&num=10"
        )

    return urls


def extract_search_result_links(
    html,
    allowed_domains,
):

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    results = []

    for a in soup.find_all(
        "a",
        href=True,
    ):

        href = a.get(
            "href",
            "",
        )

        text = clean(
            a.get_text(
                " ",
                strip=True,
            )
        )

        if href.startswith(
            "/url?q="
        ):

            href = href.split(
                "/url?q=",
                1,
            )[1].split(
                "&",
                1,
            )[0]

        href = normalize_url(
            href
        )

        if not href:
            continue

        if same_domain(
            href,
            allowed_domains,
        ):

            results.append(
                (
                    href,
                    text,
                )
            )

    seen = set()
    output = []

    for href, text in results:

        if href in seen:
            continue

        seen.add(href)

        output.append(
            (
                href,
                text,
            )
        )

    return output


def find_official_product_page(
    model,
    brand,
):

    domains = BRANDS[
        brand
    ]

    candidates = []

    for search_url in official_search_urls(
        model,
        domains,
    ):

        html = get_html(
            search_url,
            timeout=20,
        )

        if not html:
            continue

        for href, text in extract_search_result_links(
            html,
            domains,
        ):

            blob = (
                href
                + " "
                + text
            ).lower()

            score = 0

            if official_page_is_india(
                href
            ):
                score += 8

            if any(
                x in blob
                for x in (
                    "phone",
                    "smartphone",
                    "mobile",
                    "product",
                )
            ):
                score += 3

            for word in re.findall(
                r"[a-z0-9]+",
                model.lower(),
            ):

                if (
                    len(word) >= 3
                    and word in blob
                ):
                    score += 1

            if any(
                x in href.lower()
                for x in (
                    "/support",
                    "/help",
                    "/careers",
                    "/about",
                    "/privacy",
                    "/terms",
                    "/store-locator",
                )
            ):
                score -= 10

            candidates.append(
                (
                    score,
                    href,
                )
            )

    candidates.sort(
        reverse=True
    )

    for score, url in candidates:

        if score < 5:
            continue

        html = get_html(
            url,
            timeout=30,
        )

        if not html:
            continue

        soup = BeautifulSoup(
            html,
            "html.parser",
        )

        body = clean(
            soup.get_text(
                " ",
                strip=True,
            )
        )

        low = body.lower()

        model_words = [
            w.lower()
            for w in re.findall(
                r"[A-Za-z0-9]+",
                model,
            )
            if len(w) >= 3
        ]

        if (
            model_words
            and not all(
                w in low
                for w in model_words[:2]
            )
        ):
            continue

        if (
            not official_page_is_india(
                url
            )
            and "india" not in low
        ):
            continue

        return (
            url,
            html,
        )

    return (
        "",
        "",
    )


# ------------------------------------------------------------
# OFFICIAL PRODUCT DESCRIPTION
# ------------------------------------------------------------

def product_name_from_official(
    html,
    fallback,
):

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    h1 = soup.find(
        "h1"
    )

    if h1:

        value = clean(
            h1.get_text(
                " ",
                strip=True,
            )
        )

        if 2 <= len(value) <= 120:
            return value

    for attrs in (
        {"property": "og:title"},
        {"name": "twitter:title"},
    ):

        meta = soup.find(
            "meta",
            attrs=attrs,
        )

        if meta and meta.get(
            "content"
        ):

            value = clean(
                meta["content"]
            )

            value = re.sub(
                r"\s*[-|]\s*"
                r"(Official.*|Buy.*|India.*)$",
                "",
                value,
                flags=re.I,
            )

            if len(value) >= 2:
                return value

    return clean(
        fallback
    )


def official_description(
    html,
    model,
):

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    for attrs in (
        {"name": "description"},
        {"property": "og:description"},
        {"name": "twitter:description"},
    ):

        meta = soup.find(
            "meta",
            attrs=attrs,
        )

        if meta and meta.get(
            "content"
        ):

            text = clean(
                meta["content"]
            )

            if len(text) >= 60:
                return text[:2000]

    paragraphs = []

    for p in soup.find_all(
        ["p", "li"]
    ):

        text = clean(
            p.get_text(
                " ",
                strip=True,
            )
        )

        low = text.lower()

        if (
            len(text) < 60
            or len(text) > 1200
        ):
            continue

        if (
            model.lower() not in low
            and "smartphone" not in low
            and "phone" not in low
            and "mobile" not in low
        ):
            continue

        paragraphs.append(
            text
        )

    if paragraphs:
        return " ".join(
            paragraphs[:4]
        )[:2000]

    return ""


# ------------------------------------------------------------
# VARIANTS
# ------------------------------------------------------------

def normalize_capacity(value):

    value = clean(
        value
    ).upper()

    return re.sub(
        r"\s+",
        " ",
        value,
    )


def extract_variants(
    html,
    model,
):

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    found = []

    ram_storage = re.compile(
        r"(?P<ram>\d{1,2})\s*GB"
        r"\s*(?:RAM)?\s*"
        r"(?:\+|/|\||,|with)\s*"
        r"(?P<storage>\d{2,4})\s*"
        r"(?P<unit>GB|TB)",
        re.I,
    )

    ram_storage_2 = re.compile(
        r"(?P<ram>\d{1,2})\s*GB\s*RAM"
        r".{0,100}?"
        r"(?P<storage>\d{2,4})\s*"
        r"(?P<unit>GB|TB)"
        r"\s*(?:storage|memory)?",
        re.I,
    )

    color_re = re.compile(
        r"(?:colour|color|colourway|colorway)"
        r"\s*[:\-]?\s*"
        r"([A-Za-z][A-Za-z0-9 &'’()./\-]{2,45})",
        re.I,
    )

    blocks = soup.find_all(
        [
            "li",
            "div",
            "section",
            "article",
            "label",
            "option",
            "tr",
            "td",
        ]
    )

    for block in blocks:

        block_text = clean(
            block.get_text(
                " ",
                strip=True,
            )
        )

        if (
            not block_text
            or len(block_text) > 700
        ):
            continue

        pm = (
            ram_storage.search(
                block_text
            )
            or
            ram_storage_2.search(
                block_text
            )
        )

        cm = color_re.search(
            block_text
        )

        if pm and cm:

            ram = (
                f"{pm.group('ram')} GB"
            )

            storage = (
                f"{pm.group('storage')} "
                f"{pm.group('unit').upper()}"
            )

            color = clean(
                cm.group(1)
            )

            color = re.split(
                r"\s+(?:ram|storage|memory|price|₹|\|)\b",
                color,
                flags=re.I,
            )[0].strip()

            if (
                2 <= len(color) <= 45
                and color.lower()
                not in {
                    "select",
                    "choose",
                    "available",
                    "variant",
                    "color",
                }
            ):

                item = (
                    color,
                    normalize_capacity(
                        storage
                    ),
                    normalize_capacity(
                        ram
                    ),
                )

                if item not in found:
                    found.append(
                        item
                    )

    return found


# ------------------------------------------------------------
# 91MOBILES PRIMARY IMAGE
# ------------------------------------------------------------

def score_91_image(
    url,
    alt="",
    width=0,
    height=0,
):

    if (
        not url
        or not url.startswith("http")
    ):
        return -999

    u = url.lower()
    a = alt.lower()

    bad = (
        "logo",
        "icon",
        "favicon",
        "sprite",
        "banner",
        "advert",
        "ad-",
        "offer",
        "amazon",
        "flipkart",
        "review",
        "screenshot",
        "avatar",
        "placeholder",
        "video",
        "gallery-thumb",
    )

    if any(
        x in u or x in a
        for x in bad
    ):
        return -999

    score = 0

    if any(
        x in u
        for x in (
            "phone",
            "mobile",
            "product",
            "device",
        )
    ):
        score += 5

    if (
        width >= 700
        and height >= 700
    ):
        score += 5

    elif (
        width >= 400
        and height >= 400
    ):
        score += 3

    if (
        width
        and height
        and height >= width * 0.75
    ):
        score += 2

    return score


def extract_91_primary_image(
    page_url,
    model,
):

    html = get_html(
        page_url,
        timeout=30,
    )

    if not html:
        return ""

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    candidates = []

    for attrs in (
        {"property": "og:image"},
        {"name": "twitter:image"},
    ):

        meta = soup.find(
            "meta",
            attrs=attrs,
        )

        if (
            meta
            and meta.get("content")
        ):

            url = normalize_url(
                meta.get("content"),
                page_url,
            )

            score = score_91_image(
                url,
                model,
            )

            if score >= 0:
                candidates.append(
                    (
                        score + 4,
                        url,
                    )
                )

    for tag in soup.find_all(
        ["img", "source"]
    ):

        alt = (
            tag.get("alt", "")
            or tag.get("title", "")
        )

        urls = []

        for attr in (
            "src",
            "data-src",
            "data-original",
            "data-lazy-src",
            "data-image",
            "data-zoom-image",
        ):

            if tag.get(attr):
                urls.append(
                    tag.get(attr)
                )

        for part in (
            tag.get("srcset") or ""
        ).split(","):

            piece = (
                part.strip()
                .split(" ")[0]
            )

            if piece:
                urls.append(
                    piece
                )

        width = (
            int(tag.get("width"))
            if str(
                tag.get("width") or ""
            ).isdigit()
            else 0
        )

        height = (
            int(tag.get("height"))
            if str(
                tag.get("height") or ""
            ).isdigit()
            else 0
        )

        for raw in urls:

            url = normalize_url(
                raw,
                page_url,
            )

            score = score_91_image(
                url,
                alt,
                width,
                height,
            )

            if score >= 0:
                candidates.append(
                    (
                        score,
                        url,
                    )
                )

    if not candidates:
        return ""

    best = {}

    for score, url in candidates:

        best[url] = max(
            score,
            best.get(
                url,
                -999,
            ),
        )

    ranked = sorted(
        (
            (score, url)
            for url, score in best.items()
        ),
        reverse=True,
    )

    return (
        ranked[0][1]
        if ranked
        else ""
    )


# ------------------------------------------------------------
# EXCEL
# ------------------------------------------------------------

def backup_excel():

    stamp = now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup = XLSX.with_name(
        f"{XLSX.stem}_backup_"
        f"{stamp}{XLSX.suffix}"
    )

    shutil.copy2(
        XLSX,
        backup,
    )

    print(
        "Backup:",
        backup.name,
    )


def load_existing(ws):

    cols = {
        ws.cell(1, c).value: c
        for c in range(
            1,
            ws.max_column + 1,
        )
    }

    required = {
        "productName",
        "brandName",
        "variant_Color",
        "variant_Storage",
        "variant_RAM",
    }

    if not required.issubset(
        cols
    ):
        raise RuntimeError(
            "Required Excel columns are missing."
        )

    existing = set()

    for row in ws.iter_rows(
        min_row=2,
        values_only=True,
    ):

        key = tuple(
            str(
                row[
                    cols[k] - 1
                ]
                or ""
            ).strip().lower()

            for k in (
                "productName",
                "brandName",
                "variant_Color",
                "variant_Storage",
                "variant_RAM",
            )
        )

        existing.add(
            key
        )

    return existing


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():

    print("=" * 78)

    print(
        "RETAILIX DAILY SMARTPHONE AUTOMATION"
    )

    print(
        "DISCOVERY: 91MOBILES + GADGETS 360"
    )

    print(
        "GOOGLE NEWS: REMOVED"
    )

    print(
        "DATA: OFFICIAL INDIA BRAND WEBSITE"
    )

    print(
        "IMAGE: 91MOBILES - ONE PRIMARY IMAGE"
    )

    print("=" * 78)

    if not XLSX.exists():

        raise FileNotFoundError(
            f"Excel file not found: "
            f"{XLSX.resolve()}"
        )

    wb = load_workbook(
        XLSX
    )

    if (
        "PRODUCT_DATA"
        not in wb.sheetnames
    ):
        raise RuntimeError(
            "Worksheet "
            "'PRODUCT_DATA' "
            "was not found."
        )

    ws = wb[
        "PRODUCT_DATA"
    ]

    headers = [
        ws.cell(
            1,
            c,
        ).value
        for c in range(
            1,
            21,
        )
    ]

    if headers != HEADERS:

        raise RuntimeError(
            "RetailiX template headers "
            "do not match the expected "
            "20 columns."
        )

    existing = load_existing(
        ws
    )

    start, end, test_mode = (
        get_release_window()
    )

    if test_mode:

        print(
            "TEST MODE ENABLED: "
            "checking the full India "
            "calendar day "
            f"{start.date()}"
        )

    else:

        print(
            "24-HOUR INDIA RELEASE WINDOW: "
            f"{start.isoformat()} -> "
            f"{end.isoformat()}"
        )

    print(
        "Existing combinations:",
        len(existing),
    )

    # --------------------------------------------------------
    # DISCOVERY
    # --------------------------------------------------------

    launches = discover_launches(
        start,
        end,
    )

    if not launches:

        if test_mode:

            print(
                "\nNo selected-brand phone "
                "was independently confirmed "
                "by BOTH 91mobiles and "
                f"Gadgets 360 on "
                f"{start.date()}."
            )

        else:

            print(
                "\nNo selected-brand phone "
                "was independently confirmed "
                "by BOTH 91mobiles and "
                "Gadgets 360 in the last "
                "24 hours."
            )

        print(
            "AUTOMATION COMPLETED - "
            "NO EXCEL CHANGES."
        )

        return

    print(
        f"\nQualified launches: "
        f"{len(launches)}"
    )

    additions = []

    # --------------------------------------------------------
    # OFFICIAL INDIA DATA
    # --------------------------------------------------------

    for item in launches:

        brand = item["brand"]

        model = item["name"]

        print(
            "\n" + "-" * 70
        )

        print(
            f"PROCESSING: "
            f"{brand} | {model}"
        )

        print(
            "India release date:",
            item["discovery_date"].date(),
        )

        print(
            "91mobiles:",
            item["url"],
        )

        print(
            "Gadgets 360:",
            item.get(
                "g360_url",
                "",
            ),
        )

        official_url, official_html = (
            find_official_product_page(
                model,
                brand,
            )
        )

        if (
            not official_url
            or not official_html
        ):

            print(
                "SKIPPED: official India "
                "product page not found."
            )

            continue

        print(
            "Official India page:",
            official_url,
        )

        name = product_name_from_official(
            official_html,
            model,
        )

        description = official_description(
            official_html,
            name,
        )

        variants = extract_variants(
            official_html,
            name,
        )

        if len(description) < 60:

            print(
                "SKIPPED: official product "
                "description not verified."
            )

            continue

        if not variants:

            print(
                "SKIPPED: official RAM + "
                "Storage + Colour "
                "combinations could not "
                "be verified."
            )

            continue

        # ----------------------------------------------------
        # PRIMARY IMAGE
        # ----------------------------------------------------

        image = extract_91_primary_image(
            item["url"],
            model,
        )

        if not image:

            print(
                "SKIPPED: primary 91mobiles "
                "product image not found."
            )

            continue

        print(
            "Primary image:",
            image,
        )

        print(
            "Official variants:",
            len(variants),
        )

        # ----------------------------------------------------
        # DUPLICATES
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
                    "  DUPLICATE:",
                    color,
                    "|",
                    storage,
                    "|",
                    ram,
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

            existing.add(
                key
            )

            added_for_phone += 1

        print(
            "New variants to add:",
            added_for_phone,
        )

    # --------------------------------------------------------
    # WRITE EXCEL
    # --------------------------------------------------------

    if not additions:

        print(
            "\nNo new variants to add."
        )

        print(
            "AUTOMATION COMPLETED - "
            "NO EXCEL CHANGES."
        )

        return

    backup_excel()

    for product in additions:

        ws.append([

            product["name"],

            product["brand"],

            "Electronics",

            "Mobiles",

            "piece",

            "NO TAX APPLIED",

            "",

            "NO WARRANTY APPLIED",

            product["short"],

            product["long"],

            "true",

            "false",

            product["color"],

            product["storage"],

            product["ram"],

            product["image"],

            "true",

            "",

            "",

            "",
        ])

    wb.save(
        XLSX
    )

    print(
        "\n" + "=" * 78
    )

    print(
        f"ROWS ADDED: "
        f"{len(additions)}"
    )

    print(
        f"EXCEL SAVED: "
        f"{XLSX}"
    )

    print(
        "image1 = 91mobiles primary "
        "product image"
    )

    print(
        "image1Primary = true"
    )

    print(
        "image2/image3/image4 = blank"
    )

    print(
        "AUTOMATION COMPLETED "
        "SUCCESSFULLY"
    )

    print(
        "=" * 78
    )


if __name__ == "__main__":
    main()
