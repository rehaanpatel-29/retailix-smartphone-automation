from __future__ import annotations

import re
import shutil
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse, quote_plus

import requests
import feedparser
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
XLSX = BASE / "product_import_template (1).xlsx"
IST = timezone(timedelta(hours=5, minutes=30))

HEAD = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/142.0 Safari/537.36",
    "Accept-Language": "en-IN,en;q=0.9",
}
S = requests.Session()
S.headers.update(HEAD)

BRANDS = {
    "APPLE": {
        "domains": ["apple.com"],
        "sources": [
            "https://www.apple.com/in/newsroom/archive/",
            "https://www.apple.com/in/newsroom/archive/iphone/",
        ],
    },
    "SAMSUNG": {
        "domains": ["samsung.com"],
        "sources": [
            "https://news.samsung.com/in/category/products",
            "https://news.samsung.com/in/category/press-resources/press-release",
            "https://news.samsung.com/in/latest",
        ],
    },
    "VIVO": {"domains": ["vivo.com"], "sources": ["https://vivonewsroom.in/"]},
    "OPPO": {
        "domains": ["oppo.com"],
        "sources": [
            "https://www.oppo.com/in/newsroom/",
            "https://www.oppo.com/in/newsroom/all/",
        ],
    },
    "ONEPLUS": {"domains": ["oneplus.com"], "sources": ["https://www.oneplus.com/in"]},
    "REALME": {"domains": ["realme.com"], "sources": ["https://www.realme.com/in/brand/newsroom"]},
    "MOTOROLA": {
        "domains": ["motorola.in", "motorola.com"],
        "sources": ["https://www.motorola.in/"],
    },
    "INFINIX": {
        "domains": ["infinixmobility.com", "infinixmobiles.in"],
        "sources": ["https://www.infinixmobility.com/news/", "https://infinixmobiles.in/"],
    },
    "POCO": {"domains": ["mi.com"], "sources": ["https://www.mi.com/in/", "https://www.mi.com/in/phone/"]},
    "REDMI": {"domains": ["mi.com"], "sources": ["https://www.mi.com/in/", "https://www.mi.com/in/phone/"]},
}

SMART = (
    "smartphone", "iphone", "galaxy", "redmi", "poco", "mobile", "phone",
    "reno", "find x", "find n", "nord", "oneplus", "narzo", "realme",
    "moto", "edge", "razr", "zero", "note"
)
LAUNCH = (
    "launch", "launched", "launches", "introduces", "introduced",
    "unveils", "unveiled", "debuts", "debut", "announces", "announced",
    "available", "goes on sale", "sale begins", "now available",
    "arrives in stores", "available in india", "launched in india",
)
BAD = (
    "upcoming", "rumour", "rumor", "leak", "leaked", "expected",
    "may launch", "could launch", "might launch", "tipped", "teased",
    "teaser", "concept", "review", "comparison", "price hike",
    "discount", "offer", "campaign", "event", "hands-on", "speculation",
)
HEADERS = [
    "productName", "brandName", "categoryName", "subcategoryName", "unit_name",
    "taxName", "discountName", "warrantyName", "shortDescription",
    "longDescription", "is_featured", "is_digital", "variant_Color",
    "variant_Storage", "variant_RAM", "image1", "image1Primary",
    "image2", "image3", "image4"
]


def now():
    return datetime.now(timezone.utc).astimezone(IST)


def get(url, timeout=12):
    try:
        r = S.get(url, timeout=timeout, allow_redirects=True)
        if r.status_code == 200 and "text" in r.headers.get("content-type", "").lower():
            return r.text
    except requests.RequestException:
        pass
    return ""


def norm(url, base):
    return urljoin(base, url).split("#")[0]


def official(url, domains):
    host = urlparse(url).netloc.lower().split(":")[0]
    return any(host == d or host.endswith("." + d) for d in domains)


def clean(x):
    return re.sub(r"\s+", " ", x or "").strip()


def parse_date(value):
    if not value:
        return None
    value = clean(value)
    try:
        d = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return (d if d.tzinfo else d.replace(tzinfo=timezone.utc)).astimezone(IST)
    except Exception:
        pass
    for fmt in ("%d %B %Y", "%B %d, %Y", "%d %b %Y", "%b %d, %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).replace(tzinfo=IST)
        except Exception:
            pass
    return None


def article_date(html):
    soup = BeautifulSoup(html, "html.parser")
    values = []

    keys = (
        "article:published_time", "article:modified_time", "datepublished",
        "date", "published", "publishdate", "dc.date", "datecreated",
        "parsely-pub-date"
    )
    for m in soup.find_all("meta"):
        key = clean(m.get("property") or m.get("name")).lower()
        val = m.get("content")
        if val and any(k in key for k in keys):
            values.append(val)

    for t in soup.find_all("time"):
        if t.get("datetime"):
            values.append(t.get("datetime"))
        values.append(t.get_text(" ", strip=True))

    for script in soup.find_all("script", type="application/ld+json"):
        raw = script.string or script.get_text()
        for field in ("datePublished", "dateCreated", "dateModified"):
            values += re.findall(rf'"{field}"\s*:\s*"([^"]+)"', raw, re.I)

    for value in values:
        d = parse_date(value)
        if d:
            return d

    text = soup.get_text(" ", strip=True)
    patterns = [
        r"\b(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})\b",
        r"\b((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})\b",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            d = parse_date(m.group(1))
            if d:
                return d
    return None


def title_of(html, fallback=""):
    soup = BeautifulSoup(html, "html.parser")
    m = soup.find("meta", property="og:title")
    if m and m.get("content"):
        return clean(m["content"])
    h1 = soup.find("h1")
    if h1:
        return clean(h1.get_text(" ", strip=True))
    return clean(soup.title.get_text(" ", strip=True) if soup.title else fallback)


def brand_in_title(title, brand):
    aliases = {
        "APPLE": ("apple", "iphone"), "SAMSUNG": ("samsung", "galaxy"),
        "VIVO": ("vivo",), "OPPO": ("oppo",), "ONEPLUS": ("oneplus", "one plus"),
        "REALME": ("realme",), "MOTOROLA": ("motorola", "moto"),
        "INFINIX": ("infinix",), "POCO": ("poco",), "REDMI": ("redmi", "xiaomi"),
    }
    return any(x in title.lower() for x in aliases[brand])


def qualifying(title, html, brand, dt, start, end):
    if not dt or not (start <= dt <= end):
        return False
    if not brand_in_title(title, brand):
        return False
    if any(x in title.lower() for x in BAD):
        return False

    body = BeautifulSoup(html, "html.parser").get_text(" ", strip=True).lower()
    blob = title.lower() + " " + body

    if not any(x in blob for x in SMART):
        return False
    if not any(x in blob for x in LAUNCH):
        return False
    if not any(x in blob for x in ("india", "indian", "₹", "inr")):
        return False
    return True


def rss_candidates(source, domains):
    html = get(source)
    if not html:
        return []
    soup = BeautifulSoup(html, "html.parser")
    feeds = []

    for link in soup.find_all("link"):
        rel = " ".join(link.get("rel", [])).lower()
        typ = (link.get("type") or "").lower()
        href = link.get("href")
        if href and "alternate" in rel and any(x in typ for x in ("rss", "atom", "xml")):
            u = norm(href, source)
            if official(u, domains):
                feeds.append(u)

    rows = []
    for feed in dict.fromkeys(feeds):
        try:
            p = feedparser.parse(feed)
            for e in p.entries[:40]:
                u, title = e.get("link"), clean(e.get("title", ""))
                q = e.get("published_parsed") or e.get("updated_parsed")
                if not u or not official(u, domains):
                    continue
                dt = datetime(*q[:6], tzinfo=timezone.utc).astimezone(IST) if q else None
                rows.append((u, title, dt))
        except Exception:
            pass
    return rows


def html_candidates(source, domains):
    html = get(source)
    if not html:
        return []
    soup = BeautifulSoup(html, "html.parser")
    rows, seen = [], set()

    for a in soup.find_all("a", href=True):
        u = norm(a["href"], source)
        text = clean(a.get_text(" ", strip=True))
        blob = (text + " " + u).lower()

        if u in seen or not official(u, domains) or u.rstrip("/") == source.rstrip("/"):
            continue
        if not any(x in blob for x in SMART):
            continue

        seen.add(u)
        rows.append((u, text, None))
        if len(rows) >= 12:
            break
    return rows


def rendered_candidates(page, source, domains):
    try:
        page.goto(source, wait_until="domcontentloaded", timeout=35000)
        page.wait_for_timeout(1800)
        links = page.locator("a[href]").evaluate_all(
            """as => as.map(a => ({href:a.href||'',text:(a.innerText||a.textContent||'').trim()}))"""
        )
    except Exception:
        return []

    rows, seen = [], set()
    for item in links:
        u = norm(item.get("href", ""), source)
        text = clean(item.get("text", ""))
        blob = (text + " " + u).lower()

        if not u or u in seen or not official(u, domains) or u.rstrip("/") == source.rstrip("/"):
            continue
        if not any(x in blob for x in SMART):
            continue

        seen.add(u)
        rows.append((u, text, None))
        if len(rows) >= 12:
            break
    return rows


def google_news_candidates(brand, domains):
    # Discovery fallback only. It is never accepted as launch proof.
    query = f'"{brand}" smartphone India (launched OR launch OR introduced OR unveiled OR available)'
    url = f"https://news.google.com/rss/search?q={quote_plus(query)}&hl=en-IN&gl=IN&ceid=IN:en"
    rows = []
    try:
        p = feedparser.parse(url)
        for e in p.entries[:25]:
            u, title = e.get("link"), clean(e.get("title", ""))
            q = e.get("published_parsed") or e.get("updated_parsed")
            dt = datetime(*q[:6], tzinfo=timezone.utc).astimezone(IST) if q else None
            if u:
                rows.append((u, title, dt))
    except Exception:
        pass
    return rows


def discover(brand, cfg, source_page=None):
    """
    Fast discovery:
    1) official India RSS/Atom feeds
    2) lightweight links from official India pages
    3) Google News only as a discovery fallback
    Playwright is NOT used for discovery, keeping the daily job bounded.
    """
    rows = []

    for source in cfg["sources"]:
        rows += rss_candidates(source, cfg["domains"])
        rows += html_candidates(source, cfg["domains"])

    if len(rows) < 5:
        rows += google_news_candidates(brand, cfg["domains"])

    unique = {}
    for u, title, dt in rows:
        if u and u not in unique:
            unique[u] = (title, dt)

    result = [(u, t, d) for u, (t, d) in unique.items()]
    result.sort(
        key=lambda x: x[2] or datetime.min.replace(tzinfo=IST),
        reverse=True,
    )
    return result[:12]


def product_name(title, brand):
    x = clean(title)
    x = re.sub(r"^\s*(?:press release|update|quick read|photos)\s*[:\-]?\s*", "", x, flags=re.I)
    x = re.split(
        r"\s+(?:launched|launches|launch|introduced|unveiled|debuts|announces|announced)\b",
        x, maxsplit=1, flags=re.I
    )[0]
    x = re.sub(r"\s*[-–—:]\s*.*$", "", x).strip()
    if x.lower().startswith(brand.lower() + " "):
        x = x[len(brand):].strip()
    return x or "New Smartphone"


def description(html, name):
    soup = BeautifulSoup(html, "html.parser")
    for attrs in (
        {"property": "og:description"},
        {"name": "description"},
        {"name": "twitter:description"},
    ):
        m = soup.find("meta", attrs=attrs)
        if m and m.get("content") and len(m["content"]) >= 40:
            return clean(m["content"])[:2000]

    for p in soup.find_all("p"):
        txt = clean(p.get_text(" ", strip=True))
        if 80 <= len(txt) <= 2000:
            return txt
    return f"{name} smartphone launched in India."


def variants(page):
    out = []
    try:
        elements = page.locator("[data-model-code]").all()
    except Exception:
        elements = []

    for e in elements:
        try:
            text = clean(e.inner_text())
            c = re.search(r"(?:colour|color)\s*:\s*([^|]+)", text, re.I)
            q = re.search(r"(\d+\s*(?:GB|TB))\s*[|｜/]\s*(\d+\s*GB)", text, re.I)
            if c and q:
                value = (c.group(1).strip(), q.group(1).strip(), q.group(2).strip())
                if value not in out:
                    out.append(value)
        except Exception:
            pass

    if out:
        return out

    try:
        body = clean(page.locator("body").inner_text())
    except Exception:
        return []

    pairs = re.findall(r"(\d+\s*(?:GB|TB))\s*[|｜/]\s*(\d+\s*GB)", body, re.I)
    colors = []
    for m in re.finditer(r"(?:colour|color)\s*:\s*([A-Za-z][A-Za-z0-9 &'’\-]{2,40})", body, re.I):
        c = m.group(1).strip()
        if c.lower() not in [x.lower() for x in colors]:
            colors.append(c)

    return [(colors[0], pairs[0][0], pairs[0][1])] if len(pairs) == 1 and len(colors) == 1 else []


def images(page):
    try:
        data = page.locator("img").evaluate_all(
            """imgs => imgs.map(i => ({s:i.currentSrc||i.src||'',w:i.naturalWidth||0,h:i.naturalHeight||0}))"""
        )
    except Exception:
        return []

    out = []
    for x in data:
        u, low = x["s"], x["s"].lower()
        if (
            u and x["w"] >= 400 and x["h"] >= 400
            and not any(z in low for z in ("logo", "icon", "favicon", "swatch", "color_", "colour_", "flag", "avatar"))
            and u not in out
        ):
            out.append(u)
        if len(out) >= 4:
            break
    return out


def marketplace_support(page, brand, model):
    result = {"description": "", "images": []}
    q = quote_plus(f"{brand} {model} smartphone")

    for url in (
        f"https://www.amazon.in/s?k={q}",
        f"https://www.flipkart.com/search?q={q}",
    ):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=20000)
            page.wait_for_timeout(1500)

            if not result["description"]:
                text = clean(page.locator("body").inner_text())
                if len(text) >= 80:
                    result["description"] = text[:1000]

            data = page.locator("img").evaluate_all(
                """imgs => imgs.map(i => ({s:i.currentSrc||i.src||'',w:i.naturalWidth||0,h:i.naturalHeight||0,alt:i.alt||''}))"""
            )
            for x in data:
                u, low = x["s"], x["s"].lower()
                if (
                    u and x["w"] >= 300 and x["h"] >= 300
                    and not any(z in low for z in ("logo", "icon", "favicon", "banner", "sprite"))
                    and u not in result["images"]
                ):
                    result["images"].append(u)
                if len(result["images"]) >= 4:
                    break

            if result["description"] or result["images"]:
                break
        except Exception:
            pass

    return result


def best_product_url(article_url, html, domains):
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        u = norm(a["href"], article_url)
        if official(u, domains):
            blob = (clean(a.get_text(" ", strip=True)) + " " + u).lower()
            if any(x in blob for x in SMART):
                return u
    return article_url


def main():
    print("=" * 70)
    print("RETAILIX SMARTPHONE AUTOMATION - FINAL CLOUD VERSION")
    print("OFFICIAL INDIA VERIFICATION + AMAZON/FLIPKART SUPPORT")
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

    products = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent=HEAD["User-Agent"],
            locale="en-IN",
            timezone_id="Asia/Kolkata",
        )

        product_page = context.new_page()
        marketplace_page = context.new_page()

        try:
            for brand, cfg in BRANDS.items():
                print("\nChecking:", brand)
                candidates = discover(brand, cfg)

                candidates.sort(
                    key=lambda x: x[2] or datetime.min.replace(tzinfo=IST),
                    reverse=True,
                )

                found = False
                checked = 0

                for candidate_url, hint_title, candidate_dt in candidates:
                    if checked >= 12:
                        break
                    checked += 1

                    html = get(candidate_url)
                    if not html:
                        continue

                    title = title_of(html, hint_title)
                    dt = candidate_dt or article_date(html)

                    if not qualifying(title, html, brand, dt, start, end):
                        continue

                    print("  VERIFIED:", title)
                    print("  Date:", dt.isoformat())
                    print("  Official:", candidate_url)

                    product_url = best_product_url(candidate_url, html, cfg["domains"])

                    try:
                        product_page.goto(product_url, wait_until="domcontentloaded", timeout=30000)
                        product_page.wait_for_timeout(2500)

                        vs = variants(product_page)
                        if not vs:
                            print("  SKIPPED: exact Color + Storage + RAM mapping not verified.")
                            continue

                        name = product_name(title, brand)
                        desc = description(html, name)
                        img = images(product_page)

                        if not img or len(desc) < 80:
                            support = marketplace_support(marketplace_page, brand, name)
                            if not img:
                                img = support["images"]
                            if len(desc) < 80 and support["description"]:
                                desc = support["description"]

                        products.append({
                            "name": name,
                            "brand": brand,
                            "description": desc,
                            "variants": vs,
                            "images": img,
                        })

                        print("  Exact variants:", len(vs))
                        print("  Images:", len(img))
                        found = True
                        break

                    except Exception as exc:
                        print("  Product page skipped:", type(exc).__name__)

                if not found:
                    print("  No qualifying launch in previous 24 hours.")

        finally:
            for p in (product_page, marketplace_page):
                try:
                    p.close()
                except Exception:
                    pass
            browser.close()

    added = skipped = 0

    if products:
        backup = XLSX.with_name(
            XLSX.stem + "_backup_" + now().strftime("%Y%m%d_%H%M%S") + XLSX.suffix
        )
        shutil.copy2(XLSX, backup)
        print("\nBackup:", backup.name)

        for p in products:
            for color, storage, ram in p["variants"]:
                key = (
                    p["name"].lower(), p["brand"].lower(),
                    color.lower(), storage.lower(), ram.lower()
                )

                if key in existing:
                    skipped += 1
                    continue

                imgs = p["images"]
                ws.append([
                    p["name"], p["brand"], "Electronics", "Accessories", "piece",
                    "NO TAX APPLIED", "NODISCOUNT0", "1 Year Warranty",
                    p["description"][:500], p["description"][:2000],
                    "true", "false", color, storage, ram,
                    imgs[0] if len(imgs) > 0 else "", "true",
                    imgs[1] if len(imgs) > 1 else "",
                    imgs[2] if len(imgs) > 2 else "",
                    imgs[3] if len(imgs) > 3 else "",
                ])
                existing.add(key)
                added += 1

        wb.save(XLSX)
    else:
        print("\nNo qualifying smartphone launches. Excel was NOT modified.")

    print("\n" + "=" * 70)
    print("AUTOMATION COMPLETED")
    print("Rows added:", added)
    print("Rows skipped:", skipped)
    print("=" * 70)


if __name__ == "__main__":
    main()
