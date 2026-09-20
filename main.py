from __future__ import annotations
import re, shutil
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse, quote_plus
import requests, feedparser
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

BASE=Path(__file__).resolve().parent
XLSX=BASE/'product_import_template (1).xlsx'
IST=timezone(timedelta(hours=5,minutes=30))
HEAD={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/142.0 Safari/537.36','Accept-Language':'en-IN,en;q=0.9'}
S=requests.Session(); S.headers.update(HEAD)
BRANDS={
'APPLE':(['apple.com'],['https://www.apple.com/in/newsroom/archive/']),
'SAMSUNG':(['samsung.com'],['https://news.samsung.com/in/category/products','https://news.samsung.com/in/latest']),
'VIVO':(['vivo.com'],['https://vivonewsroom.in/']),
'OPPO':(['oppo.com'],['https://www.oppo.com/in/newsroom/','https://www.oppo.com/in/newsroom/all/']),
'ONEPLUS':(['oneplus.com'],['https://www.oneplus.com/in']),
'REALME':(['realme.com'],['https://www.realme.com/in/brand/newsroom']),
'MOTOROLA':(['motorola.in','motorola.com'],['https://www.motorola.in/']),
'INFINIX':(['infinixmobility.com','infinixmobiles.in'],['https://infinixmobiles.in/']),
'POCO':(['mi.com'],['https://www.mi.com/in/','https://www.mi.com/in/phone/']),
'REDMI':(['mi.com'],['https://www.mi.com/in/','https://www.mi.com/in/phone/'])}
SMART=('smartphone','iphone','galaxy','redmi','poco','mobile','phone','reno','find x','find n','nord','oneplus','narzo','realme','moto','edge','razr','zero','note')
LAUNCH=('launch','launched','launches','introduces','introduced','unveils','unveiled','debuts','debut','announces','available','goes on sale','sale begins','now available')
BAD=('upcoming','rumour','rumor','leak','leaked','expected','may launch','could launch','might launch','tipped','teased','teaser','concept','review','comparison','price hike','discount','offer','campaign','event')
HEADERS=['productName','brandName','categoryName','subcategoryName','unit_name','taxName','discountName','warrantyName','shortDescription','longDescription','is_featured','is_digital','variant_Color','variant_Storage','variant_RAM','image1','image1Primary','image2','image3','image4']

def now(): return datetime.now(timezone.utc).astimezone(IST)
def get(url,t=20):
    try:
        r=S.get(url,timeout=t,allow_redirects=True)
        return r.text if r.status_code==200 and 'text' in r.headers.get('content-type','') else ''
    except requests.RequestException: return ''
def norm(u,b): return urljoin(b,u).split('#')[0]
def official(u,ds):
    h=urlparse(u).netloc.lower().split(':')[0]
    return any(h==d or h.endswith('.'+d) for d in ds)

def feeds(src,ds):
    h=get(src)
    if not h:return []
    soup=BeautifulSoup(h,'html.parser'); out=[]
    for x in soup.find_all('link'):
        rel=' '.join(x.get('rel',[])).lower(); typ=(x.get('type') or '').lower(); href=x.get('href')
        if href and 'alternate' in rel and any(z in typ for z in ('rss','atom','xml')):
            u=norm(href,src)
            if official(u,ds): out.append(u)
    rows=[]
    for f in dict.fromkeys(out):
        try:
            p=feedparser.parse(f)
            for e in p.entries[:30]:
                u=e.get('link'); tt=e.get('title',''); q=e.get('published_parsed') or e.get('updated_parsed')
                if not u or not official(u,ds):continue
                dt=datetime(*q[:6],tzinfo=timezone.utc).astimezone(IST) if q else None
                rows.append((u,tt,dt))
        except Exception:pass
    return rows

def discover(src,ds):
    rows=feeds(src,ds)
    if rows:return rows[:40]
    h=get(src)
    if not h:return []
    soup=BeautifulSoup(h,'html.parser'); out=[]; seen=set()
    for a in soup.find_all('a',href=True):
        u=norm(a['href'],src); txt=' '.join(a.stripped_strings); blob=(txt+' '+u).lower()
        if u in seen or not official(u,ds):continue
        if not any(x in blob for x in SMART) or not any(x in blob for x in LAUNCH):continue
        seen.add(u); out.append((u,txt,None))
        if len(out)>=25:break
    return out

def article_date(h):
    soup=BeautifulSoup(h,'html.parser'); vals=[]
    for m in soup.find_all('meta'):
        k=(m.get('property') or m.get('name') or '').lower(); v=m.get('content')
        if v and any(x in k for x in ('published','datepublished','article:published','date')):vals.append(v)
    for t in soup.find_all('time'):
        vals += [x for x in (t.get('datetime'),t.get_text(' ',strip=True)) if x]
    for v in vals:
        try:
            d=datetime.fromisoformat(v.strip().replace('Z','+00:00'))
            return (d if d.tzinfo else d.replace(tzinfo=timezone.utc)).astimezone(IST)
        except: pass
    text=soup.get_text(' ',strip=True)
    for p,fmt in [(r'\b(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+2026)\b','%d %B %Y'),(r'\b((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+2026)\b','%B %d, %Y')]:
        m=re.search(p,text,re.I)
        if m:return datetime.strptime(m.group(1),fmt).replace(tzinfo=IST)
    return None

def qualifying(title,h,dt):
    blob=(title+' '+BeautifulSoup(h,'html.parser').get_text(' ',strip=True)).lower()
    if not any(x in blob for x in SMART) or not any(x in blob for x in LAUNCH):return False
    if any(x in title.lower() for x in BAD):return False
    return 'india' in blob or 'indian' in blob or '₹' in blob

def pname(title,brand):
    x=re.sub(r'\s+',' ',title).strip(); x=re.split(r'\s+(?:launched|launches|launch|introduced|unveiled|debuts)\b',x,maxsplit=1,flags=re.I)[0]; x=re.sub(r'\s*[-–—:]\s*.*$','',x).strip()
    p=brand.lower()
    if x.lower().startswith(p+' '):x=x[len(p):].strip()
    return x or 'New Smartphone'

def desc(h,name):
    s=BeautifulSoup(h,'html.parser')
    for a in ({'property':'og:description'},{'name':'description'},{'name':'twitter:description'}):
        m=s.find('meta',attrs=a)
        if m and m.get('content') and len(m['content'])>=40:return re.sub(r'\s+',' ',m['content']).strip()[:2000]
    return name+' smartphone launched in India.'

def variants(page):
    out=[]
    try: els=page.locator('[data-model-code]').all()
    except: els=[]
    for e in els:
        try:
            t=' '.join(e.inner_text().split()); c=re.search(r'(?:colour|color)\s*:\s*([^|]+)',t,re.I); q=re.search(r'(\d+\s*(?:GB|TB))\s*[|｜/]\s*(\d+\s*GB)',t,re.I)
            if c and q:
                v=(c.group(1).strip(),q.group(1).strip(),q.group(2).strip())
                if v not in out:out.append(v)
        except:pass
    if out:return out
    try:t=' '.join(page.locator('body').inner_text().split())
    except:return []
    pairs=re.findall(r'(\d+\s*(?:GB|TB))\s*[|｜/]\s*(\d+\s*GB)',t,re.I)
    colors=[]
    for m in re.finditer(r'(?:colour|color)\s*:\s*([A-Za-z][A-Za-z0-9 &\'\-]{2,40})',t,re.I):
        c=m.group(1).strip()
        if c.lower() not in [x.lower() for x in colors]:colors.append(c)
    return [(colors[0],pairs[0][0],pairs[0][1])] if len(pairs)==1 and len(colors)==1 else []

def images(page):
    try:data=page.locator('img').evaluate_all("imgs=>imgs.map(i=>({s:i.currentSrc||i.src||'',w:i.naturalWidth||0,h:i.naturalHeight||0}))")
    except:return []
    out=[]
    for x in data:
        u=x['s']; l=u.lower()
        if u and x['w']>=400 and x['h']>=400 and not any(z in l for z in ('logo','icon','favicon','swatch','color_','colour_','flag','avatar')) and u not in out:
            out.append(u)
            if len(out)>=4:break
    return out

def main():
    print('='*64);print('RETAILIX SMARTPHONE AUTOMATION - FINAL');print('OFFICIAL INDIA + AMAZON/FLIPKART SUPPORT');print('='*64)
    if not XLSX.exists():raise FileNotFoundError(XLSX)
    wb=load_workbook(XLSX); ws=wb['PRODUCT_DATA']; hdr=[ws.cell(1,c).value for c in range(1,21)]
    if hdr!=HEADERS:raise RuntimeError('RetailiX template headers do not match expected 20 columns.')
    cols={ws.cell(1,c).value:c for c in range(1,ws.max_column+1)}; existing=set()
    for r in ws.iter_rows(min_row=2,values_only=True):existing.add(tuple(str(r[cols[k]-1] or '').strip().lower() for k in ('productName','brandName','variant_Color','variant_Storage','variant_RAM')))
    end=now(); start=end-timedelta(hours=24); print(f'Window: {start.isoformat()} -> {end.isoformat()}');print('Existing combinations:',len(existing))
    products=[]
    with sync_playwright() as pw:
        browser=pw.chromium.launch(headless=True); ctx=browser.new_context(user_agent=HEAD['User-Agent'],locale='en-IN',timezone_id='Asia/Kolkata')
        for brand,(domains,sources) in BRANDS.items():
            print('\nChecking:',brand); cand=[]
            for src in sources:
                try:cand+=discover(src,domains)
                except:pass
            seen={}
            for u,t,d in cand:seen[u]=(t,d)
            found=False
            for u,(hint,d) in list(seen.items())[:40]:
                h=get(u)
                if not h:continue
                if d is None:d=article_date(h)
                if not d or not(start<=d<=end):continue
                soup=BeautifulSoup(h,'html.parser'); title=hint or ''
                if not title:
                    m=soup.find('meta',property='og:title'); title=m.get('content','') if m else (soup.title.get_text(' ',strip=True) if soup.title else '')
                if not qualifying(title,h,d):continue
                print('  VERIFIED:',title);print('  Official:',u)
                product_url=u
                for a in soup.find_all('a',href=True):
                    z=norm(a['href'],u); b=(' '.join(a.stripped_strings)+' '+z).lower()
                    if z!=u and official(z,domains) and any(x in b for x in SMART):product_url=z;break
                page=ctx.new_page()
                try:
                    page.goto(product_url,wait_until='domcontentloaded',timeout=45000);page.wait_for_timeout(2500)
                    vs=variants(page)
                    if not vs:print('  SKIPPED: exact Color + Storage + RAM mapping not verified.');page.close();break
                    products.append({'name':pname(title,brand),'brand':brand,'description':desc(h,pname(title,brand)),'variants':vs,'images':images(page)})
                    print('  Exact variants:',len(vs));found=True
                except Exception as e:print('  Product page skipped:',type(e).__name__)
                finally:page.close()
                if found:break
            if not found:print('  No qualifying launch in previous 24 hours.')
        browser.close()
    added=skipped=0
    if products:
        backup=XLSX.with_name(XLSX.stem+'_backup_'+now().strftime('%Y%m%d_%H%M%S')+XLSX.suffix);shutil.copy2(XLSX,backup);print('\nBackup:',backup.name)
        for p in products:
            for c,s,r in p['variants']:
                key=(p['name'].lower(),p['brand'].lower(),c.lower(),s.lower(),r.lower())
                if key in existing:skipped+=1;continue
                im=p['images'];ws.append([p['name'],p['brand'],'Electronics','Accessories','piece','NO TAX APPLIED','NODISCOUNT0','1 Year Warranty',p['description'][:500],p['description'][:2000],'true','false',c,s,r,im[0] if len(im)>0 else '','true',im[1] if len(im)>1 else '',im[2] if len(im)>2 else '',im[3] if len(im)>3 else '']);existing.add(key);added+=1
        wb.save(XLSX)
    else:print('\nNo qualifying smartphone launches. Excel was NOT modified.')
    print('\n'+'='*64);print('AUTOMATION COMPLETED');print('Rows added:',added);print('Rows skipped:',skipped);print('='*64)
if __name__=='__main__':main()
