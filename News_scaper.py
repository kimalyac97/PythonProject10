# Jupyter 단일 셀: 대한민국 에너지 뉴스 스크랩 (완전 랜덤 순번, HTML 미리보기 .txt 저장)
# ─ 요구사항 통합 ─
# - 수요일 기준 주차 시트명 (예: 25.10.4주차)
# - 날짜별 최대 5개, 최근 7일 → 총 35개 무작위 셔플 후 순번 부여
# - 한 줄 요약(본문→메타→제목 폴백)
# - 제외어: -배구, -학폭
# - 우선 가중치(계통, 그리드위즈, iDRS, 아이디알서비스, 전력수급기본계획, 분산에너지)
# - Google News 결과 파싱 + 언랩(unwrap) + URL 정규화 + 중복 제거
# - 한국 매체 판별(도메인/언어 휴리스틱)
# - 매체명 매핑 + 기자 직함 보존/부여
# - ★ HTML 미리보기는 .txt 파일로 저장(HTML 코드 그대로) ★

# 1) 패키지 설치
import sys, subprocess
from importlib.util import find_spec

def ensure(pkgs):
    to_install = [p for p in pkgs if find_spec(p.split("==")[0]) is None]
    if to_install:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", *to_install])

ensure([
    "requests", "beautifulsoup4", "lxml", "readability-lxml",
    "pandas", "openpyxl", "tqdm", "python-dateutil", "feedparser"
])

# 2) 임포트
import re, json, time, random, html
from datetime import datetime, timedelta, date
from urllib.parse import urlparse, parse_qs, quote_plus, urlencode, urlunparse
import requests
from bs4 import BeautifulSoup
from readability import Document
import pandas as pd
from tqdm import tqdm
import feedparser
from dateutil import parser as dateparser
from dateutil.tz import gettz

# 3) 설정
DEBUG = True
def log(*args):
    if DEBUG:
        print("[LOG]", *args)

ORIGINAL_URL = "https://www.google.com/search?lr=&cr=countryKR&sca_esv=f6d19e7077c59f8e&tbs=ctr:countryKR&q=%EC%88%98%EC%9A%94%EA%B4%80%EB%A6%AC+OR+%EA%B7%B8%EB%A6%AC%EB%93%9C%EC%9C%84%EC%A6%88+OR+DR+OR+%EC%A0%84%EB%A0%A5+OR+%ED%95%9C%EC%A0%84&tbm=nws&num=100"
PER_DAY = 5
DAYS = 7
KST = gettz("Asia/Seoul")
CAND_CAP = 40

PRIORITY_DOMAINS = {
    "etnews.com", "khan.co.kr", "electimes.com", "marketin.edaily.co.kr",
    "ekn.kr", "mk.co.kr", "energydaily.co.kr", "edaily.co.kr", "e2news.com",
}
PRIORITY_TERMS_WEIGHTS = {
    "아이디알서비스": 3, "idrs": 3, "iDRS": 3,
    "그리드위즈": 2, "전력수급기본계획": 2, "분산에너지": 2, "계통": 1,
}
EXCLUDE_TERMS = {"배구", "학폭"}

PUB_NAME_MAP = {
    "etnews.com": "전자신문","khan.co.kr": "경향신문","electimes.com": "전기신문",
    "marketin.edaily.co.kr": "이데일리 마켓인","ekn.kr": "에너지경제","mk.co.kr": "매일경제",
    "energydaily.co.kr": "에너지데일리","edaily.co.kr": "이데일리","e2news.com": "이투뉴스",
    "yonhapnews.co.kr": "연합뉴스","yna.co.kr": "연합뉴스","newsis.com": "뉴시스",
    "hankyung.com": "한국경제","chosun.com": "조선일보","hani.co.kr": "한겨레","joins.com": "중앙일보",
    "hankookilbo.com": "한국일보","seoul.co.kr": "서울신문","kmib.co.kr": "국민일보","munhwa.com": "문화일보",
    "ohmynews.com": "오마이뉴스","pressian.com": "프레시안","zdnet.co.kr": "지디넷코리아",
    "sbs.co.kr": "SBS","mbc.co.kr": "MBC","jtbc.co.kr": "JTBC","kbs.co.kr": "KBS",
    "busan.com": "부산일보","asiatoday.co.kr": "아시아투데이","naver.com": "네이버뉴스",
    "daum.net": "다음뉴스","news1.kr": "뉴스1","sisajournal-e.com": "시사저널e",
    "tongilnews.com": "통일신문","m-i.kr": "매일일보","newsspirit.kr": "뉴스스피릿",
    "h2news.kr":"이투뉴스","incheontoday.com": "인천투데이",
}

# 4) 유틸
def clean_text(s):
    if not s: return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def get_headers():
    uas = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36",
    ]
    ua = uas[int(time.time()) % len(uas)]
    return {"User-Agent": ua, "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"}

def get_soup(url, timeout=25, allow_redirects=True):
    headers = get_headers()
    for i in range(4):
        try:
            resp = requests.get(url, headers=headers, timeout=timeout, allow_redirects=allow_redirects)
            if resp.status_code == 200 and "https://www.google.com/sorry" not in resp.url:
                return BeautifulSoup(resp.text, "lxml"), resp
            time.sleep(1.2 * (i + 1))
        except requests.RequestException:
            time.sleep(1.2 * (i + 1))
    return None, None

def parse_query_from_original(url):
    try:
        q = parse_qs(urlparse(url).query).get("q", [""])[0]
        return q or ""
    except Exception:
        return ""

def compose_query(base_q: str) -> str:
    base_q = (base_q or "").strip()
    extra_terms = ["계통", "그리드위즈", "iDRS", "아이디알서비스", "전력수급기본계획", "분산에너지"]
    for t in extra_terms:
        if t not in base_q:
            base_q = (base_q + f" OR {t}") if base_q else t
    if "-배구" not in base_q: base_q += " -배구"
    if "-학폭" not in base_q: base_q += " -학폭"
    return base_q

def build_news_url_for_date(query, date_obj):
    md = date_obj.strftime("%m/%d/%Y")
    tbs = f"cdr:1,cd_min:{md},cd_max:{md},ctr:countryKR"
    q_enc = quote_plus(query)
    return f"https://www.google.com/search?tbm=nws&q={q_enc}&tbs={tbs}&num=100&hl=ko&lr=lang_ko&cr=countryKR&gl=KR"

# ─ Google News 언랩/정규화
def unwrap_google_news_link(link: str) -> str:
    try:
        u = urlparse(link)
        if "news.google." in (u.hostname or ""):
            q = parse_qs(u.query)
            real = q.get("url", [None])[0]
            if real: return real
        return link
    except Exception:
        return link

_TRACKING_PARAMS = {"utm_source","utm_medium","utm_campaign","utm_content","utm_term","fbclid","gclid","igshid","ref"}
def normalize_url_for_dedupe(u: str) -> str:
    try:
        p = urlparse(u)
        qs = parse_qs(p.query)
        clean = {k: v for k, v in qs.items() if k not in _TRACKING_PARAMS}
        return urlunparse(p._replace(query=urlencode({k: v[0] for k, v in clean.items()}, doseq=True)))
    except Exception:
        return u

def title_key(title: str) -> str:
    return re.sub(r"[\s\W]+", "", (title or "").lower())

# ─ Google HTML 파서
def _extract_card_title(card, a):
    head = card.select_one("div[role='heading']") or card.find("h3")
    if head:
        t = clean_text(head.get_text(" "))
        if t: return t
    aria = a.get("aria-label")
    if aria:
        t = clean_text(aria)
        if t: return t
    raw = a.get_text(" ")
    return clean_text(raw.splitlines()[0])

def parse_google_news_results(search_url):
    log("Google HTML 검색:", search_url)
    soup, resp = get_soup(search_url)
    if not soup:
        log("HTML 파싱 실패")
        return []
    items, seen = [], set()
    for card in soup.select("div.dbsr"):
        a = card.find("a")
        if not a or not a.get("href"): continue
        raw_title = _extract_card_title(card, a)
        raw_url = unwrap_google_news_link(a["href"])
        norm = (normalize_url_for_dedupe(raw_url), title_key(raw_title))
        if norm in seen: continue
        seen.add(norm)
        items.append({"title": raw_title, "url": raw_url})
    if not items:
        for a in soup.select("a.WlydOe, a.VDXfz"):
            link = a.get("href")
            if not link: continue
            head = a.select_one("div[role='heading'], h3")
            raw = head.get_text(" ") if head else a.get_text(" ")
            title = clean_text((raw or "").splitlines()[0])
            link = unwrap_google_news_link(link)
            norm = (normalize_url_for_dedupe(link), title_key(title))
            if norm in seen: continue
            seen.add(norm)
            items.append({"title": title, "url": link})
    log(f"HTML 결과 {len(items)}건")
    return [it for it in items if it["title"]]

# ─ RSS 대체 파서
def build_news_rss_url(query):
    q_enc = quote_plus(query)
    return f"https://news.google.com/rss/search?q={q_enc}&hl=ko&gl=KR&ceid=KR:ko"

def parse_google_news_results_rss(query, target_date):
    url = build_news_rss_url(query)
    log("Google RSS 검색:", url)
    feed = feedparser.parse(url)
    items, seen = [], set()
    for e in feed.entries:
        title = clean_text(getattr(e, 'title', ''))
        link = unwrap_google_news_link(getattr(e, 'link', ''))
        published = None
        if hasattr(e, 'published'):
            try: published = dateparser.parse(e.published)
            except Exception: published = None
        if not title or not link: continue
        if published:
            kst = published.astimezone(KST) if published.tzinfo else published
            if kst.date() != target_date: continue
        norm = (normalize_url_for_dedupe(link), title_key(title))
        if norm in seen: continue
        seen.add(norm)
        items.append({"title": title, "url": link})
    log(f"RSS 결과 {len(items)}건 (필터 후)")
    return items

# ─ 대한민국 판별
KR_TLDS = (".kr",)
KR_DOMAINS = {
    "naver.com","daum.net","nate.com","chosun.com","hani.co.kr","khan.co.kr","joins.com","hankookilbo.com",
    "seoul.co.kr","mk.co.kr","yonhapnews.co.kr","yna.co.kr","news1.kr","newspim.com","nocutnews.co.kr",
    "ohmynews.com","pressian.com","newsis.com","mbc.co.kr","sbs.co.kr","jtbc.co.kr","kbs.co.kr",
    "edaily.co.kr","etnews.com","zdnet.co.kr","asiatoday.co.kr","kmib.co.kr","munhwa.com","hankyung.com",
    "isplus.com","busan.com","e2news.com","electimes.com","energydaily.co.kr","ekn.kr",*PRIORITY_DOMAINS,
}
def domain_of(url):
    try:
        host = urlparse(url).hostname or ""
        return host.lower()
    except Exception:
        return ""
def host_matches_suffix(host, suffix_set):
    return any(host == s or host.endswith("." + s) or host.endswith(s) for s in suffix_set)
def looks_korean_by_meta_and_text(soup, html_text):
    html_tag = soup.find("html")
    if html_tag and (html_tag.get("lang") or html_tag.get("xml:lang")):
        lg = (html_tag.get("lang") or html_tag.get("xml:lang") or "").lower()
        if lg.startswith("ko"): return True
    og = soup.find("meta", {"property": "og:locale"})
    if og and og.get("content") and og["content"].lower().startswith("ko"): return True
    text = soup.get_text(" ", strip=True) or BeautifulSoup(html_text or "", "lxml").get_text(" ", strip=True)
    hangul = len(re.findall(r"[가-힣]", text))
    letters = len(re.findall(r"[A-Za-z가-힣]", text))
    return bool(hangul >= 40 and letters and (hangul / letters) >= 0.30)
def is_korean_source(final_url, soup, html_text):
    host = domain_of(final_url)
    if host.endswith(KR_TLDS) or host_matches_suffix(host, KR_DOMAINS): return True
    return looks_korean_by_meta_and_text(soup, html_text)
def is_priority_host(final_url):
    return host_matches_suffix(domain_of(final_url), PRIORITY_DOMAINS)

# ─ 요약/작성자
SENT_SPLIT_RE = re.compile(r"(?<=[.!?。！？])\s+|(?<=다\.|요\.)\s+")
def pick_meaningful_sentence(text: str) -> str:
    text = clean_text(text)
    if not text: return ""
    parts = re.split(SENT_SPLIT_RE, text)
    noise = ("무단전재","재배포","저작권","사진","기자","연합뉴스")
    for p in parts:
        s = clean_text(p)
        if not s: continue
        if any(n in s for n in noise) and len(s) < 30: continue
        if len(re.sub(r"[^가-힣A-Za-z0-9]", "", s)) < 4: continue
        return s
    return clean_text(parts[0]) if parts else ""
def shrink_to_chars(text: str, target: int = 30, hard_max: int = 36) -> str:
    t = clean_text(text)
    return t if len(t) <= target else t[:target].rstrip() + "…"
def concise_summary_from_text(text: str) -> str:
    return shrink_to_chars(pick_meaningful_sentence(text), target=30, hard_max=36)
def extract_main_text(soup, html_text):
    try:
        doc = Document(html_text)
        main_html = doc.summary()
        main_text = clean_text(BeautifulSoup(main_html, "lxml").get_text(" "))
        if len(main_text) >= 60: return main_text
    except Exception:
        pass
    for sel in [
        "article","div[itemprop='articleBody']",".article-body",".news_end",".article",
        "#newsct_article",".newsct_article",".art_txt",".article_view","#articleBodyContents",
        "#articeBody","#articleBody","#newsEndContents"
    ]:
        el = soup.select_one(sel)
        if el:
            txt = clean_text(el.get_text(" "))
            if len(txt) >= 60: return txt
    return clean_text(soup.get_text(" "))
def extract_summary(soup, html_text, title_fallback=""):
    main = extract_main_text(soup, html_text)
    if main:
        s = concise_summary_from_text(main)
        if s: return s
    og = soup.find("meta", property="og:description")
    if og and og.get("content"):
        s = concise_summary_from_text(og["content"])
        if s: return s
    md = soup.find("meta", attrs={"name":"description"})
    if md and md.get("content"):
        s = concise_summary_from_text(md["content"])
        if s: return s
    return shrink_to_chars(title_fallback or "", target=30)

_TITLES_RE = re.compile(r"(기자|특파원|논설위원|평론가|칼럼니스트|사진기자|에디터|부장|팀장|국장)")
_KO_NAME_RE = re.compile(r"([가-힣]{2,4})")
_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
def _strip_noise(txt):
    if not txt: return ""
    txt = _EMAIL_RE.sub("", txt)
    txt = re.sub(r"\(.*?\)|\[.*?\]|<.*?>"," ",txt)
    txt = re.sub(r"[-–—•·▶◇]|By|by|기자명?:?"," ",txt,flags=re.I)
    return clean_text(txt)
def extract_json_ld(soup):
    out=[]
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            block=json.loads(tag.string or "")
            out.extend(block if isinstance(block,list) else [block])
        except Exception:
            pass
    return out
def extract_authors_from_jsonld(jsonld):
    names=[]
    for b in (jsonld if isinstance(jsonld,list) else [jsonld]):
        for key in ("author","creator"):
            author=b.get(key)
            if not author: continue
            vals = author if isinstance(author,list) else [author]
            for a in vals:
                if isinstance(a,dict) and a.get("name"): names.append(a["name"])
                elif isinstance(a,str): names.append(a)
    names=[_strip_noise(n) for n in names if n]
    return [n for n in names if n]
def extract_author_meta_and_dom(soup):
    cands=[]
    for sel in [
        ('meta', {'name':'author'}),('meta', {'property':'article:author'}),
        ('meta', {'name':'parsely-author'}),('meta', {'name':'byl'}),
    ]:
        tag=soup.find(*sel)
        if tag and tag.get("content"): cands.append(tag["content"])
    for sel in [
        "a[rel='author']","[itemprop='author']","[itemprop='author'] [itemprop='name']",
        "address.byline","p.byline","span.byline","div.byline",
        "span[class*=author]","div[class*=author]","p[class*=author]",
        "span[class*=writer]","div[class*=writer]","p[class*=writer]",
        "span[class*=reporter]","div[class*=reporter]","p[class*=reporter]",
        "span[class*=journalist]","div[class*=journalist]","p[class*=journalist]",
        "span.article_writer","div.article_writer","em.article_writer",
        ".info_view .writer",".press_writer","#news_writer","#author",
        "strong.name","span.name",
    ]:
        for el in soup.select(sel):
            txt=clean_text(el.get_text(" "))
            if txt: cands.append(txt)
    for el in soup.find_all(True):
        for attr in el.attrs:
            if "author" in attr or "writer" in attr or "reporter" in attr:
                val=el.get(attr)
                if isinstance(val,str) and len(val)<=40: cands.append(val)
    cleaned=[]
    for c in cands:
        c=_strip_noise(c)
        for p in re.split(r"[,/;•·]| 그리고 | 및 | and ", c):
            p=clean_text(p)
            if not p: continue
            name_m=_KO_NAME_RE.search(p)
            if not name_m: continue
            name=name_m.group(1)
            title_m=_TITLES_RE.search(p)
            cleaned.append(f"{name} {title_m.group(1)}" if title_m else name)
    def _score(s):
        sc=0
        if _TITLES_RE.search(s): sc+=2
        if len(s)<=20: sc+=1
        if re.search(r"[가-힣]{2,4}", s): sc+=1
        return sc
    uniq=list(dict.fromkeys([x for x in cleaned if x]))
    uniq.sort(key=lambda x:_score(x), reverse=True)
    return uniq[:2]
def try_parse_date(text):
    try: return dateparser.parse(text)
    except: return None

def map_publisher_name(host):
    host=host.lower()
    for key,name in PUB_NAME_MAP.items():
        if host==key or host.endswith("."+key) or host.endswith(key):
            return name
    parts=host.split(".")
    return ".".join(parts[-2:]) if len(parts)>=2 else host
def normalize_author_with_title(author: str) -> str:
    s=clean_text(author)
    if not s: return ""
    m=re.search(r"([가-힣]{2,4})\s*(기자|특파원|논설위원|평론가|칼럼니스트|사진기자|에디터|부장|팀장|국장)?", s)
    if m:
        name=m.group(1); title=m.group(2)
        return f"{name} {title}" if title else f"{name} 기자"
    return s if s.endswith("기자") else f"{s} 기자"
def build_reporter_cell(host, authors):
    pub_name=map_publisher_name(host.replace("www.",""))
    if authors:
        author_fmt=normalize_author_with_title(authors[0])
        return f"{pub_name} {author_fmt}"
    return pub_name

# ─ 수요일 기준 시트명
def wednesday_based_week_info(d: date):
    delta=(d.weekday()-2)%7
    anchor=d - timedelta(days=delta)
    y,m=anchor.year, anchor.month
    first=date(y,m,1)
    days_to_wed=(2-first.weekday())%7
    first_wed=first + timedelta(days=days_to_wed)
    week_no=1 + (anchor-first_wed).days//7
    return y,m,week_no
def week_sheet_name_wed_kst(now_dt):
    d=now_dt.date()
    y,m,w=wednesday_based_week_info(d)
    return f"{str(y)[-2:]}.{m:02d}.{w}주차"

# ─ 핵심: 기사 상세 (★ 언랩 복원: 원문 본문을 파싱하도록)
def fetch_article_details(url: str) -> dict:
    try:
        url = unwrap_google_news_link(url)  # ← 원문 URL로 언랩 (요청 사항)
    except Exception:
        pass
    soup, resp = get_soup(url, allow_redirects=True)
    if not soup or not resp:
        return {"final_url": url, "soup": None, "html": "", "authors": [], "published": None}
    final_url = resp.url
    html_text = resp.text

    authors=[]
    try:
        jsonlds=extract_json_ld(soup)
        authors=extract_authors_from_jsonld(jsonlds)
    except Exception:
        pass
    if not authors:
        authors=extract_author_meta_and_dom(soup)

    published=None
    try:
        for b in (jsonlds if isinstance(jsonlds,list) else []):
            for key in ("datePublished","dateCreated","uploadDate"):
                if isinstance(b,dict) and b.get(key):
                    dt=try_parse_date(b.get(key))
                    if dt: published=dt; break
            if published: break
    except Exception:
        pass
    if not published:
        meta=soup.find("meta", {"property":"article:published_time"}) or soup.find("meta", {"name":"pubdate"})
        if meta and meta.get("content"):
            published=try_parse_date(meta["content"])
    if not published:
        ttag=soup.find("time")
        if ttag and ttag.get("datetime"):
            published=try_parse_date(ttag.get("datetime"))

    return {"final_url": final_url, "soup": soup, "html": html_text, "authors": authors, "published": published}

# ─ 실행
base_query = parse_query_from_original(ORIGINAL_URL)
query = compose_query(base_query)
today_kst = datetime.now(tz=KST)
date_range = [today_kst.date() - timedelta(days=i) for i in range(DAYS)]

rows_all = []  # ["", None, title, url, summary, 기자, 일자]
for d in date_range:
    day_url = build_news_url_for_date(query, d)
    candidates = parse_google_news_results(day_url)
    if not candidates:
        candidates = parse_google_news_results_rss(query, d)
    log(f"{d} 후보 {len(candidates)}건")
    if CAND_CAP:
        candidates = candidates[:CAND_CAP]

    detailed=[]
    for idx, item in enumerate(candidates):
        if any(x in (item.get("title","")+item.get("url","")) for x in EXCLUDE_TERMS):
            continue
        det = fetch_article_details(item["url"])
        if det.get("soup") is None or not is_korean_source(det["final_url"], det["soup"], det["html"]):
            continue
        summary = extract_summary(det["soup"], det["html"], title_fallback=item["title"])
        if any(x in summary for x in EXCLUDE_TERMS):
            continue
        host = domain_of(det["final_url"]) or domain_of(item["url"]) or ""
        prio_score = (20 if is_priority_host(det["final_url"]) else 0) \
                   + sum(w for term, w in PRIORITY_TERMS_WEIGHTS.items() if term.lower() in (item["title"]+" "+summary+" "+det["final_url"]).lower())
        published = det.get("published")
        pub_date = (published.astimezone(KST).date() if published and published.tzinfo else (published.date() if published else d))
        detailed.append({
            "rank": idx, "prio_score": prio_score,
            "title": item["title"], "url": det["final_url"] or item["url"],
            "summary": clean_text(summary), "authors_list": det.get("authors", []),
            "published": pub_date, "host": host,
        })

    detailed.sort(key=lambda x: (-x["prio_score"], x["rank"]))
    picked = detailed[:PER_DAY]
    log(f"{d} 최종 선별 {len(picked)}건")
    for r in picked:
        reporter_cell = build_reporter_cell(r["host"], r["authors_list"])
        rows_all.append(["", None, r["title"], r["url"], r["summary"], reporter_cell, r["published"].strftime("%Y-%m-%d")])

# 5) 완전 랜덤 셔플 → 순번 부여
random.shuffle(rows_all)
rows = rows_all[:DAYS*PER_DAY]
for idx, row in enumerate(rows, start=1):
    row[1] = idx

# 6) 엑셀 저장
header_row = ["", "순번", "타이틀", "링크", "세부내용", "기자", "일자"]
data_for_excel = [[""] * 7, header_row] + rows
df_out = pd.DataFrame(data_for_excel)
sheet_name = week_sheet_name_wed_kst(today_kst)
stamp = today_kst.strftime("%Y%m%d_%H%M%S")
EXCEL_PATH = f"에너지뉴스_{stamp}.xlsx"

with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
    df_out.to_excel(writer, index=False, header=False, sheet_name=sheet_name)

# 7) HTML 미리보기(코드 자체)를 txt로 저장
def build_html_from_rows(rows, sheet_name):
    def html_br(n=1): return "<br>" * int(n)
    def html_p(txt, small=False, bold=False):
        t = html.escape(str(txt)).replace("\n","<br>")
        if bold: t = f"<b>{t}</b>"
        if small: t = f'<span style="font-size:90%">{t}</span>'
        return f'<p align="left">{t}</p>'
    def week_label_from_sheet(sheet_name):
        m = re.match(r"(\d{2})\.(\d{2})\.(\d+)주차", sheet_name)
        return f"{m.group(1)}년 {m.group(2)}월 {m.group(3)}주차" if m else sheet_name

    parts = []
    parts.append('<div align="">')
    parts.append(html_p("안녕하세요."))
    parts.append(html_br(1))
    parts.append(html_p("아이디알서비스 입니다."))
    parts.append(html_br(1))
    parts.append(html_p(f"{week_label_from_sheet(sheet_name)} 에너지 뉴스 모음 입니다."))
    parts.append(html_br(5))

    for idx, r in enumerate(rows, start=1):
        title, url, summary, reporter, d = r[2], r[3], r[4], r[5], r[6]
        parts.append(html_p(f"{idx}.\u00A0{title}", bold=True))
        parts.append(html_br(2))
        if url:
            esc = html.escape(url, quote=True)
            parts.append(f'<p align="left"><a href="{esc}" target="_blank" rel="noopener noreferrer">기사원문</a></p>')
            parts.append(html_br(2))
        if summary: parts.append(html_p(summary))
        if reporter: parts.append(html_p(reporter, small=True, bold=True))
        if d: parts.append(html_p(d, small=True, bold=True))
        parts.append(html_br(6))

    parts.append("</div>")
    return "".join(parts)

HTML_TXT_PATH = f"에너지뉴스_{stamp}.txt"   # ← txt로 저장 (내용은 HTML 코드)
with open(HTML_TXT_PATH, "w", encoding="utf-8") as f:
    f.write(build_html_from_rows(rows, sheet_name))

print("\n=== 저장 완료 ===")
print(f"- 엑셀 파일: {EXCEL_PATH} (시트명: {sheet_name})")
print(f"- HTML 코드(txt): {HTML_TXT_PATH}")
print(f"- 총 기사 수: {len(rows)} (최대 {DAYS*PER_DAY})")
