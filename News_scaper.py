# ===== 대한민국 에너지 뉴스 스크랩 (Streamlit GUI, Tk 기능 이식) =====
# - 체크박스 우선 키워드 / 직접 입력
# - 저장 경로/파일명 입력(서버 로컬 저장 + 다운로드)
# - 진행도 + 로그 스트림
# - 전체 HTML 미리보기(렌더) + HTML 코드 미리보기(txt) + 다운로드
# - 제목 클릭 → 모달 팝업에서 해당 기사 HTML 미리보기
# - Google 뉴스 언랩 복원으로 원문 본문 파싱

import os, io, re, json, time, random, html as pyhtml
from datetime import datetime, timedelta, date
from urllib.parse import urlparse, parse_qs, quote_plus, urlencode, urlunparse

import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from readability import Document
import feedparser
from dateutil import parser as dateparser
from dateutil.tz import gettz

# -------------------- 전역 설정 --------------------
KST = gettz("Asia/Seoul")
DEBUG = False
DEFAULT_ORIGINAL_URL = (
    "https://www.google.com/search?lr=&cr=countryKR&sca_esv=f6d19e7077c59f8e"
    "&tbs=ctr:countryKR&q=%EC%88%98%EC%9A%94%EA%B4%80%EB%A6%AC+OR+%EA%B7%B8%EB%A6%AC%EB%93%9C%EC%9C%84%EC%A6%88+OR+DR+OR+%EC%A0%84%EB%A0%A5+OR+%ED%95%9C%EC%A0%84&tbm=nws&num=100"
)

BASE_PRIORITY_TERMS = [
    "아이디알서비스", "iDRS", "idrs", "그리드위즈", "전력수급기본계획", "분산에너지", "계통"
]
EXCLUDE_TERMS = {"배구", "학폭"}
PRIORITY_DOMAINS = {
    "etnews.com", "khan.co.kr", "electimes.com", "marketin.edaily.co.kr",
    "ekn.kr", "mk.co.kr", "energydaily.co.kr", "edaily.co.kr", "e2news.com",
}
PRIORITY_TERMS_WEIGHTS = {
    "아이디알서비스": 3, "idrs": 3, "iDRS": 3,
    "그리드위즈": 2, "전력수급기본계획": 2, "분산에너지": 2, "계통": 1,
}
PUB_NAME_MAP = {
    "etnews.com":"전자신문","khan.co.kr":"경향신문","electimes.com":"전기신문","marketin.edaily.co.kr":"이데일리 마켓인",
    "ekn.kr":"에너지경제","mk.co.kr":"매일경제","energydaily.co.kr":"에너지데일리","edaily.co.kr":"이데일리","e2news.com":"이투뉴스",
    "yonhapnews.co.kr":"연합뉴스","yna.co.kr":"연합뉴스","newsis.com":"뉴시스","hankyung.com":"한국경제","chosun.com":"조선일보",
    "hani.co.kr":"한겨레","joins.com":"중앙일보","hankookilbo.com":"한국일보","seoul.co.kr":"서울신문","kmib.co.kr":"국민일보",
    "munhwa.com":"문화일보","ohmynews.com":"오마이뉴스","pressian.com":"프레시안","zdnet.co.kr":"지디넷코리아",
    "sbs.co.kr":"SBS","mbc.co.kr":"MBC","jtbc.co.kr":"JTBC","kbs.co.kr":"KBS","busan.com":"부산일보","asiatoday.co.kr":"아시아투데이",
    "naver.com":"네이버뉴스","daum.net":"다음뉴스","news1.kr":"뉴스1","sisajournal-e.com":"시사저널e",
    "tongilnews.com":"통일신문","m-i.kr":"매일일보","newsspirit.kr":"뉴스스피릿","h2news.kr":"이투뉴스","incheontoday.com":"인천투데이",
}
_TRACKING_PARAMS = {"utm_source","utm_medium","utm_campaign","utm_content","utm_term","fbclid","gclid","igshid","ref"}
SENT_SPLIT_RE = re.compile(r"(?<=[.!?。！？])\s+|(?<=다\.|요\.)\s+")

# -------------------- Streamlit 페이지 설정 --------------------
st.set_page_config(page_title="대한민국 에너지 뉴스 스크랩", layout="wide")
st.title("대한민국 에너지 뉴스 스크랩 (Streamlit GUI 확장판)")

# -------------------- 로그 유틸 --------------------
if "logs" not in st.session_state:
    st.session_state.logs = []

def log(msg):
    st.session_state.logs.append(msg)
    if len(st.session_state.logs) > 1000:
        st.session_state.logs = st.session_state.logs[-1000:]

def get_headers():
    uas = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36",
    ]
    ua = uas[int(time.time()) % len(uas)]
    return {"User-Agent": ua, "Accept-Language":"ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"}

def clean_text(s):
    return re.sub(r"\s+"," ",str(s)).strip() if s is not None else ""

def get_soup(url, timeout=25, allow_redirects=True):
    for i in range(4):
        try:
            r = requests.get(url, headers=get_headers(), timeout=timeout, allow_redirects=allow_redirects)
            if r.status_code == 200 and "https://www.google.com/sorry" not in r.url:
                return BeautifulSoup(r.text, "lxml"), r
            time.sleep(1.0*(i+1))
        except requests.RequestException:
            time.sleep(1.0*(i+1))
    return None, None

def domain_of(url):
    try:
        host = urlparse(url).hostname or ""
        return host.lower()
    except Exception:
        return ""

def host_matches_suffix(host, suffix_set):
    return any(host == s or host.endswith("."+s) or host.endswith(s) for s in suffix_set)

def parse_query_from_original(url):
    try:
        return parse_qs(urlparse(url).query).get("q",[""])[0] or ""
    except Exception:
        return ""

def compose_query(base_q: str, selected_terms, custom_terms) -> str:
    base_q = (base_q or "").strip()
    extra_terms = list(selected_terms or BASE_PRIORITY_TERMS)
    for t in (custom_terms or []):
        if t and t not in extra_terms:
            extra_terms.append(t)
    for t in extra_terms:
        if t not in base_q:
            base_q = (base_q + f" OR {t}") if base_q else t
    if "-배구" not in base_q: base_q += " -배구"
    if "-학폭" not in base_q: base_q += " -학폭"
    return base_q

def build_news_url_for_date(query, date_obj):
    md = date_obj.strftime("%m/%d/%Y")
    tbs = f"cdr:1,cd_min:{md},cd_max:{md},ctr:countryKR"
    return f"https://www.google.com/search?tbm=nws&q={quote_plus(query)}&tbs={tbs}&num=100&hl=ko&lr=lang_ko&cr=countryKR&gl=KR"

def unwrap_google_news_link(link: str) -> str:
    try:
        u = urlparse(link)
        if "news.google." in (u.hostname or ""):
            real = parse_qs(u.query).get("url", [None])[0]
            if real: return real
        return link
    except Exception:
        return link

def normalize_url_for_dedupe(u: str) -> str:
    try:
        p = urlparse(u); qs = parse_qs(p.query)
        clean = {k:v for k,v in qs.items() if k not in _TRACKING_PARAMS}
        return urlunparse(p._replace(query=urlencode({k:v[0] for k,v in clean.items()}, doseq=True)))
    except Exception:
        return u

def title_key(title: str) -> str:
    return re.sub(r"[\s\W]+","", (title or "").lower())

def _extract_card_title(card, a):
    head = card.select_one("div[role='heading']") or card.find("h3")
    if head:
        t = clean_text(head.get_text(" "))
        if t: return t
    aria = a.get("aria-label")
    if aria:
        t = clean_text(aria)
        if t: return t
    return clean_text((a.get_text(" ") or "").splitlines()[0])

def parse_google_news_results(search_url):
    log(f"[검색] {search_url}")
    soup, _ = get_soup(search_url)
    if not soup:
        log(" - HTML 파싱 실패")
        return []
    items, seen = [], set()
    for card in soup.select("div.dbsr"):
        a = card.find("a")
        if not a or not a.get("href"): continue
        raw_title = _extract_card_title(card, a)
        raw_url = unwrap_google_news_link(a["href"])  # 원문 언랩
        norm = (normalize_url_for_dedupe(raw_url), title_key(raw_title))
        if norm in seen: continue
        seen.add(norm); items.append({"title": raw_title, "url": raw_url})
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
            seen.add(norm); items.append({"title": title, "url": link})
    log(f" - HTML 결과 {len(items)}건")
    return [it for it in items if it["title"]]

def build_news_rss_url(query):
    return f"https://news.google.com/rss/search?q={quote_plus(query)}&hl=ko&gl=KR&ceid=KR:ko"

def parse_google_news_results_rss(query, target_date):
    url = build_news_rss_url(query)
    log(f"[RSS] {url}")
    feed = feedparser.parse(url)
    items, seen = [], set()
    for e in feed.entries:
        title = clean_text(getattr(e,'title','')); link  = unwrap_google_news_link(getattr(e,'link',''))
        published = dateparser.parse(getattr(e,'published','')) if hasattr(e,'published') else None
        if not title or not link: continue
        if published:
            kst = published.astimezone(KST) if published.tzinfo else published
            if kst.date() != target_date: continue
        norm = (normalize_url_for_dedupe(link), title_key(title))
        if norm in seen: continue
        seen.add(norm); items.append({"title": title, "url": link})
    log(f" - RSS 결과 {len(items)}건(필터 후)")
    return items

KR_TLDS = (".kr",)
KR_DOMAINS = {"naver.com","daum.net","nate.com","chosun.com","hani.co.kr","khan.co.kr","joins.com",
              "hankookilbo.com","seoul.co.kr","mk.co.kr","yonhapnews.co.kr","yna.co.kr","news1.kr",
              "newspim.com","nocutnews.co.kr","ohmynews.com","pressian.com","newsis.com","mbc.co.kr",
              "sbs.co.kr","jtbc.co.kr","kbs.co.kr","edaily.co.kr","etnews.com","zdnet.co.kr","asiatoday.co.kr",
              "kmib.co.kr","munhwa.com","hankyung.com","isplus.com","busan.com","e2news.com","electimes.com",
              "energydaily.co.kr","ekn.kr", *PRIORITY_DOMAINS}

def looks_korean_by_meta_and_text(soup, html_text):
    html_tag = soup.find("html")
    if html_tag and (html_tag.get("lang") or html_tag.get("xml:lang")):
        if (html_tag.get("lang") or html_tag.get("xml:lang") or "").lower().startswith("ko"): return True
    og = soup.find("meta", {"property":"og:locale"})
    if og and og.get("content") and og["content"].lower().startswith("ko"): return True
    text = soup.get_text(" ", strip=True) or BeautifulSoup(html_text or "", "lxml").get_text(" ", strip=True)
    hangul = len(re.findall(r"[가-힣]", text)); letters = len(re.findall(r"[A-Za-z가-힣]", text))
    return bool(hangul >= 40 and letters and (hangul/letters) >= 0.30)

def is_korean_source(final_url, soup, html_text):
    host = domain_of(final_url)
    if host.endswith(KR_TLDS) or host_matches_suffix(host, KR_DOMAINS): return True
    return looks_korean_by_meta_and_text(soup, html_text)

def is_priority_host(final_url): return host_matches_suffix(domain_of(final_url), PRIORITY_DOMAINS)

def pick_meaningful_sentence(text: str) -> str:
    text = clean_text(text); 
    if not text: return ""
    parts = re.split(SENT_SPLIT_RE, text)
    for p in parts:
        s = clean_text(p)
        if not s: continue
        if any(n in s for n in ("무단전재","재배포","저작권","사진","기자","연합뉴스")) and len(s) < 30: continue
        if len(re.sub(r"[^가-힣A-Za-z0-9]","",s)) < 4: continue
        return s
    return clean_text(parts[0]) if parts else ""

def concise_summary_from_text(text: str) -> str:
    s = pick_meaningful_sentence(text)
    return s if len(s) <= 30 else s[:30].rstrip()+"…"

def extract_main_text(soup, html_text):
    try:
        doc = Document(html_text); main_html = doc.summary()
        main_text = clean_text(BeautifulSoup(main_html,"lxml").get_text(" "))
        if len(main_text) >= 60: return main_text
    except Exception: pass
    for sel in ["article","div[itemprop='articleBody']", ".article-body",".news_end",".article",
                "#newsct_article",".newsct_article",".art_txt",".article_view","#articleBodyContents",
                "#articeBody","#articleBody","#newsEndContents"]:
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
    tf = title_fallback or ""
    return tf[:30] + ("…" if len(tf) > 30 else "")

_TITLES_RE = re.compile(r"(기자|특파원|논설위원|평론가|칼럼니스트|사진기자|에디터|부장|팀장|국장)")
_KO_NAME_RE = re.compile(r"([가-힣]{2,4})")
_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

def _strip_noise(txt):
    if not txt: return ""
    txt = _EMAIL_RE.sub("", txt)
    txt = re.sub(r"\(.*?\)|\[.*?\]|<.*?>"," ", txt)
    txt = re.sub(r"[-–—•·▶◇]|By|by|기자명?:?"," ", txt, flags=re.I)
    return clean_text(txt)

def extract_json_ld(soup):
    out=[]
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            block=json.loads(tag.string or "")
            out.extend(block if isinstance(block,list) else [block])
        except Exception: pass
    return out

def extract_authors_from_jsonld(jsonld):
    names=[]
    for b in (jsonld if isinstance(jsonld,list) else [jsonld]):
        for key in ("author","creator"):
            author=b.get(key); 
            if not author: continue
            vals = author if isinstance(author,list) else [author]
            for a in vals:
                if isinstance(a,dict) and a.get("name"): names.append(a["name"])
                elif isinstance(a,str): names.append(a)
    names=[_strip_noise(n) for n in names if n]
    return [n for n in names if n]

def extract_author_meta_and_dom(soup):
    cands=[]
    for sel in [('meta',{'name':'author'}),('meta',{'property':'article:author'}),
                ('meta',{'name':'parsely-author'}),('meta',{'name':'byl'})]:
        tag = soup.find(*sel)
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
            t = clean_text(el.get_text(" "))
            if t: cands.append(t)
    for el in soup.find_all(True):
        for attr in el.attrs:
            if "author" in attr or "writer" in attr or "reporter" in attr:
                val = el.get(attr)
                if isinstance(val,str) and len(val)<=40: cands.append(val)
    cleaned=[]
    for c in cands:
        c=_strip_noise(c)
        for p in re.split(r"[,/;•·]| 그리고 | 및 | and ", c):
            p = clean_text(p)
            if not p: continue
            m=_KO_NAME_RE.search(p)
            if not m: continue
            name=m.group(1); t=_TITLES_RE.search(p)
            cleaned.append(f"{name} {t.group(1)}" if t else name)
    uniq=list(dict.fromkeys(cleaned))
    uniq.sort(key=lambda s:(bool(_TITLES_RE.search(s)),-len(s)), reverse=True)
    return uniq[:2]

def try_parse_date(text):
    try: return dateparser.parse(text)
    except: return None

def map_publisher_name(host):
    host=host.lower()
    for key,name in PUB_NAME_MAP.items():
        if host==key or host.endswith("."+key) or host.endswith(key): return name
    parts=host.split(".")
    return ".".join(parts[-2:]) if len(parts)>=2 else host

def normalize_author_with_title(author: str) -> str:
    s=clean_text(author)
    m=re.search(r"([가-힣]{2,4})\s*(기자|특파원|논설위원|평론가|칼럼니스트|사진기자|에디터|부장|팀장|국장)?", s)
    if m:
        name=m.group(1); title=m.group(2)
        return f"{name} {title}" if title else f"{name} 기자"
    return s if s.endswith("기자") else (s+" 기자" if s else "")

def build_reporter_cell(host, authors):
    pub = map_publisher_name(host.replace("www.",""))
    if authors:
        return f"{pub} {normalize_author_with_title(authors[0])}"
    return pub

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

def contains_excluded(text:str)->bool:
    if not text: return False
    return any(x in str(text) for x in EXCLUDE_TERMS)

def has_priority_term(title:str, summary:str, url:str)->int:
    blob = " ".join([title or "", summary or "", url or ""]).lower()
    sc=0
    for term,w in PRIORITY_TERMS_WEIGHTS.items():
        if term.lower() in blob: sc+=w
    return sc

def fetch_article_details(url:str)->dict:
    # ★ 원문으로 언랩
    try: url = unwrap_google_news_link(url)
    except: pass
    soup, resp = get_soup(url, allow_redirects=True)
    if not soup or not resp:
        return {"final_url": url, "soup": None, "html":"", "authors":[], "published":None}
    final_url = resp.url; html_text = resp.text

    authors=[]
    try:
        jsonlds = extract_json_ld(soup)
        authors = extract_authors_from_jsonld(jsonlds)
    except: pass
    if not authors: authors = extract_author_meta_and_dom(soup)

    published=None
    try:
        for b in (jsonlds if isinstance(jsonlds,list) else []):
            for key in ("datePublished","dateCreated","uploadDate"):
                if isinstance(b,dict) and b.get(key):
                    dt = try_parse_date(b.get(key))
                    if dt: published = dt; break
            if published: break
    except: pass
    if not published:
        meta = soup.find("meta", {"property":"article:published_time"}) or soup.find("meta", {"name":"pubdate"})
        if meta and meta.get("content"): published = try_parse_date(meta["content"])
    if not published:
        ttag = soup.find("time")
        if ttag and ttag.get("datetime"): published = try_parse_date(ttag.get("datetime"))

    return {"final_url": final_url, "soup": soup, "html": html_text, "authors": authors, "published": published}

def parse_one_day(query, d, cand_cap, status=None, progress=None):
    items = parse_google_news_results(build_news_url_for_date(query, d))
    if not items:
        items = parse_google_news_results_rss(query, d)
    if cand_cap:
        items = items[:cand_cap]
    detailed=[]
    total=len(items)
    for idx, item in enumerate(items):
        if status: status.info(f"[파싱] {d} {idx+1}/{total}")
        if progress and total: progress.progress(min((idx+1)/total,1.0))
        if contains_excluded(item.get("title","")) or contains_excluded(item.get("url","")): 
            continue
        det = fetch_article_details(item["url"])
        if det.get("soup") is None or not is_korean_source(det["final_url"], det["soup"], det["html"]):
            continue
        summary = extract_summary(det["soup"], det["html"], title_fallback=item["title"])
        if contains_excluded(summary):
            continue
        host = domain_of(det["final_url"]) or domain_of(item["url"]) or ""
        prio = (20 if is_priority_host(det["final_url"]) else 0) + has_priority_term(item["title"], summary, det.get("final_url",""))
        published = det.get("published")
        pub_date = (published.astimezone(KST).date() if published and published.tzinfo else (published.date() if published else d))
        detailed.append({
            "rank": idx, "prio_score": prio, "title": item["title"],
            "url": det["final_url"] or item["url"], "summary": clean_text(summary),
            "authors_list": det.get("authors", []), "published": pub_date, "host": host,
        })
    return detailed

def run_pipeline(selected_terms, custom_terms, per_day=5, days=7, cand_cap=40, status=None, progress=None):
    base_q = parse_query_from_original(DEFAULT_ORIGINAL_URL)
    query = compose_query(base_q, selected_terms, custom_terms)

    today_kst = datetime.now(tz=KST)
    date_list = [today_kst.date() - timedelta(days=i) for i in range(days)]

    rows_all = []  # ["", None, title, url, summary, 기자, 일자]
    for d in date_list:
        if status: status.info(f"[검색] {d} 뉴스 검색 중…")
        detailed = parse_one_day(query, d, cand_cap, status=status, progress=progress)
        detailed.sort(key=lambda x: (-x["prio_score"], x["rank"]))
        picked = detailed[:per_day]
        for r in picked:
            reporter_cell = build_reporter_cell(r["host"], r["authors_list"])
            rows_all.append(["", None, r["title"], r["url"], r["summary"], reporter_cell, r["published"].strftime("%Y-%m-%d")])

    random.shuffle(rows_all)
    rows = rows_all[:days*per_day]
    for i, row in enumerate(rows, start=1): row[1] = i

    df = pd.DataFrame([[""]*7, ["", "순번","타이틀","링크","세부내용","기자","일자"], *rows])
    sheet_name = week_sheet_name_wed_kst(datetime.now(tz=KST))
    return df, sheet_name, rows

def build_html_from_rows(rows, sheet_name):
    def html_br(n=1): return "<br>"*int(n)
    def html_p(txt, small=False, bold=False):
        t = pyhtml.escape(str(txt)).replace("\n","<br>")
        if bold: t=f"<b>{t}</b>"
        if small: t=f'<span style="font-size:90%">{t}</span>'
        return f'<p align="left">{t}</p>'
    def week_label_from_sheet(s):
        m = re.match(r"(\d{2})\.(\d{2})\.(\d+)주차", s)
        return f"{m.group(1)}년 {m.group(2)}월 {m.group(3)}주차" if m else s

    parts=['<div align="">', html_p("안녕하세요."), html_br(1),
           html_p("아이디알서비스 입니다."), html_br(1),
           html_p(f"{week_label_from_sheet(sheet_name)} 에너지 뉴스 모음 입니다."), html_br(5)]
    for idx, r in enumerate(rows, start=1):
        title, url, summary, reporter, d = r[2], r[3], r[4], r[5], r[6]
        parts.append(html_p(f"{idx}.\u00A0{title}", bold=True)); parts.append(html_br(2))
        if url:
            parts.append(f'<p align="left"><a href="{pyhtml.escape(url,quote=True)}" target="_blank" rel="noopener noreferrer">기사원문</a></p>')
            parts.append(html_br(2))
        if summary: parts.append(html_p(summary))
        if reporter: parts.append(html_p(reporter, small=True, bold=True))
        if d: parts.append(html_p(d, small=True, bold=True))
        parts.append(html_br(6))
    parts.append("</div>")
    return "".join(parts)

# -------------------- GUI (Tk 기능 이식) --------------------
left, right = st.columns([1.1,1])

with left:
    st.markdown("#### 우선 키워드")
    selected = st.multiselect("체크/해제", BASE_PRIORITY_TERMS, default=BASE_PRIORITY_TERMS)
    custom_raw = st.text_input("사용자 직접 입력 키워드(쉼표로 구분)", placeholder="예: 전력거래소, 한전, 송전망, 재생에너지")

    c1, c2, c3 = st.columns(3)
    with c1:
        per_day = st.number_input("일자별 최대 건수", 1, 20, 5, 1)
    with c2:
        days = st.number_input("최근 N일(오늘 포함)", 1, 14, 7, 1)
    with c3:
        cand_cap = st.number_input("일자별 후보 파싱 상한", 10, 200, 40, 10)

    st.markdown("#### 저장 설정")
    save_dir = st.text_input("저장 폴더 경로", value="./outputs")
    save_name = st.text_input("저장 파일명(확장자 없이)", value=f"에너지뉴스_{datetime.now(tz=KST).strftime('%Y%m%d_%H%M%S')}")

    run = st.button("실행", type="primary")
    reset = st.button("초기화")

with right:
    st.markdown("#### 진행도")
    status_box = st.empty()
    progress_bar = st.progress(0)

    st.markdown("#### 로그")
    logs_area = st.empty()

if reset:
    st.session_state.logs = []

def write_logs():
    logs_area.text_area("실시간 로그", value="\n".join(st.session_state.logs[-500:]), height=260)

# -------------------- 실행 --------------------
if run:
    # 폴더 보장
    try:
        os.makedirs(save_dir, exist_ok=True)
    except Exception as e:
        st.warning(f"저장 폴더 생성 실패: {e}")

    custom_terms = [s.strip() for s in (custom_raw or "").split(",") if s.strip()]
    st.session_state.logs = []
    log("[시작] 파이프라인 실행")
    write_logs()

    df_out, sheet_name, rows = run_pipeline(
        selected_terms=set(selected),
        custom_terms=custom_terms,
        per_day=int(per_day),
        days=int(days),
        cand_cap=int(cand_cap),
        status=status_box,
        progress=progress_bar
    )

    status_box.success("완료!")
    log("[완료] 파이프라인 종료")
    write_logs()

    # 엑셀 저장(서버)
    excel_path = os.path.join(save_dir, f"{save_name}.xlsx")
    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, header=False, sheet_name=sheet_name)
        log(f"[저장] 엑셀: {excel_path}")
    except Exception as e:
        log(f"[저장 실패] 엑셀: {e}")
    write_logs()

    # HTML 코드(txt) 생성/저장
    html_txt = build_html_from_rows(rows, sheet_name)
    txt_path = os.path.join(save_dir, f"{save_name}.txt")
    try:
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(html_txt)
        log(f"[저장] HTML 코드(txt): {txt_path}")
    except Exception as e:
        log(f"[저장 실패] HTML txt: {e}")
    write_logs()

    # 다운로드 버튼
    d1, d2 = st.columns(2)
    with d1:
        buf_xlsx = io.BytesIO()
        with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, header=False, sheet_name=sheet_name)
        buf_xlsx.seek(0)
        st.download_button("엑셀 다운로드 (.xlsx)", data=buf_xlsx,
            file_name=f"{save_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with d2:
        st.download_button("HTML 코드 다운로드 (.txt)", data=html_txt.encode("utf-8"),
            file_name=f"{save_name}.txt", mime="text/plain; charset=utf-8")

    st.divider()

    # 표 미리보기
    try:
        df_prev = pd.read_excel(io.BytesIO(buf_xlsx.getvalue()), sheet_name=sheet_name, header=None)
        df_prev.columns = df_prev.iloc[1]; df_prev = df_prev.iloc[2:].reset_index(drop=True)
        st.subheader("수집 결과 (표 미리보기)")
        st.dataframe(df_prev, use_container_width=True, height=380)
    except Exception as e:
        st.warning(f"표 미리보기를 만들 수 없습니다: {e}")

    # HTML 코드 미리보기(렌더 + 코드)
    st.subheader("전체 HTML 렌더 미리보기")
    st.components.v1.html(html_txt, height=600, scrolling=True)

    st.subheader("전체 HTML 코드 (txt 미리보기)")
    st.code(html_txt, language="html")

    # -------- 제목 클릭 → 모달 팝업(개별 HTML) --------
    st.subheader("제목 목록 (클릭하면 팝업에서 개별 미리보기)")

    # dialog 정의
    @st.dialog("기사 미리보기", width="large")
    def show_preview_dialog(item_html):
        st.components.v1.html(item_html, height=420, scrolling=True)

    # 행 유틸
    def choose_title_col(df: pd.DataFrame) -> str:
        for c in df.columns:
            if "타이틀" in str(c): return c
        return "타이틀"

    # rows는 이미 확정된 리스트 형식이므로 각 항목을 버튼으로 노출
    for idx, r in enumerate(rows, start=1):
        title = r[2] or "(제목 없음)"
        url   = r[3]
        colA, colB = st.columns([0.75, 0.25])
        with colA:
            if st.button(f"{idx}. {title}", key=f"title_btn_{idx}"):
                # 개별 HTML 구성 후 모달 오픈
                # (df 없이 rows 직접 사용)
                def build_item_html():
                    def html_br(n=1): return "<br>"*int(n)
                    def html_p(txt, small=False, bold=False):
                        t = pyhtml.escape(str(txt)).replace("\n","<br>")
                        if bold: t=f"<b>{t}</b>"
                        if small: t=f'<span style="font-size:90%">{t}</span>'
                        return f'<p align="left">{t}</p>'
                    parts=['<div align="">', html_p(f"{idx}. {title}", bold=True), html_br(1)]
                    if url:
                        parts.append(f'<p align="left"><a href="{pyhtml.escape(url,quote=True)}" target="_blank" rel="noopener noreferrer">기사원문</a></p>')
                        parts.append(html_br(1))
                    if r[4]: parts.append(html_p(r[4]))
                    if r[5]: parts.append(html_p(r[5], small=True, bold=True))
                    if r[6]: parts.append(html_p(r[6], small=True, bold=True))
                    parts.append("</div>")
                    return "".join(parts)
                show_preview_dialog(build_item_html())
        with colB:
            if url:
                st.link_button("원문 열기", url, use_container_width=True)
            else:
                st.button("원문 없음", disabled=True, use_container_width=True)

    progress_bar.progress(1.0)
    status_box.empty()

# 항상 최신 로그 보이기
write_logs()
