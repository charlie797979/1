"""네이버 뉴스 검색 API로 키워드 기사를 수집하고, 시간 범위·중복·키워드 적합성을 1차 필터링한다."""
import re
import html
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
import requests

KST = timezone(timedelta(hours=9))
API = "https://openapi.naver.com/v1/search/news.json"


def window_start(now=None):
    """전일 13:00 (KST) ~ 현재"""
    now = now or datetime.now(KST)
    y = now - timedelta(days=1)
    return y.replace(hour=13, minute=0, second=0, microsecond=0)


def clean(t):
    t = re.sub(r"<[^>]+>", "", t or "")
    return html.unescape(t).strip()


def norm_title(t):
    return re.sub(r"[^0-9a-zA-Z가-힣]", "", t).lower()


def fetch(keyword, client_id, client_secret, display=100, pages=2):
    headers = {"X-Naver-Client-Id": client_id, "X-Naver-Client-Secret": client_secret}
    out = []
    for p in range(pages):
        r = requests.get(
            API,
            headers=headers,
            params={"query": keyword, "display": display, "start": 1 + p * display, "sort": "date"},
            timeout=10,
        )
        r.raise_for_status()
        items = r.json().get("items", [])
        for it in items:
            try:
                pub = parsedate_to_datetime(it["pubDate"]).astimezone(KST)
            except Exception:
                continue
            out.append({
                "keyword": keyword,
                "title": clean(it["title"]),
                "desc": clean(it["description"]),
                "url": it.get("originallink") or it["link"],
                "naver_url": it["link"],
                "pub": pub,
            })
        if len(items) < display:
            break
    return out


def relevance(article, keyword):
    """키워드가 제목에 있으면 '높음', 제목엔 없고 본문 요약에만 있으면 '검토'. 없으면 '낮음'."""
    kw = keyword.lower()
    if kw in article["title"].lower():
        return "높음"
    if kw in article["desc"].lower():
        return "검토"
    return "낮음"


def collect(keyword, client_id, client_secret, now=None):
    start = window_start(now)
    rows = fetch(keyword, client_id, client_secret)
    rows = [a for a in rows if a["pub"] >= start]
    seen, uniq = set(), []
    for a in sorted(rows, key=lambda x: x["pub"], reverse=True):
        key = norm_title(a["title"])[:30]
        if key in seen:
            continue
        seen.add(key)
        a["relevance"] = relevance(a, keyword)
        uniq.append(a)
    return [a for a in uniq if a["relevance"] != "낮음"]


def draft_summary(article, keyword):
    """검색 API가 주는 요약문을 1~2문장으로 정리한 초안. 반드시 사용자가 검토·수정."""
    d = article["desc"]
    sents = re.split(r"(?<=[.다요])\s+", d)
    s = " ".join(sents[:2]).strip()
    if keyword not in s:
        s = f"{keyword} 관련: {s}"
    return s
