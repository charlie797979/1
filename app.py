import html as htmllib
import io
import re
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, time
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook

try:
    import anthropic
except ImportError:  # requirements.txt에 anthropic이 없으면 기존 요약 방식 사용
    anthropic = None

import news_search as ns
import excel_writer as xw

st.set_page_config(page_title="뉴스 클리핑 모니터링", page_icon="📰", layout="wide")

BASE_DIR = Path(__file__).parent

# 원래의 조회 시작 시각 함수를 보관 (재실행되어도 덮어쓰이지 않도록 한 번만 저장)
if not hasattr(ns, "_orig_window_start"):
    ns._orig_window_start = ns.window_start

# ---------- 제외 규칙 (필요하면 단어를 추가/삭제하세요) ----------
STOCK_WORDS = [
    "주가", "목표주가", "목표가", "투자의견", "투자 의견", "증권", "코스피", "코스닥",
    "종목", "증시", "상한가", "하한가", "급등", "급락", "순매수", "순매도",
    "매매동향", "시가총액", "테마주", "특징주", "컨센서스", "공매도", "ETF",
    "개미", "외국인 매수", "외국인 매도", "기관 매수", "기관 매도", "리포트",
    "매수 의견", "매도 의견", "주식", "배당",
]
# 제목에 '매출/영업이익/실적' 등이 '증가/성장/호조...'와 함께 나오는 단순 실적 기사
REVENUE_RE = re.compile(
    r"(매출|영업이익|영업익|순이익|실적|수익|흑자)[^,.…]{0,12}"
    r"(증가|상승|성장|개선|호조|급증|늘|최대|돌파|훨훨|껑충|기대|전망|웃돌)"
)
EARNINGS_WORDS = ["어닝", "서프라이즈", "분기 실적", "잠정실적", "잠정 실적"]


def _aware(dt):
    return dt if dt.tzinfo else dt.replace(tzinfo=ns.KST)


def exclude_reason(article, kw):
    """제외해야 하는 기사면 사유(문자열), 아니면 None."""
    title = str(article.get("title", ""))
    text = title + " " + str(ns.draft_summary(article, kw))
    for w in STOCK_WORDS:
        if w in title:
            return f"주식 관련 ('{w}')"
    # 요약에는 주식 단어가 2개 이상 있을 때만 제외 (오탐 방지)
    hits = [w for w in STOCK_WORDS if w in text]
    if len(hits) >= 2:
        return f"주식 관련 ('{hits[0]}', '{hits[1]}')"
    if REVENUE_RE.search(title) or any(w in title for w in EARNINGS_WORDS):
        return "단순 실적/매출 증가"
    return None


# ---------- 한 문장 요약 (Claude API) ----------
SUMMARY_MODELS = ["claude-haiku-5-5", "claude-haiku-4-5"]  # 앞에서부터 시도


def fetch_body(url, limit=4000):
    """기사 본문을 최대한 가져온다. 실패하면 빈 문자열."""
    try:
        r = requests.get(url, timeout=6, headers={"User-Agent": "Mozilla/5.0"})
        r.encoding = r.apparent_encoding or r.encoding
        page = re.sub(r"(?is)<(script|style).*?</\1>", " ", r.text)
        paras = re.findall(r"(?is)<p[^>]*>(.*?)</p>", page)
        out = []
        for p in paras:
            t = htmllib.unescape(re.sub(r"(?s)<[^>]+>", " ", p))
            t = re.sub(r"\s+", " ", t).strip()
            if len(t) >= 30:
                out.append(t)
        return " ".join(out)[:limit]
    except Exception:
        return ""


def ai_summary(client, article, kw):
    """기사를 키워드 관점에서 한 문장으로 요약. 실패하면 None."""
    title = str(article.get("title", ""))
    body = fetch_body(article["url"])
    if len(body) < 200:  # 본문 확보 실패 → 제목 + 검색 발췌문으로 대체
        body = str(ns.draft_summary(article, kw))
    prompt = (
        f"뉴스클리핑 보고서에 넣을 요약을 작성하세요.\n"
        f"- 모니터링 키워드: {kw}\n"
        f"- 한국어 한 문장, 110~140자, 핵심 사실(누가/무엇을/수치)에 배경이나 목적을 한 가지 덧붙여 구체적으로\n"
        f"- 문장은 '~했다.' 형태의 평서문으로 끝내기, 수식어·평가·광고성 표현 제외\n"
        f"- 요약문 한 문장만 출력 (따옴표, 설명, 줄바꿈 금지)\n\n"
        f"제목: {title}\n본문: {body}"
    )
    return call_claude(client, prompt)


def call_claude(client, prompt, max_tokens=200):
    """모델을 차례로 시도. (요약문, 오류메시지) 반환."""
    err = None
    for model in SUMMARY_MODELS:
        try:
            msg = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                messages=[{"role": "user", "content": prompt}],
            )
            text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
            text = re.sub(r"\s+", " ", text).strip()
            if text:
                return text, None
            err = f"{model}: 빈 응답"
        except Exception as e:
            err = f"{model}: {type(e).__name__}: {e}"
    return None, err


# ---------- 접속 비밀번호(선택) ----------
pw = st.secrets.get("APP_PASSWORD", "")
if pw:
    if not st.session_state.get("ok"):
        x = st.text_input("비밀번호", type="password")
        if x == pw:
            st.session_state["ok"] = True
            st.rerun()
        st.stop()

st.title("📰 ㈜셀바이오휴먼텍 뉴스클리핑 업무")

cid = st.secrets.get("NAVER_CLIENT_ID", "")
csec = st.secrets.get("NAVER_CLIENT_SECRET", "")
if not cid or not csec:
    st.error("네이버 검색 API 키가 설정되지 않았습니다. Streamlit 앱 설정의 Secrets를 확인하세요.")
    st.stop()


def find_template():
    for p in (BASE_DIR / "template.xlsx", BASE_DIR / "template" / "template.xlsx"):
        if p.exists():
            return p
    return None


# ---------- 1. 업무파일 ----------
with st.container():
    # 업무파일은 항상 저장소의 기본 양식(template.xlsx)을 사용
    if "work_bytes" not in st.session_state:
        tpl = find_template()
        if tpl is None:
            st.error("기본 양식(template.xlsx)이 저장소에 없습니다. GitHub 저장소를 확인하세요.")
            st.stop()
        st.session_state["work_bytes"] = tpl.read_bytes()
        st.session_state.pop("written", None)

    wb = load_workbook(io.BytesIO(st.session_state["work_bytes"]))
    sheets = [s for s in wb.sheetnames if s != "월말보고서"]

    st.subheader("1. 키워드(시트)")
    chosen = st.multiselect("모니터링할 키워드", sheets, default=sheets)

    # ---------- 3. 조회 범위 (직접 변경 가능) ----------
    st.subheader("2. 조회 범위")
    now = datetime.now(ns.KST)
    default_start = _aware(ns._orig_window_start(now))

    c1, c2, c3, c4 = st.columns(4)
    start_date = c1.date_input("시작 날짜", value=default_start.date(), key="start_date")
    start_time = c2.time_input("시작 시각", value=time(0, 0), key="start_time")
    end_date = c3.date_input("종료 날짜", value=now.date(), key="end_date")
    end_time = c4.time_input("종료 시각", value=time(0, 0), key="end_time")
    end_dt = datetime.combine(end_date, end_time).replace(tzinfo=ns.KST)

    start_dt = datetime.combine(start_date, start_time).replace(tzinfo=ns.KST)
    if start_dt >= end_dt:
        st.error("시작 일시가 종료 일시보다 빠르거나 같아야 합니다.")
        st.stop()

    # ---------- 4. 제외 설정 ----------
    st.subheader("3. 기사 제외")
    use_filter = st.checkbox("주식·단순 매출 증가 기사 제외", value=True)

    work_date = st.text_input("업무일자 (0000.00.00)", value=now.strftime("%Y.%m.%d"))

    # ---------- AI 요약 연결 테스트 ----------
    if st.button("🔧 AI 요약 연결 테스트"):
        _key = str(st.secrets.get("ANTHROPIC_API_KEY", "")).strip()
        if not _key:
            st.error("Secrets에 ANTHROPIC_API_KEY가 없습니다.")
        elif anthropic is None:
            st.error("requirements.txt에 anthropic이 없습니다.")
        else:
            _text, _err = call_claude(
                anthropic.Anthropic(api_key=_key), "'연결 성공'이라고만 답하세요.", max_tokens=20
            )
            if _text:
                st.success(f"연결 성공: {_text}")
            else:
                st.error(f"연결 실패 - {_err}")

# ---------- 기사 검색 ----------
st.divider()


def wide_button(label, **kw):
    """가로 전체 폭 버튼 (streamlit 버전에 따라 옵션 이름이 달라 둘 다 대응)."""
    try:
        return st.button(label, width="stretch", **kw)
    except TypeError:
        return st.button(label, use_container_width=True, **kw)


if wide_button("🔍 기사 검색", type="primary"):
    if not chosen:
        st.warning("키워드를 하나 이상 선택하세요.")
    else:
        res, excluded = {}, {}
        prog = st.progress(0.0)
        # 검색 동안만 조회 시작 시각을 사용자가 고른 값으로 바꿨다가 원래대로 되돌림
        ns.window_start = lambda _now, _s=start_dt: _s
        try:
            for i, kw in enumerate(chosen):
                try:
                    arts = ns.collect(kw, cid, csec)
                except Exception as e:
                    st.warning(f"'{kw}' 검색 실패: {e}")
                    arts = []

                # 선택한 기간 안의 기사만 남김
                arts = [a for a in arts if start_dt <= _aware(a["pub"]) <= end_dt]

                kept, dropped = [], []
                for a in arts:
                    reason = exclude_reason(a, kw) if use_filter else None
                    if reason:
                        dropped.append((a, reason))
                    else:
                        kept.append(a)
                res[kw] = kept
                excluded[kw] = dropped
                prog.progress((i + 1) / len(chosen))
        finally:
            ns.window_start = ns._orig_window_start

        # 검색 단계에서는 요약하지 않음 (API 절약) - 선택한 기사만 저장 단계에서 요약
        st.session_state["results"] = res
        st.session_state["excluded"] = excluded
        st.session_state.pop("written", None)

results = st.session_state.get("results")
if not results:
    st.write("위에서 키워드를 고른 뒤 **기사 검색**을 누르세요.")
    st.stop()

excluded = st.session_state.get("excluded", {})

# ---------- 기사 선택 ----------
st.subheader("업무파일에 넣을 기사를 선택하세요")
st.caption("체크한 기사만 요약·저장됩니다. 기사의 주 내용이 키워드와 관련 있는지 직접 확인하고 체크하세요.")

edited = {}
for kw, arts in results.items():
    n_drop = len(excluded.get(kw, []))
    label = f"{kw}  ({len(arts)}건" + (f", 제외 {n_drop}건)" if n_drop else ")")
    with st.expander(label, expanded=bool(arts)):
        if not arts:
            st.write("해당 기간에 조건에 맞는 기사가 없습니다.")
        else:
            df = pd.DataFrame(
                [
                    {
                        "선택": False,  # 전부 해제 상태, 직접 체크
                        "기사발행일": a["pub"].strftime("%Y.%m.%d"),
                        "제목": a["title"],
                        "URL": a["url"],
                    }
                    for a in arts
                ]
            )
            edited[kw] = st.data_editor(
                df,
                key=f"ed_{kw}",
                hide_index=True,
                column_order=["선택", "기사발행일", "제목"],
                column_config={
                    "선택": st.column_config.CheckboxColumn(width="small"),
                    "제목": st.column_config.TextColumn(width="large"),
                },
                disabled=["기사발행일", "제목", "URL"],
            )

        if n_drop:
            st.caption("제외된 기사 (잘못 제외됐는지 확인용)")
            st.dataframe(
                pd.DataFrame(
                    [
                        {"제외 사유": r, "제목": a["title"], "URL": a["url"]}
                        for a, r in excluded[kw]
                    ]
                ),
                hide_index=True,
                column_config={"URL": st.column_config.LinkColumn()},
            )

# ---------- 선택한 기사만 요약 + 엑셀 저장 ----------
n_selected = sum(int(df["선택"].sum()) for df in edited.values())
st.write(f"선택한 기사: **{n_selected}건**")

if st.button("📝 선택한 기사 요약하고 업무파일에 저장", type="primary", disabled=n_selected == 0):
    # 체크한 기사 모으기
    picked = []  # (키워드, 기사발행일, 제목, URL)
    for kw, df in edited.items():
        for _, r in df[df["선택"]].iterrows():
            if not str(r["제목"]).strip() or not str(r["URL"]).startswith("http"):
                st.warning(f"{kw}: 제목/URL이 비어 있는 행은 제외했습니다.")
                continue
            picked.append((kw, r["기사발행일"], str(r["제목"]).strip(), r["URL"]))

    # 선택한 기사만 Claude로 요약 (실패하면 검색 API 발췌문 기반 기존 요약 사용)
    art_by_url = {a["url"]: a for arts in results.values() for a in arts}
    summaries = {}
    api_key = str(st.secrets.get("ANTHROPIC_API_KEY", "")).strip()
    if not api_key:
        st.warning("Secrets에 ANTHROPIC_API_KEY가 없어 기존 요약을 사용합니다.")
    elif anthropic is None:
        st.warning("requirements.txt에 anthropic이 없어 기존 요약을 사용합니다.")
    else:
        client = anthropic.Anthropic(api_key=api_key)
        sprog = st.progress(0.0, text="선택한 기사 요약 중...")
        errors = []
        with ThreadPoolExecutor(max_workers=5) as pool:
            futures = [
                (url, pool.submit(ai_summary, client, art_by_url[url], kw))
                for kw, _, _, url in picked
            ]
            for j, (url, fut) in enumerate(futures):
                text, err = fut.result()
                summaries[url] = text
                if err:
                    errors.append(err)
                sprog.progress((j + 1) / len(futures), text="선택한 기사 요약 중...")
        sprog.empty()
        if errors:
            st.warning(f"AI 요약 {len(errors)}건 실패 (해당 기사는 기존 요약 사용). 실패 원인:\n\n{errors[0]}")

    by_sheet = {}
    for kw, pub_date, title, url in picked:
        summary = summaries.get(url) or ns.draft_summary(art_by_url[url], kw)
        by_sheet.setdefault(kw, []).append(
            {"pub_date": pub_date, "keyword": kw, "summary": summary, "title": title, "url": url}
        )

    out, written = xw.write_articles(st.session_state["work_bytes"], by_sheet, work_date)
    st.session_state["work_bytes"] = out
    st.session_state["written"] = written
    st.session_state["done_name"] = (
        f"㈜셀바이오휴먼텍 뉴스클리핑 업무파일({datetime.now(ns.KST).month}월_한철희).xlsx"
    )

if st.session_state.get("written") is not None:
    w = st.session_state["written"]
    if w:
        st.success("기록 완료: " + ", ".join(f"{k} {v}건" for k, v in w.items()))
    else:
        st.info("선택된 기사가 없습니다.")
    st.download_button(
        "⬇️ 업무파일 다운로드",
        st.session_state["work_bytes"],
        file_name=st.session_state["done_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.warning("제출 전 확인: 다운로드한 파일에서 하이퍼링크를 클릭해 정상 연결되는지 확인하세요.")
