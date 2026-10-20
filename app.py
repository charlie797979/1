import datetime as dt
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import core

st.set_page_config(page_title="뉴스클리핑 업무", page_icon="📰", layout="wide")

TEMPLATE = Path(__file__).parent / "template.xlsx"  # 업무파일(엑셀) 기본 양식
WORKER = "한철희"

if not TEMPLATE.exists():
    st.error("template.xlsx(업무파일 양식)를 app.py와 같은 폴더에 올려주세요.")
    st.stop()


@st.cache_data
def load_keywords():
    return core.keyword_sheets(TEMPLATE)


def secret(name):
    try:
        return st.secrets[name]
    except Exception:
        return os.environ.get(name)


ss = st.session_state
ss.setdefault("articles", [])
ss.setdefault("search_id", 0)
ss.setdefault("stats", "")
ss.setdefault("result", None)

st.title("📰 ㈜셀바이오휴먼텍 뉴스클리핑 업무")

# 1. 키워드
st.header("1. 키워드(시트)")
options = load_keywords()
keywords = st.multiselect("모니터링할 키워드", options, default=options)

# 2. 조회 범위
st.header("2. 조회 범위")
now = dt.datetime.now(core.KST)
c1, c2, c3, c4 = st.columns(4)
start_d = c1.date_input("시작 날짜", now.date() - dt.timedelta(days=1), format="YYYY/MM/DD")
start_t = c2.time_input("시작 시각", dt.time(0, 0))
end_d = c3.date_input("종료 날짜", now.date(), format="YYYY/MM/DD")
end_t = c4.time_input("종료 시각", dt.time(0, 0))

# 3. 기사 제외
st.header("3. 기사 제외")
e1, e2, e3 = st.columns(3)
ex_stock = e1.checkbox("주식 관련 뉴스 제외", value=True)
ex_sales = e2.checkbox("단순 매출 증가 뉴스 제외", value=True)
ex_nokw = e3.checkbox("제목에 키워드가 없는 기사 제외", value=False)
extra = st.text_input("추가로 제외할 단어 (쉼표로 구분)", "")

# 기사 검색
if st.button("기사 검색", type="primary"):
    start_dt = dt.datetime.combine(start_d, start_t, tzinfo=core.KST)
    end_dt = dt.datetime.combine(end_d, end_t, tzinfo=core.KST)
    nid, nsec = secret("NAVER_CLIENT_ID"), secret("NAVER_CLIENT_SECRET")
    if not (nid and nsec):
        st.error("Streamlit Secrets에 NAVER_CLIENT_ID와 NAVER_CLIENT_SECRET을 먼저 등록해 주세요.")
    elif not keywords:
        st.warning("키워드를 하나 이상 선택해 주세요.")
    elif start_dt >= end_dt:
        st.warning("종료 일시가 시작 일시보다 늦어야 합니다.")
    else:
        extra_words = [w.strip() for w in extra.split(",") if w.strip()]
        found, failed = [], []
        with st.spinner("기사를 검색하는 중입니다..."):
            for kw in keywords:
                try:
                    found += core.search_news(kw, start_dt, end_dt, nid, nsec)
                except Exception as e:
                    failed.append(f"{kw} ({type(e).__name__})")
        kw_order = {k: i for i, k in enumerate(keywords)}
        found.sort(key=lambda x: (x["pub"], kw_order[x["keyword"]]))
        total = len(found)
        kept = [a for a in found if not core.is_excluded(a["title"], ex_stock, ex_sales, extra_words)]
        if ex_nokw:
            kept = [a for a in kept if a["keyword"] in a["title"]]
        n_excluded = total - len(kept)
        unique = core.dedupe(kept, strip_words=keywords)
        unique.sort(key=lambda x: x["pub"], reverse=True)
        ss.articles = unique
        ss.search_id += 1
        ss.result = None
        ss.stats = f"검색 {total}건 · 제외 {n_excluded}건 · 중복 {len(kept) - len(unique)}건 → {len(unique)}건"
        if failed:
            st.warning("검색에 실패한 키워드: " + ", ".join(failed))

# 검색 결과 표
if ss.articles:
    st.caption(ss.stats)
    table = pd.DataFrame({
        "선택": [False] * len(ss.articles),  # 전부 해제 상태로 시작
        "기사발행일": [a["pub"].strftime("%Y.%m.%d") for a in ss.articles],
        "제목": [a["title"] for a in ss.articles],
    })
    edited = st.data_editor(
        table,
        hide_index=True,
        disabled=["기사발행일", "제목"],
        column_config={
            "선택": st.column_config.CheckboxColumn("선택", width="small"),
            "기사발행일": st.column_config.TextColumn("기사발행일", width="small"),
            "제목": st.column_config.TextColumn("제목", width="large"),
        },
        height=min(36 * len(table) + 40, 640),
        key=f"editor_{ss.search_id}",
    )

    # 요약하기
    if st.button("요약하기", type="primary"):
        picked = [ss.articles[i] for i in edited.index[edited["선택"]]]
        if not picked:
            st.warning("요약할 기사를 먼저 체크해 주세요.")
        else:
            api_key = secret("ANTHROPIC_API_KEY")
            rows, fallbacks = [], 0
            bar = st.progress(0.0, text="기사 내용을 요약하는 중입니다...")
            for n, a in enumerate(picked, 1):
                body = core.fetch_body(a["url"])
                summary, how = core.summarize(a["title"], body, api_key)
                fallbacks += how == "fallback"
                rows.append({**a, "summary": summary})
                bar.progress(n / len(picked), text=f"요약 중… ({n}/{len(picked)})")
            bar.empty()
            data, counts = core.build_workbook(TEMPLATE, rows, now.date())
            ss.result = {
                "bytes": data,
                "name": f"㈜셀바이오휴먼텍 뉴스클리핑 업무파일({now.month}월_{WORKER}).xlsx",
                "counts": counts,
                "fallbacks": fallbacks,
                "preview": pd.DataFrame({
                    "키워드": [r["keyword"] for r in rows],
                    "기사발행일": [r["pub"].strftime("%Y.%m.%d") for r in rows],
                    "제목": [r["title"] for r in rows],
                    "요약": [r["summary"] for r in rows],
                }),
            }

# 결과 / 다운로드
if ss.result:
    res = ss.result
    st.success("엑셀 입력 완료: " + ", ".join(f"{k} {v}건" for k, v in res["counts"].items()))
    if res["fallbacks"]:
        st.warning(f"{res['fallbacks']}건은 AI 요약을 쓰지 못해 본문 앞부분(또는 제목)으로 대신 입력했습니다. "
                   "ANTHROPIC_API_KEY 설정과 기사 본문 접근 여부를 확인해 주세요.")
    st.dataframe(res["preview"], hide_index=True)
    st.download_button(
        "엑셀 파일 다운로드",
        data=res["bytes"],
        file_name=res["name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
