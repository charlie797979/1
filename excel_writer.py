"""업무파일(xlsx)에 기사 행을 기록하는 모듈. 기존 서식·셀 간격은 건드리지 않고 값/하이퍼링크만 입력한다."""
import io
import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment

FONT = "맑은 고딕"
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _is_example_row(ws, r):
    v = ws.cell(r, 1).value
    return isinstance(v, str) and "예시" in v


def first_empty_row(ws):
    """순번(A)이 숫자이고 업무일자(B)가 비어 있는 첫 행. 없으면 마지막 입력행 다음 행."""
    last = ws.max_row
    for r in range(3, last + 1):
        if _is_example_row(ws, r):
            continue
        if ws.cell(r, 2).value in (None, ""):
            return r
    return last + 1


def count_filled(ws):
    return sum(
        1 for r in range(3, ws.max_row + 1)
        if not _is_example_row(ws, r) and ws.cell(r, 2).value not in (None, "")
    )


def write_articles(template_bytes, articles_by_sheet, work_date):
    """
    articles_by_sheet: {시트명: [{'pub_date': '2026.10.10', 'keyword': '..', 'summary': '..',
                                   'title': '..', 'url': '..'}, ...]}
    work_date: 'YYYY.MM.DD'
    return: (xlsx bytes, {시트명: 기록건수})
    """
    wb = load_workbook(io.BytesIO(template_bytes))
    written = {}
    for sheet, items in articles_by_sheet.items():
        if sheet not in wb.sheetnames or not items:
            continue
        ws = wb[sheet]
        r = first_empty_row(ws)
        for it in items:
            # 순번이 없는 행(양식 끝을 넘어가는 경우)은 직접 부여
            if ws.cell(r, 1).value in (None, ""):
                ws.cell(r, 1).value = r - 2 - sum(1 for x in range(3, r) if _is_example_row(ws, x))
            row_vals = {2: work_date, 3: it["pub_date"], 4: it["keyword"], 5: it["summary"]}
            for col, val in row_vals.items():
                c = ws.cell(r, col)
                c.value = val
                c.number_format = "@"
            link = ws.cell(r, 6)
            link.value = it["title"]
            link.hyperlink = it["url"]
            link.font = Font(name=FONT, size=10, bold=False, underline="single", color="0563C1")
            for col in range(1, 7):
                c = ws.cell(r, col)
                if col != 6:
                    c.font = Font(name=FONT, size=10, bold=False)
                c.border = BORDER
                if col in (5, 6):
                    c.alignment = Alignment(wrap_text=True, horizontal="left", vertical="center")
                elif c.alignment.horizontal is None:
                    c.alignment = Alignment(horizontal="center", vertical="center")
            r += 1
        written[sheet] = len(items)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), written


def update_monthly_report(xlsx_bytes, week_row, text):
    """월말보고서 시트의 주차별 문구 셀(B열)을 갱신. week_row 예: 5(1주차), 10(2주차) ..."""
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb["월말보고서"]
    ws.cell(week_row, 2).value = text
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
