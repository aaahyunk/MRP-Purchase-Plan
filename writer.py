"""결과 엑셀 생성: 템플릿 서식 + MRP(값) 시트 + 구매계획_통합(수식) + 휴일/신규품목/검증로그.

'구매계획_통합'은 기존과 동일한 수식을 사용한다 → 엑셀에서 C3/E3(기간)을 바꾸면 그대로 재계산.
변경점: V(기간말 과부족) 정의, L(발주일) WORKDAY에 '휴일' 시트 반영.
"""
import io
import datetime as dt
from copy import copy

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

from . import config as C
from . import custom as CU
from .engine import order_rows, is_non_mass, NON_MASS_ZERO_RUN, NON_MASS_IGNORE_TAIL
from .loader import _num

NA_FILL = "FFFF0000"          # 대체품목(규격 N/A) 행 배경 = 빨간색
MASS_RED = "FFFF0000"         # '양산품목X' 글자색


def _remark(base, non_mass):
    """비고 셀 값: 사람이 쓴 내용 유지 + 필요 시 '양산품목X'(빨간 글자) 덧붙임."""
    base = "" if base in (None, "") else str(base)
    if not non_mass:
        return base or None
    red = TextBlock(InlineFont(rFont=FONT, sz=10, color=MASS_RED, b=True), CU.MASS_TAG)
    return CellRichText([base + CU.MASS_SEP, red]) if base else CellRichText([red])

FONT = "Pretendard"
HDR_FILL = PatternFill("solid", fgColor="FF1E2761")
THIN = Side(style="thin", color="FFBFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _txt(cell):
    """'#N/A' 같은 문자열이 엑셀 오류값으로 저장되지 않도록 텍스트로 고정."""
    if cell.data_type == "e":
        cell.data_type = "s"
    return cell


def _load(template):
    if isinstance(template, (bytes, bytearray)):
        template = io.BytesIO(template)
    return openpyxl.load_workbook(template)


def _hdr(ws, row, values, widths=None):
    for j, v in enumerate(values, 1):
        c = ws.cell(row, j, v)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFFFF")
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[CL(j)].width = w


def _body(ws, r0, rows, fmts=None):
    f = Font(name=FONT, size=10)
    for i, vals in enumerate(rows):
        for j, v in enumerate(vals, 1):
            c = _txt(ws.cell(r0 + i, j, v))
            c.font, c.border = f, BOX
            if fmts and fmts.get(j):
                c.number_format = fmts[j]


def write_mrp_sheet(wb, mrp):
    ws = wb[C.MRP_SHEET]
    styles = [copy(ws.cell(2, c)._style) for c in range(1, C.N_MRP_COLS + 1)]
    for j, d in enumerate(mrp.dates):
        c = ws.cell(1, C.N_FIXED + 1 + j, dt.datetime.combine(d, dt.time()))
        c.number_format = "mm/dd(aaa)"
    for i, row in enumerate(mrp.rows):
        r = 2 + i
        for j in range(C.N_MRP_COLS):
            v = row[j]
            cell = _txt(ws.cell(r, j + 1, v))
            cell._style = copy(styles[j])
    ws.freeze_panes = "K2"
    ws.auto_filter.ref = f"A1:{CL(C.N_MRP_COLS)}{len(mrp.rows) + 1}"


def write_holiday_sheet(wb, eff):
    """A열: 실제 적용 휴일 목록(WORKDAY 참조 범위) / C~G열: 전체 공휴일과 적용 판단."""
    ws = wb.create_sheet(C.HOLIDAY_SHEET)
    f = Font(name=FONT, size=10)
    heads = {1: "적용 휴일\n(WORKDAY 참조)", 3: "날짜", 4: "명칭", 5: "해당일 생산계획\n소요합계", 6: "적용", 7: "사유"}
    for j, w in zip(range(1, 8), [16, 3, 15, 24, 16, 7, 36]):
        ws.column_dimensions[CL(j)].width = w
    for j, v in heads.items():
        c = ws.cell(1, j, v)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFFFF")
        c.fill, c.border = HDR_FILL, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    applied = [h["날짜"] for h in eff if h["적용"]]
    for i, d in enumerate(applied):
        c = ws.cell(2 + i, 1, dt.datetime.combine(d, dt.time()))
        c.number_format, c.font, c.border = "yyyy-mm-dd(aaa)", f, BOX
    for i, h in enumerate(eff):
        vals = [dt.datetime.combine(h["날짜"], dt.time()), h["명칭"], h["소요합계"], "O" if h["적용"] else "X", h["사유"]]
        for j, v in enumerate(vals, 3):
            c = ws.cell(2 + i, j, v)
            c.font, c.border = f, BOX
        ws.cell(2 + i, 3).number_format = "yyyy-mm-dd(aaa)"
        ws.cell(2 + i, 5).number_format = "#,##0"
        ws.cell(2 + i, 6).alignment = Alignment(horizontal="center")
    note = len(eff) + 3
    ws.cell(note, 3, "※ 생산계획(그날 전체 품목 소요량 합계 > 0)이 있는 공휴일은 근무일로 보고 휴일에서 제외합니다. "
                     "A열 목록이 '부족일 D-3' WORKDAY 계산에 쓰입니다. 공휴일 추가·수정은 holidays.csv 에서 합니다."
            ).font = Font(name=FONT, size=9, color="FF595959")
    ws.freeze_panes = "A2"


def write_new_items_sheet(wb, new_items):
    ws = wb.create_sheet("신규품목")
    cols = ["포함", "품번", "출처", "BPA", "공급자", "규격", "구분", "비고", "MOQ",
            "담당자", "공정", "재고", "기간소요", "기간최저", "첫부족일", "BPA출처"]
    _hdr(ws, 1, ["오늘 추가"] + cols[1:], [9, 14, 9, 13, 16, 40, 8, 12, 8, 10, 8, 11, 11, 11, 11, 10])
    _body(ws, 2, [["O" if r.get("포함", True) else "제외"] + [r.get(k, "") for k in cols[1:]] for r in new_items],
          {12: "#,##0", 13: "#,##0", 14: "#,##0"})
    ws.freeze_panes = "C2"
    if new_items:
        ws.auto_filter.ref = f"A1:{CL(len(cols))}{len(new_items) + 1}"


def write_log_sheet(wb, checks, params, summary_, exc):
    ws = wb.create_sheet("검증로그")
    _hdr(ws, 1, ["구분", "항목", "내용"], [10, 34, 90])
    mark = {"ok": "✅ 정상", "info": "ℹ 참고", "warn": "⚠ 확인", "error": "⛔ 오류"}
    rows = [["설정", "생성시각", params.get("생성시각", "")],
            ["설정", "MRP 파일", params.get("파일명", "")],
            ["설정", "최저값 기간", f"{params['start']} ~ {params['end']} (C3~E3, 엑셀에서 변경 가능)"],
            ["설정", "기간말 과부족 정의", "210 + 미출 + 대여 + 발주잔량 + 재고 − Σ소요량(C3~E3)"],
            ["설정", "대체품목", "규격에 'N/A' 포함 → 품목 두 행 전체 빨간 배경 (조건부서식, 규격 수정 시 자동 반영)"],
            ["설정", "양산품목X", f"일별 소요량(X~AZ) 0이 {NON_MASS_ZERO_RUN}번 이상 연속 → 비고에 '양산품목X'(빨간 글자) 추가"
             + (" · MRP 기간 끝의 연속 0은 제외" if params.get("ignore_tail", NON_MASS_IGNORE_TAIL) else "")]]
    rows += [[mark.get(c.level, c.level), c.title, c.detail] for c in checks]
    rows += [["결과", k, f"{v:,}"] for k, v in summary_.items()]
    rows += [["예외", "마스터에 있으나 MRP에 없는 품목 (0으로 표시)", f"{len(exc['not_in_mrp']):,}건"],
             ["예외", "BPA가 단가시트와 다른 품목", f"{len(exc['bpa_diff'])}건: " +
              ", ".join(f"{d['품번']}({d['마스터BPA']}→{d['단가시트BPA']})" for d in exc["bpa_diff"][:15])],
             ["예외", "담당자 #N/A (담당자 시트에 지점 없음)", f"{len(exc['owner_na'])}건: " + ", ".join(exc["owner_na"][:20])]]
    _body(ws, 2, rows)
    for r in range(2, len(rows) + 2):
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")


def write_plan_sheet(wb, mrp, master, params, n_hol):
    ws = wb[C.PLAN_SHEET]
    ncol = C.N_PLAN_COLS
    protos = {}
    for i, k in enumerate(C.PROTO_KEYS):
        r = C.FIRST_DATA_ROW + 2 * i
        protos[k] = [[copy(ws.cell(r + d, c)._style) for c in range(1, ncol + 1)] for d in (0, 1)]
    ws.delete_rows(C.FIRST_DATA_ROW, ws.max_row)

    last_m = max(C.MIN_MRP_ROWS, len(mrp.rows) + 1)
    MR = lambda col: f"MRP!${col}$2:${col}${last_m}"                     # noqa: E731
    HOL = f"{C.HOLIDAY_SHEET}!$A$2:$A${1 + max(C.HOLIDAY_RANGE_ROWS, n_hol)}"
    X0, XN = CL(C.DAY_COL0), CL(C.DAY_COL0 + C.N_DAYS - 1)              # X, AZ
    AW = CL(C.DAY_COL0 + C.T_BALANCE_OFFSET)

    rows = order_rows(master)
    end = C.FIRST_DATA_ROW + 2 * len(rows) - 1
    for i, m in enumerate(rows):
        r, g = C.FIRST_DATA_ROW + 2 * i, C.FIRST_DATA_ROW + 2 * i + 1
        src = m.get("출처") or None
        st = protos.get(src, protos[C.NEW_SOURCE])
        for d in (0, 1):
            for c in range(1, ncol + 1):
                ws.cell(r + d, c)._style = copy(st[d][c - 1])
        moq = m.get("MOQ")
        vals = {1: src, 2: "MRP계획", 4: m.get("BPA") if m.get("BPA") != "" else None,
                5: m.get("공급자") or None, 6: m["품번"], 7: m.get("규격") or None,
                8: m.get("구분") or None, 10: moq if moq != "" else None}
        for c, v in vals.items():
            _txt(ws.cell(r, c, v))
        mrow = mrp.item(m["품번"])
        days = [_num(v) for v in mrow[C.N_FIXED:]] if mrow is not None else [0] * C.N_DAYS
        ws.cell(r, 9).value = _remark(m.get("비고"), is_non_mass(days, ignore_tail=params.get("ignore_tail", NON_MASS_IGNORE_TAIL)))
        status = f'=IF($U{g}>=0,"{C.STATUS_OK}",IF(($O{r}+$N{r}/2)>=-$U{g},"{C.STATUS_SHORT}","{C.STATUS_ORDER}"))'
        ws.cell(r, 3, status)
        ws.cell(r, 11, f"=IF($U{g}>=0,0,IF(($O{r}+$N{r}/2)>=-$U{g},0,IF(OR(NOT(ISNUMBER($J{r})),$J{r}=0),"
                       f"MAX(0,-$U{g}-($O{r}+$N{r}/2)),CEILING(MAX(0,-$U{g}-($O{r}+$N{r}/2)),$J{r}))))")
        first = (f"INDEX(${X0}$4:${XN}$4,MATCH(1,INDEX((${X0}{g}:${XN}{g}<0)*(${X0}$4:${XN}$4>=$C$3)"
                 f"*(${X0}$4:${XN}$4<=$E$3),0),0))")
        ws.cell(r, 12, f'=IF(N($K{r})>0,IFERROR(IF(WORKDAY({first},-{C.LEAD_DAYS},{HOL})<$C$3,'
                       f'WORKDAY($C$3,1,{HOL}),WORKDAY({first},-{C.LEAD_DAYS},{HOL})),""),"")')
        for c, col in ((13, "E"), (14, "G"), (15, "H"), (16, "C"), (17, "D")):
            ws.cell(r, c, f"=IFERROR(INDEX({MR(col)},$W{r}),0)")
        ws.cell(r, 18, f'=IFERROR(INDEX({MR("S")},$W{r}),"")')
        ws.cell(r, 19, f'=IFERROR(INDEX({MR("T")},$W{r}),"")')
        ws.cell(r, 20, f"=$P{r}-{AW}{g}")
        ws.cell(r, 23, f'=IFERROR(MATCH($F{r},{MR("A")},0),"")')
        for j in range(C.N_DAYS):
            col = CL(C.DAY_COL0 + j)
            mcol = CL(C.N_FIXED + 1 + j)
            ws.cell(r, C.DAY_COL0 + j, f"=IFERROR(INDEX({MR(mcol)},$W{r}),0)")
            prev = f"$P{r}" if j == 0 else f"{CL(C.DAY_COL0 + j - 1)}{g}"
            ws.cell(g, C.DAY_COL0 + j, f"={prev}-{col}{r}")
        # 과부족 행
        ws.cell(g, 1, src)
        ws.cell(g, 2, "과부족")
        ws.cell(g, 3, status)
        ws.cell(g, 6, f"=$F{r}")
        ws.cell(g, 7, f"=$G{r}")
        ws.cell(g, 21, f'=_xlfn.MINIFS({X0}{g}:{XN}{g},${X0}$4:${XN}$4,">="&$C$3,${X0}$4:${XN}$4,"<="&$E$3)')
        ws.cell(g, 22, f'=N{r}+Q{r}+M{r}+O{r}+P{r}-SUMIFS(${X0}{r}:${XN}{r},${X0}$4:${XN}$4,">="&$C$3,'
                       f'${X0}$4:${XN}$4,"<="&$E$3)')
        CU.apply_item(ws, r, CU.loads(m))        # 수기 수정사항(서식·메모·추가입력) 재적용

    # 상단 (요약 / 기간 / 날짜 헤더)
    ws["A2"] = (f'="총 품목: "&COUNTIF($B$5:$B${end},"MRP계획")&"건   |   🔴 기간 중 부족 예상: "'
                f'&COUNTIF($U$5:$U${end},"<0")&"건   |   🟡 발주 필요: "&SUMPRODUCT(ISNUMBER($K$5:$K${end})*($K$5:$K${end}>0))'
                f'&"건   |   출처: MRP {params.get("파일명", "")}"')
    ws["C3"] = "=Y4" if params["start"] == (mrp.dates[1] if len(mrp.dates) > 1 else None) else \
        dt.datetime.combine(params["start"], dt.time())
    ws["E3"] = dt.datetime.combine(params["end"], dt.time())
    ws["F3"] = (f'=IF($C$3>$E$3,"⚠ 시작일>종료일",IF(COUNTIFS(${X0}$4:${XN}$4,">="&$C$3,${X0}$4:${XN}$4,"<="&$E$3)=0,'
                f'"⚠ 날짜범위 밖",COUNTIFS(${X0}$4:${XN}$4,">="&$C$3,${X0}$4:${XN}$4,"<="&$E$3)&"일 반영"))')
    ws["C3"].number_format = ws["E3"].number_format = "yyyy-mm-dd"
    ws["L4"] = "부족일 D-3\n(주말·휴일제외)"
    ws["V4"] = "기간말\n과부족"
    for j in range(C.N_DAYS):
        c = ws.cell(4, C.DAY_COL0 + j, f"=MRP!{CL(C.N_FIXED + 1 + j)}$1")
        c.number_format = "mm/dd"          # 열 너비 8 에 맞춤 (요일 포함 시 ### 표시됨)

    # 조건부서식 (2행 패턴 하나로 깔끔하게 재구성)
    rng = lambda a, b: f"{a}{C.FIRST_DATA_ROW}:{b}{end}"                  # noqa: E731
    f0 = C.FIRST_DATA_ROW
    # ① 대체품목(규격에 N/A) → 품목 두 행 전체 빨간 배경. 가장 우선 적용(규격을 엑셀에서 고쳐도 자동 반영)
    ws.conditional_formatting.add(rng("A", "AZ"), FormulaRule(formula=[f'ISNUMBER(SEARCH("N/A",$G{f0}))'],
        font=Font(color="FF000000"), fill=PatternFill("solid", bgColor=NA_FILL)))
    ws.conditional_formatting.add(rng("K", "K"), CellIsRule(operator="greaterThan", formula=["0"],
        font=Font(color="FF7F6000", bold=True), fill=PatternFill("solid", bgColor="FFFFE699")))
    ws.conditional_formatting.add(rng("U", "U"), CellIsRule(operator="lessThan", formula=["0"],
        font=Font(color="FF7A0000", bold=True), fill=PatternFill("solid", bgColor="FFFF8A8A")))
    ws.conditional_formatting.add(rng(X0, XN), FormulaRule(formula=[f'AND($B{f0}="과부족",{X0}{f0}<0)'],
        font=Font(color="FF9C0006", bold=True), fill=PatternFill("solid", bgColor="FFFFC7CE")))
    ws.conditional_formatting.add(rng("A", "V"), FormulaRule(formula=[f'AND($B{f0}="과부족",$U{f0}<0)'],
        font=Font(color="FFC00000")))
    ws.conditional_formatting.add(rng("A", "V"), FormulaRule(formula=[f'AND($B{f0}="과부족",$U{f0}>=0)'],
        font=Font(color="FF0070C0")))

    ws.auto_filter.ref = f"A4:{CL(ncol)}{end}"
    ws.freeze_panes = "N5"
    ws.column_dimensions["R"].hidden = False          # 전체 품목 대상 → 담당자 열 표시
    CU.apply_layout(ws, CU.get_layout(master))        # 사용자가 바꾼 열너비·틀고정 우선
    return end


def build_workbook(template, mrp, master, params, eff_holidays, new_items, summary_, exc):
    """결과 xlsx bytes 반환."""
    wb = _load(template)
    write_mrp_sheet(wb, mrp)
    write_holiday_sheet(wb, eff_holidays)
    n_hol = sum(1 for h in eff_holidays if h["적용"])
    write_plan_sheet(wb, mrp, master, params, n_hol)
    write_new_items_sheet(wb, new_items)
    write_log_sheet(wb, mrp.checks, params, summary_, exc)
    wb.active = wb.sheetnames.index(C.PLAN_SHEET)
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == C.PLAN_SHEET
    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
