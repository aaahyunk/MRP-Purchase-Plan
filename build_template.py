"""기존 'MRP_구매계획_통합' 파일에서 서식 템플릿을 추출한다 (데이터는 모두 제거).

사용법:  python tools/build_template.py <기존_통합파일.xlsx> web/assets/template.xlsx

- '구매계획_통합' 1~4행(제목/요약/기간/헤더) 서식과 열 너비, 데이터 유효성, 병합을 유지
- 출처별 서식 견본(MRP계획행+과부족행) 1쌍씩을 5행부터 배치 (config.PROTO_KEYS 순서)
- 'MRP' 시트는 1행(헤더)과 2행(데이터 서식 견본)만 남김
- 그 외 시트, 조건부서식, 값은 모두 제거 → 템플릿에는 업무 데이터가 남지 않는다
"""
import sys
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from mrp_plan import config as C  # noqa: E402


def build(src, dst):
    wb = openpyxl.load_workbook(src)
    for name in list(wb.sheetnames):
        if name not in (C.PLAN_SHEET, C.MRP_SHEET):
            del wb[name]
    ws = wb[C.PLAN_SHEET]

    # 출처별 첫 행쌍 찾기
    first = {}
    for r in range(C.FIRST_DATA_ROW, ws.max_row + 1, 2):
        if ws.cell(r, 2).value != "MRP계획":
            continue
        a = ws.cell(r, 1).value
        first.setdefault(a, r)
    missing = [k for k in C.PROTO_KEYS if k not in first]
    if missing:
        raise SystemExit(f"서식 견본을 찾지 못한 출처: {missing}")

    ncol = C.N_PLAN_COLS
    styles = []
    for k in C.PROTO_KEYS:
        r = first[k]
        styles.append([[copy(ws.cell(r + d, c)._style) for c in range(1, ncol + 1)] for d in (0, 1)])
    heights = ws.row_dimensions[C.FIRST_DATA_ROW].height

    # 5행 이후 삭제 후 견본 재배치
    ws.delete_rows(C.FIRST_DATA_ROW, ws.max_row)
    for i, pair in enumerate(styles):
        for d in (0, 1):
            r = C.FIRST_DATA_ROW + 2 * i + d
            for c in range(1, ncol + 1):
                cell = ws.cell(r, c)
                cell._style = pair[d][c - 1]
                cell.value = None
            ws.row_dimensions[r].height = heights
    ws.conditional_formatting = ConditionalFormattingList()
    ws.auto_filter.ref = None
    ws["A2"] = None
    for c in range(1, ncol + 1):          # 헤더 날짜 수식 / 기간 셀 값 제거
        if ws.cell(4, c).value and str(ws.cell(4, c).value).startswith("="):
            ws.cell(4, c).value = None
    for ref in ("C3", "E3", "F3"):
        ws[ref] = None

    m = wb[C.MRP_SHEET]
    if m.max_row > 2:
        m.delete_rows(3, m.max_row)
    for c in range(1, m.max_column + 1):
        m.cell(2, c).value = None
        if c > C.N_FIXED:
            m.cell(1, c).value = None
    m.conditional_formatting = ConditionalFormattingList()
    m.auto_filter.ref = None
    wb.active = wb.sheetnames.index(C.PLAN_SHEET)
    wb.save(dst)
    print(f"template saved → {dst}  (견본 {len(styles)}쌍)")


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
