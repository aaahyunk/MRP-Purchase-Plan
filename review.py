"""노트북용 HITL: 신규 품목 검토 파일(xlsx) 내보내기 / 읽어오기.

노란 칸(포함·출처·BPA·공급자·규격·구분·비고·MOQ)만 수정하고 저장하면 된다.
"""
import io

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as CL

from . import config as C
from .loader import norm_code

EDIT = ["포함", "출처", "BPA", "공급자", "규격", "구분", "비고", "MOQ"]
REF = ["담당자", "공정", "재고", "기간소요", "기간최저", "첫부족일", "BPA출처"]
COLS = ["품번"] + EDIT + REF


def write_review_xlsx(items, path=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "신규품목_검토"
    ws.append(["※ 노란 칸만 수정 후 저장하세요. 포함=X 이면 오늘 결과와 마스터에서 제외(내일 다시 표시). "
               "MOQ가 비어 있으면 발주수량은 MOQ 올림 없이 계산됩니다."])
    ws.append(COLS)
    hdr = Font(name="Pretendard", bold=True, color="FFFFFFFF")
    for j, k in enumerate(COLS, 1):
        c = ws.cell(2, j)
        c.font = hdr
        c.fill = PatternFill("solid", fgColor="FF1E2761" if k not in EDIT else "FFBF8F00")
        c.alignment = Alignment(horizontal="center")
    yellow = PatternFill("solid", fgColor="FFFFF2CC")
    for r in items:
        ws.append([r["품번"], "O" if r.get("포함", True) else "X"] + [r.get(k, "") for k in EDIT[1:]]
                  + [r.get(k, "") for k in REF])
    n = len(items) + 2
    for j, k in enumerate(COLS, 1):
        if k in EDIT:
            for i in range(3, n + 1):
                ws.cell(i, j).fill = yellow
    dv1 = DataValidation(type="list", formula1='"O,X"', allow_blank=False)
    dv2 = DataValidation(type="list", formula1='"' + ",".join(C.GUBUN_OPTIONS) + '"', allow_blank=True)
    dv3 = DataValidation(type="list", formula1='"' + ",".join(k for k in C.SOURCE_ORDER if k) + '"',
                         allow_blank=True, showErrorMessage=False)   # 목록 외 직접 입력 허용
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    ws.add_data_validation(dv3)
    dv1.add(f"B3:B{n}")
    dv2.add(f"{CL(COLS.index('구분') + 1)}3:{CL(COLS.index('구분') + 1)}{n}")
    dv3.add(f"C3:C{n}")
    for j, w in enumerate([14, 6, 10, 13, 16, 40, 9, 14, 8, 10, 8, 11, 11, 11, 11, 10], 1):
        ws.column_dimensions[CL(j)].width = w
    ws.freeze_panes = "B3"
    ws.auto_filter.ref = f"A2:{CL(len(COLS))}{n}"
    if path:
        wb.save(path)
        return path
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def read_review_xlsx(path, items):
    """검토 파일을 읽어 items(원본 제안 목록)에 사람이 수정한 값을 덮어쓴다."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    head = [c.value for c in ws[2]]
    by = {r["품번"]: dict(r) for r in items}
    changed = 0
    for row in ws.iter_rows(min_row=3, values_only=True):
        rec = dict(zip(head, row))
        code = norm_code(rec.get("품번"))
        if code not in by:
            continue
        cur = by[code]
        for k in EDIT:
            v = rec.get(k)
            v = "" if v is None else v
            if k == "포함":
                v = str(v).strip().upper() != "X"
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            if v != cur.get(k, ""):
                changed += 1
            cur[k] = v
    return list(by.values()), changed
