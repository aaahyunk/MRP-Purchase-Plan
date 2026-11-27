"""일일 'MRP_yymmdd.xlsx' 로드 + 값 변환 + 입력 검증.

엑셀에 저장된 '마지막 계산값'을 읽는다(= 'MRP' 시트 전체를 값 복사한 것과 같음).
"""
import io
import re
import datetime as dt
from dataclasses import dataclass, field

import openpyxl

from . import config as C


@dataclass
class Check:
    level: str      # "ok" | "info" | "warn" | "error"
    title: str
    detail: str = ""


@dataclass
class MrpData:
    header: list                 # 51개 헤더 (W~AY 는 date)
    rows: list                   # [[51개 값], ...]  품목 있는 행만
    dates: list                  # 29개 date
    index: dict                  # 품번 → rows 인덱스(0-base)
    bpa: dict                    # 단가 시트: 품번 → BPA No
    checks: list = field(default_factory=list)
    file_date: object = None     # 파일명 yymmdd

    def item(self, code):
        i = self.index.get(code)
        return None if i is None else self.rows[i]

    def day_totals(self):
        """날짜별 전체 품목 소요량 합계 (생산계획 유무 판단용)."""
        tot = [0.0] * len(self.dates)
        for r in self.rows:
            for j in range(len(self.dates)):
                tot[j] += _num(r[C.N_FIXED + j])
        return tot


def norm_code(v):
    """품번을 문자열로 통일 (53404702 / 53404702.0 / ' 53404702 ' → '53404702')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _num(v):
    if v is None or isinstance(v, str):
        return 0.0
    if isinstance(v, bool):
        return float(v)
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 30000 < v < 80000:   # 엑셀 일련번호
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    return None


def parse_file_date(filename):
    m = re.search(r"(\d{6})(?!\d)", filename or "")
    if not m:
        return None
    s = m.group(1)
    try:
        return dt.date(2000 + int(s[:2]), int(s[2:4]), int(s[4:6]))
    except ValueError:
        return None


def load_mrp(src, filename=None):
    """src: 파일 경로 또는 bytes. MrpData 반환. 치명적 오류는 checks에 error로 기록."""
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    if filename is None and isinstance(src, str):
        filename = src.replace("\\", "/").split("/")[-1]
    checks = []
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    if C.MRP_SHEET not in wb.sheetnames:
        wb.close()
        return MrpData([], [], [], {}, {}, [Check("error", "'MRP' 시트가 없습니다", f"시트 목록: {wb.sheetnames}")])

    ws = wb[C.MRP_SHEET]
    it = ws.iter_rows(min_row=1, max_col=C.N_MRP_COLS, values_only=True)
    header = list(next(it, []))
    header += [None] * (C.N_MRP_COLS - len(header))

    # 1) 헤더 구조
    bad = [(i, header[i], h) for i, h in enumerate(C.MRP_HEADERS) if str(header[i] or "").strip() != h]
    if bad:
        checks.append(Check("error", "MRP 헤더(A~V)가 예상과 다릅니다",
                            ", ".join(f"{openpyxl.utils.get_column_letter(i+1)}1='{a}'(예상 '{b}')" for i, a, b in bad[:5])))
    dates = [_as_date(v) for v in header[C.N_FIXED:]]
    if any(d is None for d in dates):
        checks.append(Check("error", "W1:AY1 날짜를 읽지 못했습니다",
                            "MRP 시트가 계산되지 않은 상태일 수 있습니다. 엑셀에서 열어 저장한 뒤 다시 올려주세요."))
    else:
        gaps = [i for i in range(1, len(dates)) if (dates[i] - dates[i-1]).days != 1]
        if gaps:
            checks.append(Check("warn", "날짜가 연속되지 않습니다", f"{len(gaps)}곳"))
        checks.append(Check("ok", "일별 소요량 기간", f"{dates[0]:%m/%d} ~ {dates[-1]:%m/%d} ({len(dates)}일)"))
    header[C.N_FIXED:] = dates

    # 2) 데이터 행 (품목이 비었거나 0인 행 = 수식만 있는 빈 행 → 제외)
    rows, index, dups, n_formula_blank = [], {}, [], 0
    err_cnt = {}
    for raw in it:
        raw = list(raw) + [None] * (C.N_MRP_COLS - len(raw))
        code = norm_code(raw[0])
        if code in ("", "0"):
            continue
        raw[0] = code
        for j, v in enumerate(raw):
            if isinstance(v, str) and v.startswith("#"):
                err_cnt[(j, v)] = err_cnt.get((j, v), 0) + 1
        if raw[C.I_STOCK] is None and all(v is None for v in raw[C.N_FIXED:]):
            n_formula_blank += 1
        if code in index:
            dups.append(code)
            continue
        index[code] = len(rows)
        rows.append(raw)

    if not rows:
        checks.append(Check("error", "품목 데이터가 없습니다",
                            "저장된 계산값이 없을 수 있습니다. 엑셀에서 열어 [저장] 후 다시 업로드하세요."))
    elif n_formula_blank > len(rows) * 0.5:
        checks.append(Check("error", "계산값이 비어 있습니다",
                            f"{n_formula_blank}행의 재고/소요량이 비어 있음 → 엑셀에서 열어 저장 후 업로드"))
    else:
        checks.append(Check("ok", "MRP 품목 수", f"{len(rows):,}개"))
    if dups:
        checks.append(Check("warn", "MRP 시트에 중복 품번", f"{len(dups)}건 (첫 행만 사용): {', '.join(dups[:5])}"))
    for (j, v), n in sorted(err_cnt.items()):
        name = header[j] if j < C.N_FIXED else "일별소요"
        checks.append(Check("warn", f"'{name}' 열 {v} {n}건",
                            "담당자/공정 매핑이 없는 지점·품목입니다. 결과에는 그대로 표시됩니다." if j in (C.I_OWNER, C.I_PROCESS)
                            else "숫자 계산에서는 0으로 처리됩니다."))

    # 3) TODAY 시트 품목 수 vs MRP 수식 범위(3000행) 점검
    if "TODAY" in wb.sheetnames:
        t = wb["TODAY"]
        n_today = sum(1 for (v,) in t.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True) if v not in (None, "", 0))
        if n_today > len(rows) + len(dups):
            checks.append(Check("warn", "TODAY 품목 수가 MRP보다 많습니다",
                                f"TODAY {n_today:,}개 / MRP {len(rows):,}개 → MRP 시트 수식 범위(3000행) 확인 필요"))

    # 4) 단가 시트 BPA (품번 → BPA No)
    bpa = {}
    if "단가" in wb.sheetnames:
        for r in wb["단가"].iter_rows(min_row=5, min_col=3, max_col=4, values_only=True):
            b, p = r
            p = norm_code(p)
            if p and b not in (None, "") and p not in bpa:
                bpa[p] = b

    wb.close()

    # 5) 파일명 날짜
    fdate = parse_file_date(filename)
    if fdate and dates and dates[0]:
        if fdate != dates[0]:
            checks.append(Check("warn", "파일명 날짜와 MRP 시작일이 다릅니다",
                                f"파일명 {fdate:%y%m%d} / W1 {dates[0]:%y%m%d}"))
        else:
            checks.append(Check("ok", "파일명 날짜 일치", f"{fdate:%Y-%m-%d}"))

    return MrpData(header, rows, dates, index, bpa, checks, fdate)
