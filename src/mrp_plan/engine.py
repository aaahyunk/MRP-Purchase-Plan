"""계산 엔진: '구매계획_통합' 시트의 수식을 Python으로 동일하게 재현.

엑셀 파일에는 수식을 그대로 넣고(사용자가 C3/E3을 바꾸면 재계산), 이 엔진은
  ① 웹 미리보기  ② 엑셀 수식 결과와의 교차검증  에 사용한다.

수식 정의 (품목 1개 = MRP계획행 r + 과부족행 r+1)
  잔량[d]     = 재고 - Σ(일별소요 ~d)                         (과부족행 X~AZ)
  U 기간중최저 = MIN(잔량[d]  | C3 ≤ d ≤ E3)
  C 상태      = U≥0 → 🟢정상 / (발주잔량+210/2) ≥ -U → 🔴부족 / 그 외 🟡발주필요
  K 발주수량   = 🟡일 때 CEILING(-U-(발주잔량+210/2), MOQ)   (MOQ 없거나 0 → 올림 없음)
  L 발주일     = K>0 일 때 WORKDAY(첫 마이너스일, -3, 휴일). 시작일보다 앞이면 WORKDAY(C3, 1, 휴일)
  T 당월소요량 = 재고 - 잔량[AW]                               (기존 정의 유지)
  V 기간말과부족 = 210 + 미출 + 대여 + 발주잔량 + 재고 - Σ(일별소요 | C3 ≤ d ≤ E3)   ← 2026-11-24 변경
"""
import csv
import io
import math
import datetime as dt

from . import config as C
from .loader import _num


# ── 휴일 ────────────────────────────────────────────────────────
def load_holidays(src):
    """CSV(날짜,명칭) → [(date, name)]"""
    if isinstance(src, (bytes, bytearray)):
        src = src.decode("utf-8-sig")
    if "\n" not in src and src.endswith(".csv"):
        with open(src, encoding="utf-8-sig") as f:
            src = f.read()
    out = []
    for r in csv.DictReader(io.StringIO(src.lstrip("﻿"))):
        try:
            d = dt.date.fromisoformat((r.get("날짜") or "").strip())
        except ValueError:
            continue
        out.append((d, (r.get("명칭") or "").strip()))
    return sorted(out)


def effective_holidays(holidays, mrp):
    """생산계획(그날 전체 소요량>0)이 있는 공휴일은 휴일로 보지 않는다.
    반환: [{날짜, 명칭, 소요합계, 적용, 사유}]"""
    tot = dict(zip(mrp.dates, mrp.day_totals())) if mrp.dates else {}
    out = []
    for d, name in holidays:
        if d in tot:
            s = tot[d]
            applied = s <= 0
            why = "생산계획 없음 → 휴일 반영" if applied else "생산계획 있음 → 휴일 미반영(근무일)"
        else:
            s, applied, why = None, True, "MRP 기간 밖 → 휴일 반영"
        out.append({"날짜": d, "명칭": name, "소요합계": s, "적용": applied, "사유": why})
    return out


def workday(start, n, holidays=()):
    """엑셀 WORKDAY(start, n, holidays)와 동일 (토/일 + 휴일 제외, start 자신은 세지 않음)."""
    hs = set(holidays)
    d, step, left = start, (1 if n > 0 else -1), abs(n)
    while left:
        d += dt.timedelta(days=step)
        if d.weekday() < 5 and d not in hs:
            left -= 1
    return d


# ── 계산 ────────────────────────────────────────────────────────
def default_period(mrp):
    """기존 파일과 동일: 시작일 = MRP 둘째 날짜(Y4, D+1), 종료일 = 시작일 + 7일."""
    s = mrp.dates[1] if len(mrp.dates) > 1 else mrp.dates[0]
    return s, s + dt.timedelta(days=7)


def order_rows(master):
    """출처 블록 순서 → 순번 순서 (기존 파일 배치 유지)."""
    rank = {k: i for i, k in enumerate(C.SOURCE_ORDER)}
    other = rank[C.NEW_SOURCE] + 0.5            # 목록에 없는 출처는 '신규' 다음

    def key(m):
        s = m.get("출처") or None
        return (rank.get(s, other), int(m.get("순번") or 0))
    return sorted(master, key=key)


def compute(mrp, master, start, end, holidays_applied=()):
    """마스터 순서대로 품목별 결과(dict) 리스트 반환."""
    hol = [h if isinstance(h, dt.date) else h["날짜"] for h in holidays_applied]
    dates = mrp.dates
    win = [start <= d <= end for d in dates]
    out = []
    for m in order_rows(master):
        row = mrp.item(m["품번"])
        if row is not None:
            P, Q, M = _num(row[C.I_STOCK]), _num(row[C.I_MICHUL]), _num(row[C.I_DAEYEO])
            N, O = _num(row[C.I_210]), _num(row[C.I_PO_REMAIN])
            days = [_num(v) for v in row[C.N_FIXED:C.N_FIXED + len(dates)]]
            owner, proc = row[C.I_OWNER], row[C.I_PROCESS]
            mrp_row = mrp.index[m["품번"]] + 1
        else:                                  # MRP에 없는 품목 → 0 표시
            P = Q = M = N = O = 0.0
            days = [0.0] * len(dates)
            owner, proc, mrp_row = "", "", ""
        bal, b = [], P
        for x in days:
            b -= x
            bal.append(b)
        inwin = [v for v, w in zip(bal, win) if w]
        U = min(inwin) if inwin else 0.0
        cover = O + N * C.HALF_210
        J = m.get("MOQ")
        J = J if isinstance(J, (int, float)) and not isinstance(J, bool) else None
        if U >= 0:
            status, K = C.STATUS_OK, 0
        elif cover >= -U:
            status, K = C.STATUS_SHORT, 0
        else:
            status = C.STATUS_ORDER
            need = max(0.0, -U - cover)
            K = need if not J else math.ceil(round(need / J, 9)) * J
        L = None
        if K > 0:
            first = next((d for v, d, w in zip(bal, dates, win) if v < 0 and w), None)
            if first:
                x = workday(first, -C.LEAD_DAYS, hol)
                L = workday(start, 1, hol) if x < start else x
        T = P - bal[C.T_BALANCE_OFFSET] if len(bal) > C.T_BALANCE_OFFSET else 0.0
        V = N + Q + M + O + P - sum(x for x, w in zip(days, win) if w)
        out.append({**m, "상태": status, "발주수량": K, "발주일": L, "대여": M, "210": N,
                    "발주잔량": O, "재고": P, "미출": Q, "담당자": owner, "공정": proc,
                    "당월소요량": T, "기간중최저값": U, "기간말과부족": V, "MRP행": mrp_row,
                    "일별소요": days, "잔량": bal})
    return out


def summary(result):
    return {
        "총품목": len(result),
        "부족예상": sum(1 for r in result if r["기간중최저값"] < 0),
        "발주필요": sum(1 for r in result if r["발주수량"] > 0),
        "잔량커버": sum(1 for r in result if r["상태"] == C.STATUS_SHORT),
        "MRP미존재": sum(1 for r in result if r["MRP행"] == ""),
    }
