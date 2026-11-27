"""합성 데이터 단위 테스트 (실데이터 불필요 → GitHub Actions 에서 실행)."""
import sys
import datetime as dt
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from mrp_plan import config as C, engine as E, master as M, review as R  # noqa: E402
from mrp_plan.loader import MrpData  # noqa: E402

D0 = dt.date(2026, 12, 21)                     # 월요일
DATES = [D0 + dt.timedelta(days=i) for i in range(C.N_DAYS)]


def make_mrp(items):
    rows, index = [], {}
    for code, stock, michul, daeyeo, s210, po, days in items:
        r = [code, f"{code} 품명", stock, michul, daeyeo, stock + michul + daeyeo, s210, po] + [0] * 10 \
            + ["담당A", "자삽", "KR_공급", "지점A"] + days + [0] * (C.N_DAYS - len(days))
        index[code] = len(rows)
        rows.append(r)
    return MrpData(list(C.MRP_HEADERS) + DATES, rows, DATES, index, {"B1": 5000000001})


def test_workday_matches_excel():
    sat = dt.date(2026, 12, 26)
    assert E.workday(sat, -3) == dt.date(2026, 12, 23)                         # 토 → 수
    assert E.workday(dt.date(2026, 12, 28), -3, [dt.date(2026, 12, 25)]) == dt.date(2026, 12, 22)
    assert E.workday(dt.date(2026, 12, 24), 1, [dt.date(2026, 12, 25)]) == dt.date(2026, 12, 28)


def test_holiday_with_production_is_not_applied():
    days = [10] * 5                                   # 12/21~12/25 매일 생산 → 12/25 생산계획 있음
    mrp = make_mrp([("A1", 100, 0, 0, 0, 0, days)])
    eff = E.effective_holidays([(dt.date(2026, 12, 25), "성탄절"), (dt.date(2027, 1, 1), "신정")], mrp)
    assert eff[0]["적용"] is False and eff[0]["소요합계"] == 10
    assert eff[1]["적용"] is True                       # 1/1 은 소요 0 → 휴일


def test_compute_order_qty_moq_and_v():
    # 재고 100, 매일 30 소요 → 4일째 -20. 발주잔량 0, 210=10(→5 인정)
    mrp = make_mrp([("A1", 100, -5, 2, 10, 0, [30, 30, 30, 30, 30])])
    master = [{"품번": "A1", "출처": "신규", "MOQ": 50, "순번": 1}, {"품번": "ZZ", "출처": "신규", "MOQ": "", "순번": 2}]
    start, end = DATES[0], DATES[6]
    res = E.compute(mrp, master, start, end, [])
    a, z = res
    assert a["기간중최저값"] == -50                     # 100-150
    assert a["상태"] == C.STATUS_ORDER
    assert a["발주수량"] == 50                          # 50-5=45 → MOQ 50 올림
    assert a["발주일"] == E.workday(DATES[3], -3) or a["발주일"] == E.workday(start, 1)
    assert a["기간말과부족"] == 10 + (-5) + 2 + 0 + 100 - 150
    assert z["MRP행"] == "" and z["발주수량"] == 0 and z["상태"] == C.STATUS_OK


def test_review_roundtrip(tmp_path):
    mrp = make_mrp([("B1", 0, 0, 0, 0, 0, [5]), ("B2", 0, 0, 0, 0, 0, [1])])
    items = M.find_new_items(mrp, [], DATES[0], DATES[5])
    assert items[0]["BPA"] == 5000000001 or items[1]["BPA"] == 5000000001
    p = tmp_path / "rv.xlsx"
    R.write_review_xlsx(items, p)
    import openpyxl
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    for r in range(3, ws.max_row + 1):
        if ws.cell(r, 1).value == "B2":
            ws.cell(r, 2).value = "X"
        else:
            ws.cell(r, 9).value = 3000                   # MOQ
    wb.save(p)
    got, changed = R.read_review_xlsx(p, items)
    by = {g["품번"]: g for g in got}
    assert by["B2"]["포함"] is False and by["B1"]["MOQ"] == 3000 and changed == 2
    new = M.add_items([], got, dt.date(2026, 11, 24))
    assert [n["품번"] for n in new] == ["B1"] and new[0]["등록일"] == "2026-11-24"
    assert M.from_csv(M.to_csv(new))[0]["MOQ"] == 3000


# ── 2026-11-27: 대체품목 / 양산품목X / 수기 수정 보존 ───────────────
def test_replaced_and_non_mass():
    from mrp_plan import custom as CU
    assert E.is_replaced("N/A SMW250-03(WH)") and E.is_replaced("abc n/a") and not E.is_replaced("NA 123")
    assert E.is_non_mass([1] + [0] * 8 + [2]) and not E.is_non_mass([1] + [0] * 7 + [2])
    tail = [5, 0, 5] + [0] * 10
    assert not E.is_non_mass(tail) and E.is_non_mass(tail, ignore_tail=False)   # 기본: 끝부분 0 제외
    assert E.is_non_mass([0] * 29, ignore_tail=True)           # 전체 0 은 항상 양산품목X
    assert CU.strip_mass_tag("로봇 / 양산품목X") == "로봇" and CU.strip_mass_tag("양산품목X") == ""


def test_custom_apply_translates_formula_and_comment():
    import openpyxl
    from mrp_plan import custom as CU
    wb = openpyxl.Workbook()
    ws = wb.active
    cu = {"x": {"0V": {"v": "=SUM($N2117:$P2117)-SUM($Z2117:$AD2117)", "o": "V2117"}},
          "m": {"1K": {"t": "메모", "a": "구매"}},
          "s": {"0E": {"fill": ["solid", ["r", "FFFF0000"]]}, "0J": {"nf": "#,##0"}}}
    CU.apply_item(ws, 9, CU.loads(CU.dumps(cu)))
    assert ws["V9"].value == "=SUM($N9:$P9)-SUM($Z9:$AD9)"
    assert ws["K10"].comment.text == "메모"
    assert ws["E9"].fill.fgColor.rgb == "FFFF0000" and ws["J9"].number_format == "#,##0"
