"""기존 1124 통합파일의 엑셀 계산값과 엔진 결과 비교 (실데이터 필요, 저장소에는 포함 안 함).
사용: python tests/golden_check.py <MRP_yymmdd.xlsx> <기존_통합파일.xlsx>"""
import sys, datetime as dt, openpyxl
sys.path.insert(0, "src")
from mrp_plan import loader, master as M, engine as E

def main(mrp_path, plan_path):
    mrp = loader.load_mrp(mrp_path)
    ms, rep = M.extract_from_plan(plan_path)
    pv = openpyxl.load_workbook(plan_path, data_only=True)["구매계획_통합"]
    s = pv["C3"].value.date(); e = pv["E3"].value.date()
    res = {r["품번"]: r for r in E.compute(mrp, ms, s, e, [])}   # 기존 파일 = 휴일 미반영
    stat = {"K": 0, "L": 0, "C": 0, "U": 0, "T": 0}; bad = []; n = 0
    for r in range(5, pv.max_row + 1, 2):
        if pv.cell(r, 2).value != "MRP계획": continue
        it = loader.norm_code(pv.cell(r, 6).value); x = res[it]; n += 1
        L = pv.cell(r, 12).value; L = L.date() if isinstance(L, dt.datetime) else None
        chk = {"K": abs((pv.cell(r, 11).value or 0) - x["발주수량"]) < 1e-6, "L": L == x["발주일"],
               "C": pv.cell(r, 3).value == x["상태"], "U": abs((pv.cell(r+1, 21).value or 0) - x["기간중최저값"]) < 1e-6,
               "T": abs((pv.cell(r, 20).value or 0) - x["당월소요량"]) < 1e-6}
        for k, v in chk.items():
            stat[k] += v
            if not v: bad.append((r, it, k))
    print(f"비교 {n}행:", {k: f"{v}/{n}" for k, v in stat.items()}, "불일치:", bad)
    print("마스터", rep["items"], "중복", len(rep["duplicates"]), rep["duplicates"][:3])
    return bad

if __name__ == "__main__":
    main(*sys.argv[1:3])
