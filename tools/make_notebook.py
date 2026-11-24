"""notebooks/01_MRP_구매계획_자동화.ipynb 생성 스크립트 (셀 내용 관리용)."""
import nbformat as nbf
from pathlib import Path

md = nbf.v4.new_markdown_cell
code = nbf.v4.new_code_cell
cells = [
md("""# MRP → 구매계획_통합 자동화

매일 받는 **`MRP_yymmdd.xlsx`** 하나만 넣으면 **`구매계획_통합_yymmdd.xlsx`** 를 만들어 줍니다.

| 단계 | 내용 | 사람 확인(HITL) |
|---|---|---|
| 1 | 설정 (파일 경로) | 파일명만 바꾸기 |
| 2 | MRP 로드 · 값 변환 · 검증 | ⚠/⛔ 항목 확인 |
| 3 | 품목 마스터 불러오기 (최초 1회는 기존 통합파일에서 추출) | |
| 4 | 기간 · 휴일 | 기간 바꾸려면 START/END |
| 5 | 신규 품목 감지 → 검토 파일 생성 | 👤 **검토 파일 수정·저장** |
| 6 | 검토 결과 반영 → 마스터 갱신 | |
| 7 | 계산 · 미리보기 | 발주필요 목록 확인 |
| 8 | 엑셀 생성 | |
| 9 | (선택) 엑셀 수식 ↔ 엔진 교차검증 | |

> **처음 사용**: `work/input/` 에 MRP 파일과 기존 `MRP_구매계획_통합_...xlsx` 를 넣고 위에서부터 실행합니다.
> **매일**: 1번 셀의 `MRP_FILE` 만 바꾸고 1~5 실행 → 검토 파일 수정 → 6~8 실행.
> ⚠ *Run All* 로 한 번에 돌리면 신규 품목은 **제안값 그대로 승인**됩니다."""),

md("## 1. 설정"),
code("""from pathlib import Path
import sys, datetime as dt

ROOT = next(p for p in [Path.cwd(), *Path.cwd().parents] if (p / "src" / "mrp_plan").exists())
sys.path.insert(0, str(ROOT / "src"))

WORK = ROOT / "work"                                  # ← 업무 데이터 폴더 (.gitignore 로 커밋 제외)
MRP_FILE   = WORK / "input" / "MRP_261123.xlsx"       # ← 매일 이 파일명만 바꾸세요
PLAN_FILE  = WORK / "input" / "MRP_구매계획_통합_업데이트1124.xlsx"   # 최초 마스터 추출용 (마스터가 있으면 안 씀)
MASTER_CSV = WORK / "master" / "item_master.csv"
OUT_DIR    = WORK / "output"
TEMPLATE   = ROOT / "web" / "assets" / "template.xlsx"
HOLIDAYS   = ROOT / "web" / "assets" / "holidays.csv"

START = None   # 예: dt.date(2026, 11, 24)  /  None 이면 MRP 둘째 날짜(D+1)
END   = None   # 예: dt.date(2026, 12, 1)   /  None 이면 시작일 + 7일

for d in (WORK / "master", OUT_DIR):
    d.mkdir(parents=True, exist_ok=True)
print("ROOT:", ROOT); print("MRP :", MRP_FILE.name, "존재" if MRP_FILE.exists() else "❌ 없음")"""),

md("## 2. MRP 로드 · 값 변환 · 검증\n엑셀에 저장된 마지막 계산값을 읽습니다 (= 'MRP' 시트 전체 값 복사)."),
code("""import pandas as pd
from mrp_plan import loader, master as M, engine as E, writer as W, review as R, config as C
pd.set_option("display.max_columns", 40); pd.set_option("display.width", 200)

mrp = loader.load_mrp(str(MRP_FILE))
ICON = {"ok": "✅", "info": "ℹ️", "warn": "⚠️", "error": "⛔"}
display(pd.DataFrame([{"": ICON[c.level], "항목": c.title, "내용": c.detail} for c in mrp.checks]))
assert not any(c.level == "error" for c in mrp.checks), "⛔ 오류가 있습니다. 위 내용을 확인하세요."
mrp_df = pd.DataFrame(mrp.rows, columns=[str(h) if not isinstance(h, dt.date) else h.strftime("%m/%d") for h in mrp.header])
mrp_df.head(3)"""),

md("## 3. 품목 마스터\n수기로 관리하던 정보(출처·BPA·공급자·규격·구분·비고·MOQ)를 `item_master.csv` 로 따로 관리합니다.\n파일이 없으면 기존 통합파일에서 **1회 추출**합니다 (중복 품번은 첫 행 유지)."),
code("""if MASTER_CSV.exists():
    master = M.load_csv(MASTER_CSV)
    print(f"마스터 로드: {len(master):,}개  ({MASTER_CSV.name})")
else:
    master, rep = M.extract_from_plan(str(PLAN_FILE))
    M.save_csv(master, MASTER_CSV)
    print(f"기존 통합파일에서 마스터 추출: {rep['items']:,}개 → {MASTER_CSV}")
    if rep["duplicates"]:
        print(f"⚠ 중복 품번 {len(rep['duplicates'])}건 (첫 행 유지)"); display(pd.DataFrame(rep["duplicates"]))
    if rep["cleared"]:
        print(f"⚠ #N/A 값 {len(rep['cleared'])}건 → 빈칸 처리"); display(pd.DataFrame(rep["cleared"]))
pd.DataFrame(master).head()"""),

md("## 4. 기간 · 휴일\n- 최저값 기간 기본값: 시작일 = MRP 둘째 날짜(D+1), 종료일 = 시작일+7일 (결과 엑셀 C3/E3 에서 다시 바꿀 수 있음)\n- 공휴일: **그날 생산계획(전체 소요량 합계>0)이 있으면 휴일로 보지 않음**"),
code("""s0, e0 = E.default_period(mrp)
start, end = START or s0, END or e0
hol = E.load_holidays(HOLIDAYS.read_text(encoding="utf-8"))
eff = E.effective_holidays(hol, mrp)
applied = [h for h in eff if h["적용"]]
print(f"기간: {start} ~ {end}")
display(pd.DataFrame([h for h in eff if mrp.dates[0] <= h["날짜"] <= mrp.dates[-1] + dt.timedelta(days=30)]
                     or [{"안내": "MRP 기간 근처 공휴일 없음"}]))"""),

md("## 5. 신규 품목 감지 → 👤 검토 파일\nMRP에는 있는데 마스터에 없는 품목입니다. 제안값: **BPA**=단가시트, **공급자**=지점, **규격**=품명, **출처**='신규'.\n부족 예상 품목이 위에 정렬됩니다."),
code("""new_items = M.find_new_items(mrp, master, start, end)
REVIEW = OUT_DIR / f"신규품목_검토_{mrp.dates[0]:%y%m%d}.xlsx"
print(f"신규 품목 {len(new_items):,}개  (BPA 자동 {sum(1 for n in new_items if n['BPA'] != ''):,}개, "
      f"기간 중 부족 {sum(1 for n in new_items if n['기간최저'] < 0):,}개)")
if new_items:
    nd = pd.DataFrame(new_items)
    display(nd.groupby("담당자").agg(품목수=("품번", "size"), 부족예상=("기간최저", lambda s: (s < 0).sum())))
    if not REVIEW.exists():
        R.write_review_xlsx(new_items, REVIEW)
    print(f"👤 검토 파일: {REVIEW}\\n   노란 칸(포함/출처/BPA/공급자/규격/구분/비고/MOQ)을 수정·저장한 뒤 다음 셀을 실행하세요.")
    display(nd.head(10))"""),

md("## 6. 검토 결과 반영 → 마스터 갱신\n저장된 검토 파일을 읽어 `포함=O` 품목만 마스터에 추가합니다. 기존 마스터는 `item_master_백업_*.csv` 로 보관합니다."),
code("""if new_items:
    reviewed, changed = R.read_review_xlsx(REVIEW, new_items)
    take = [r for r in reviewed if r["포함"]]
    print(f"검토 반영: 수정 {changed}칸, 포함 {len(take):,}개 / 제외 {len(reviewed) - len(take):,}개")
    backup = MASTER_CSV.with_name(f"item_master_백업_{dt.datetime.now():%y%m%d_%H%M}.csv")
    backup.write_text(M.to_csv(master), encoding="utf-8")
    master = M.add_items(master, take)
    M.save_csv(master, MASTER_CSV)
    print(f"마스터 저장: {len(master):,}개 (백업 {backup.name})")
else:
    reviewed = []
    print("신규 품목 없음")"""),

md("## 7. 계산 · 미리보기"),
code("""result = E.compute(mrp, master, start, end, applied)
summ = E.summary(result)
exc = M.exceptions(mrp, master)
display(pd.DataFrame([summ]))
order = pd.DataFrame([{k: r[k] for k in ["상태", "품번", "규격", "공급자", "담당자", "MOQ", "발주수량", "발주일",
                                         "기간중최저값", "발주잔량", "210", "기간말과부족"]} for r in result if r["발주수량"] > 0])
print(f"🟡 발주필요 {len(order):,}건 (발주일 순)")
order.sort_values(["발주일", "담당자"]).head(30) if len(order) else order"""),

md("## 8. 엑셀 생성\n`구매계획_통합` 시트는 기존과 같은 수식이 들어갑니다 → 엑셀에서 C3/E3 을 바꾸면 그대로 재계산됩니다."),
code("""params = {"start": start, "end": end, "파일명": MRP_FILE.name, "생성시각": f"{dt.datetime.now():%Y-%m-%d %H:%M}"}
xbytes = W.build_workbook(TEMPLATE.read_bytes(), mrp, master, params, eff, reviewed, summ, exc)
OUT = OUT_DIR / f"구매계획_통합_{mrp.dates[0]:%y%m%d}.xlsx"
OUT.write_bytes(xbytes)
print(f"✅ 저장: {OUT}  ({len(xbytes)/1e6:.1f} MB)")"""),

md("## 9. (선택) 교차검증 — 엑셀 수식 결과 = Python 엔진 결과?\nLibreOffice 가 설치된 PC에서만 실행됩니다. 수식을 실제로 계산해 엔진 결과와 전 품목을 비교합니다."),
code("""import shutil, subprocess, tempfile, openpyxl
soffice = shutil.which("soffice") or shutil.which("libreoffice")
if not soffice:
    print("LibreOffice 없음 → 건너뜀")
else:
    tmp = Path(tempfile.mkdtemp())
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp), str(OUT)],
                   check=True, capture_output=True, timeout=900)
    V = openpyxl.load_workbook(tmp / OUT.name, data_only=True)[C.PLAN_SHEET]
    bad = []
    for i, r in enumerate(result):
        a, g = 5 + 2 * i, 6 + 2 * i
        L = V.cell(a, 12).value
        L = L.date() if isinstance(L, dt.datetime) else None
        ok = (abs((V.cell(a, 11).value or 0) - r["발주수량"]) < 1e-6 and L == r["발주일"] and V.cell(a, 3).value == r["상태"]
              and abs((V.cell(g, 21).value or 0) - r["기간중최저값"]) < 1e-6 and abs((V.cell(g, 22).value or 0) - r["기간말과부족"]) < 1e-6)
        if not ok:
            bad.append(r["품번"])
    print(f"교차검증: {len(result) - len(bad):,}/{len(result):,} 일치", "✅" if not bad else f"❌ {bad[:10]}")"""),
]

nb = nbf.v4.new_notebook(cells=cells, metadata={"kernelspec": {"name": "python3", "display_name": "Python 3", "language": "python"},
                                                "language_info": {"name": "python"}})
out = Path(__file__).resolve().parents[1] / "notebooks" / "01_MRP_구매계획_자동화.ipynb"
nbf.write(nb, out)
print("written", out)
