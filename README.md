# MRP → 구매계획_통합 자동화

매일 받는 `MRP_yymmdd.xlsx` 하나만 올리면 `구매계획_통합_yymmdd.xlsx` 를 바로 내려받습니다.
기존에 손으로 하던 **① MRP 시트 값 복사 → 통합파일에 붙여넣기 ② 없는 품목 수기 추가** 를 없앱니다.

- **웹 도구** (GitHub Pages): 브라우저 안에서 Python(Pyodide)으로 처리 → **파일이 PC 밖으로 나가지 않음**
- **노트북** (`notebooks/01_MRP_구매계획_자동화.ipynb`): 같은 코드(`src/mrp_plan`)를 단계별로 실행·검증

## 매일 사용법 (웹)

| 단계 | 화면 | 하는 일 |
|---|---|---|
| 0 | 품목 마스터 | 처음 1회만: 기존 `MRP_구매계획_통합_…xlsx` 또는 `item_master.csv` 업로드 |
| 1 | MRP 파일 | `MRP_yymmdd.xlsx` 끌어놓기 → 자동 검증(✅/⚠/⛔) |
| 2 | 기간·휴일 | 기본 시작일 = MRP 둘째 날(D+1), 종료일 = +7일. 필요 시 변경 |
| 3 | **신규 품목 검토** 👤 | 마스터에 없는 품목 + 제안값. MOQ·구분 등 채우기, 일괄 변경, 포함/제외 |
| 4 | 결과 미리보기 | KPI + 발주필요 목록(발주일 순) |
| 5 | 승인·다운로드 | 신규 품목을 마스터에 저장 + 엑셀 다운로드 (+ 마스터 CSV 백업) |

> 마스터는 **이 브라우저의 저장소**(localStorage)에 보관됩니다. 브라우저 데이터 삭제·다른 PC 사용에 대비해
> 5단계에서 함께 받는 `item_master_YYYY-MM-DD.csv` 를 보관하세요. 다른 PC에서는 0단계에서 그 CSV를 올리면 됩니다.

## 결과 엑셀 구성

| 시트 | 내용 |
|---|---|
| 구매계획_통합 | 기존과 같은 서식·수식(품목당 MRP계획/과부족 2행). **C3/E3 기간을 바꾸면 엑셀에서 재계산** |
| MRP | 업로드한 MRP 시트를 값으로 변환한 것 (수식 참조 대상) |
| 휴일 | 공휴일 목록과 적용 여부. A열이 '부족일 D-3' WORKDAY 휴일 범위 |
| 신규품목 | 오늘 추가(또는 제외)한 품목과 제안값·근거 |
| 검증로그 | 입력 검증 결과, 기간, 예외(MRP 미존재, BPA 불일치, 담당자 #N/A) |

## 계산 규칙

| 항목 | 정의 |
|---|---|
| 과부족(일별) | 재고 − 일별 소요량 누계 |
| 기간중 최저값 U | C3~E3 기간 안의 과부족 최솟값 |
| 상태 C | U≥0 🟢정상 / (발주잔량+210/2) ≥ 부족량 🔴부족 / 그 외 🟡발주필요 |
| 발주수량 K | 🟡일 때 `부족량 − (발주잔량 + 210/2)` 를 MOQ 배수로 올림 (MOQ 빈칸·0 → 올림 없음) |
| 부족일 D-3 L | 기간 내 첫 마이너스일의 3영업일 전. 시작일보다 앞이면 시작일 다음 영업일. **주말 + 휴일 시트 제외** |
| 휴일 판단 | 공휴일이라도 **그날 생산계획(전체 소요량 합계>0)이 있으면 근무일** |
| 당월 소요량 T | 재고 − AW열(과부족행) (기존 정의 유지) |
| **기간말 과부족 V** | **210 + 미출 + 대여 + 발주잔량 + 재고 − Σ소요량(C3~E3)** ← 2026-11-24 변경 |
| MRP에 없는 품목 | 마스터에 있으면 계속 0으로 표시 |

## 배포 (최초 1회)

1. GitHub에 새 저장소를 만들고 이 폴더 내용을 push (`work/`·실데이터는 `.gitignore` 로 제외됨)
2. 저장소 **Settings → Pages → Build and deployment → Source: GitHub Actions**
3. `main` 에 push 하면 `.github/workflows/pages.yml` 이 테스트 → 사이트 조립 → 배포
4. 주소: `https://<계정>.github.io/<저장소>/`

- 무료 계정의 GitHub Pages는 **공개 저장소**에서만 됩니다. 코드·템플릿(서식만, 데이터 없음)은 공개되지만 업무 파일은 저장소에 올라가지 않습니다.
- Pyodide는 사이트에 함께 올라갑니다(`vendor/`). 사내망이 `github.io` 를 막으면 로컬에서 `bash tools/build_site.sh && cd _site && python -m http.server 8000` 후 `http://localhost:8000` 으로 사용하세요.

## 노트북 사용법

```bash
pip install -r requirements.txt jupyter
mkdir -p work/input   # MRP_yymmdd.xlsx (최초엔 기존 통합파일도) 넣기
jupyter lab notebooks/01_MRP_구매계획_자동화.ipynb
```
1번 셀의 `MRP_FILE` 만 바꿔서 1~5 실행 → `work/output/신규품목_검토_yymmdd.xlsx` 노란 칸 수정·저장 → 6~8 실행.
9번 셀은 LibreOffice가 있으면 엑셀 수식 결과와 Python 엔진 결과를 전 품목 비교합니다.

## 폴더 구조

```
src/mrp_plan/   loader(로드·값변환·검증) · master(마스터·신규감지) · engine(계산) · writer(엑셀) · review(노트북 검토파일)
web/            index.html · app.js(UI) · worker.js(Pyodide) · bridge.py · assets/(template.xlsx, holidays.csv)
notebooks/      01_MRP_구매계획_자동화.ipynb
tools/          build_template.py(기존 파일→서식 템플릿) · build_site.sh · make_notebook.py
tests/          test_engine.py(합성 데이터) · golden_check.py(실데이터 대조, 로컬 전용)
```

## 유지보수

- **공휴일**: `web/assets/holidays.csv` 에 행 추가 (2027년 대체공휴일 일부는 정부 고시 후 확인)
- **서식 변경**: 엑셀에서 기존 통합파일 서식을 고친 뒤 `python tools/build_template.py <파일> web/assets/template.xlsx`
- **노트북 셀 수정**: `tools/make_notebook.py` 수정 후 실행
- **검증**: `python tests/golden_check.py <MRP파일> <기존 통합파일>` → 기존 엑셀 계산값과 엔진 비교
