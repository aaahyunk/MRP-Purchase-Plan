"""공통 상수 정의 (MRP 시트 / 구매계획_통합 시트 구조)."""

# ── 일일 MRP 파일의 'MRP' 시트 (A~AY, 51열) ─────────────────────────
MRP_SHEET = "MRP"
MRP_HEADERS = [
    "품목", "품명", "재고", "미출", "대여", "재고합계", "생산재고", "구매발주 잔량",
    "M-2", "M-1", "기간 소요량", "M+1", "M+2", "기간 과부족", "장기재고율", "재고일수",
    "단가", "현 재고금액", "담당자", "공정", "공급사", "지점",
]  # A~V (22열). W~AY(29열)은 일별 소요량, 1행은 날짜
N_FIXED = len(MRP_HEADERS)          # 22
N_DAYS = 29                         # W~AY
N_MRP_COLS = N_FIXED + N_DAYS       # 51

# MRP 시트 열 인덱스(0-base)
I_ITEM, I_NAME, I_STOCK, I_MICHUL, I_DAEYEO = 0, 1, 2, 3, 4
I_210, I_PO_REMAIN = 6, 7
I_OWNER, I_PROCESS, I_SUPPLIER, I_SITE = 18, 19, 20, 21

# ── 품목 마스터 (사람이 관리하는 정보) ─────────────────────────────
MASTER_FIELDS = ["품번", "출처", "BPA", "공급자", "규격", "구분", "비고", "MOQ", "순번", "등록일"]

# 구매계획_통합 시트에서 출처 블록 순서 (기존 파일 순서 유지)
SOURCE_ORDER = ["와이앤제이", "월발주", "Wafer", "R.CHIP", "Shoe,Tiiun", "가스켓",
                "Sheet", "SME", "신규", None]
# 템플릿의 프로토타입(서식 견본) 행쌍 순서. 5행부터 2행씩.
PROTO_KEYS = SOURCE_ORDER
NEW_SOURCE = "신규"
GUBUN_OPTIONS = ["양산", "SVC", "개발", "단종", "자재단종", "미사용"]

# ── 구매계획_통합 시트 ───────────────────────────────────────────
PLAN_SHEET = "구매계획_통합"
HOLIDAY_SHEET = "휴일"
FIRST_DATA_ROW = 5
N_PLAN_COLS = 52                     # A~AZ
DAY_COL0 = 24                        # X열 (1-base)
T_BALANCE_OFFSET = 25                # 당월 소요량 T = 재고 - AW(과부족행)  → X+25 = AW
LEAD_DAYS = 3                        # 부족일 D-3 (영업일)
HALF_210 = 0.5                       # 210(생산재고)는 1/2만 인정 (기존 수식 = N/2)
MIN_MRP_ROWS = 3000                  # INDEX/MATCH 참조 범위 최소 행
HOLIDAY_RANGE_ROWS = 200

STATUS_OK = "🟢 정상"
STATUS_SHORT = "🔴 부족"
STATUS_ORDER = "🟡 발주필요"
