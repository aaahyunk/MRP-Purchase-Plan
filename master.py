"""품목 마스터: 사람이 관리하는 정보(출처·BPA·공급자·규격·구분·비고·MOQ).

- 최초 1회: 기존 '구매계획_통합' 시트에서 추출 (extract_from_plan)
- 매일: MRP에 새로 나온 품목을 찾아 제안값과 함께 검토 목록 생성 (find_new_items)
- 검토(HITL) 후 승인된 품목을 마스터에 추가 (add_items)
"""
import csv
import io
import datetime as dt

import openpyxl

from . import config as C
from .loader import norm_code, _num
from . import custom as CU

ERR = {"#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}


def _clean(v):
    if v is None:
        return ""
    if isinstance(v, str) and v.strip().startswith("#") and v.strip().upper() in ERR:
        return ""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return v


def extract_from_plan(src, template=None):
    """기존/수정한 통합 파일에서 마스터 추출. (master, report) 반환. 중복 품번은 첫 행을 유지.

    template(서식 템플릿 경로/bytes)을 주면 메모·하이라이트·추가 수식·열너비 같은
    수기 수정사항까지 읽어 '수기수정' 필드와 '#LAYOUT' 행에 저장한다.
    """
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    full = template is not None
    wb = openpyxl.load_workbook(src, read_only=not full, data_only=False)
    ws = wb[C.PLAN_SHEET]
    protos = None
    if full:
        if isinstance(template, (bytes, bytearray)):
            template = io.BytesIO(template)
        protos = CU.proto_keys(openpyxl.load_workbook(template))
    master, seen, dups, cleared = [], {}, [], []
    cols = {"출처": 0, "BPA": 3, "공급자": 4, "품번": 5, "규격": 6, "구분": 7, "비고": 8, "MOQ": 9}
    for r, row in enumerate(ws.iter_rows(min_row=C.FIRST_DATA_ROW, max_col=10, values_only=True), C.FIRST_DATA_ROW):
        if row[1] != "MRP계획":
            continue
        for k, i in cols.items():
            v = row[i]
            if isinstance(v, str) and v.strip().upper() in ERR:
                cleared.append({"품번": norm_code(row[5]), "항목": k, "원래값": v})
        rec = {k: _clean(row[i]) for k, i in cols.items()}
        rec["품번"] = norm_code(rec["품번"])
        if not rec["품번"]:
            continue
        rec["비고"] = CU.strip_mass_tag(rec["비고"])
        if rec["품번"] in seen:
            prev = seen[rec["품번"]]
            diff = [k for k in cols if k != "품번" and str(prev[k]) != str(rec[k])]
            dups.append({"품번": rec["품번"], "다른항목": ",".join(diff) or "-"})
            continue
        rec["순번"] = len(master) + 1
        rec["등록일"] = ""
        rec["수기수정"] = CU.dumps(CU.capture_item(ws, r, rec["출처"] or None, protos)) if full else ""
        seen[rec["품번"]] = rec
        master.append(rec)
    layout = CU.capture_layout(ws) if full else None
    wb.close()
    if layout:
        master.append(CU.layout_record(layout))
    return master, {"items": len(CU.items_only(master)), "duplicates": dups, "cleared": cleared,
                    "custom": CU.summarize(master) if full else None}


def diff_master(old, new):
    """마스터 교체 전후 비교 (추가/삭제/값변경/수기수정변경 품목 수)."""
    o = {m["품번"]: m for m in CU.items_only(old)}
    n = {m["품번"]: m for m in CU.items_only(new)}
    keys = ["출처", "BPA", "공급자", "규격", "구분", "비고", "MOQ"]
    changed = [k for k in n if k in o and any(str(o[k].get(f, "")) != str(n[k].get(f, "")) for f in keys)]
    cust = [k for k in n if k in o and CU.loads(o[k]) != CU.loads(n[k])]
    return {"추가": len(set(n) - set(o)), "삭제": len(set(o) - set(n)), "값변경": len(changed),
            "수기수정변경": len(cust), "예시": changed[:10]}


def to_csv(master):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=C.MASTER_FIELDS, extrasaction="ignore")
    w.writeheader()
    for r in master:
        w.writerow({k: r.get(k, "") for k in C.MASTER_FIELDS})
    return "﻿" + buf.getvalue()          # 엑셀에서 한글 깨짐 방지(BOM)


def from_csv(text):
    if isinstance(text, (bytes, bytearray)):
        text = text.decode("utf-8-sig")
    text = text.lstrip("﻿")
    out = []
    for i, r in enumerate(csv.DictReader(io.StringIO(text)), 1):
        rec = {k: (r.get(k) or "").strip() for k in C.MASTER_FIELDS}
        rec["품번"] = norm_code(rec["품번"])
        if not rec["품번"]:
            continue
        for k in ("BPA", "MOQ"):
            v = rec[k]
            if v != "":
                try:
                    f = float(v.replace(",", ""))
                    rec[k] = int(f) if f.is_integer() else f
                except ValueError:
                    pass
        try:
            rec["순번"] = int(float(rec["순번"])) if rec["순번"] else i
        except ValueError:
            rec["순번"] = i
        out.append(rec)
    return out


def save_csv(master, path):
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(to_csv(master))


def load_csv(path):
    with open(path, encoding="utf-8-sig") as f:
        return from_csv(f.read())


# ── 신규 품목 감지 + 제안값 ───────────────────────────────────────
def shortage_profile(row, dates, start=None, end=None):
    """재고에서 출발한 누적 잔량. (기간중 최저값, 첫 마이너스일) 반환."""
    bal, mn, first = _num(row[C.I_STOCK]), 0.0, None
    lo = None
    for j, d in enumerate(dates):
        bal -= _num(row[C.N_FIXED + j])
        if (start is None or d >= start) and (end is None or d <= end):
            lo = bal if lo is None else min(lo, bal)
            if bal < 0 and first is None:
                first = d
    return (lo if lo is not None else 0.0), first


def _txt0(v):
    """MRP의 빈 참조(0)나 오류값은 제안하지 않는다."""
    v = _clean(v)
    return "" if v in (0, "0") or (isinstance(v, str) and v.startswith("#")) else v


def find_new_items(mrp, master, start=None, end=None):
    """MRP에는 있고 마스터에는 없는 품목 → 제안값이 채워진 검토 목록."""
    known = {m["품번"] for m in master}
    out = []
    for row in mrp.rows:
        code = row[C.I_ITEM]
        if code in known:
            continue
        lo, first = shortage_profile(row, mrp.dates, start, end)
        bpa = mrp.bpa.get(code, "")
        out.append({
            "포함": True,
            "품번": code,
            "출처": C.NEW_SOURCE,
            "BPA": bpa,
            "공급자": _txt0(row[C.I_SITE]),
            "규격": _txt0(row[C.I_NAME]),
            "구분": "",
            "비고": "",
            "MOQ": "",
            # ↓ 검토 참고용 (마스터에는 저장 안 함)
            "담당자": _clean(row[C.I_OWNER]),
            "공정": _clean(row[C.I_PROCESS]),
            "재고": _num(row[C.I_STOCK]),
            "기간소요": sum(_num(v) for v in row[C.N_FIXED:]),
            "기간최저": lo,
            "첫부족일": first.strftime("%Y-%m-%d") if first else "",
            "BPA출처": "단가시트" if bpa != "" else "없음",
        })
    # 부족 예상 품목을 위로
    out.sort(key=lambda r: (r["기간최저"] >= 0, r["첫부족일"] or "9999", r["담당자"], r["품번"]))
    return out


def add_items(master, approved, today=None):
    """승인된 신규 품목을 마스터에 추가. 새 마스터(list) 반환."""
    today = (today or dt.date.today()).strftime("%Y-%m-%d")
    known = {m["품번"] for m in master}
    nxt = max([int(m.get("순번") or 0) for m in master] + [0]) + 1
    new = list(master)
    for a in approved:
        if not a.get("포함", True) or a["품번"] in known:
            continue
        rec = {k: a.get(k, "") for k in C.MASTER_FIELDS}
        rec["순번"], rec["등록일"] = nxt, today
        nxt += 1
        known.add(a["품번"])
        new.append(rec)
    return new


def exceptions(mrp, master):
    """검토 화면용 예외 목록."""
    master = CU.items_only(master)
    in_mrp = mrp.index
    not_in_mrp = [m["품번"] for m in master if m["품번"] not in in_mrp]
    bpa_diff = []
    for m in master:
        b = mrp.bpa.get(m["품번"])
        if b is not None and str(m.get("BPA", "")) not in ("", str(b)):
            bpa_diff.append({"품번": m["품번"], "마스터BPA": m["BPA"], "단가시트BPA": b})
    owner_na = [r[C.I_ITEM] for r in mrp.rows if str(r[C.I_OWNER]).startswith("#") or r[C.I_OWNER] in (None, "")]
    return {"not_in_mrp": not_in_mrp, "bpa_diff": bpa_diff, "owner_na": owner_na}
