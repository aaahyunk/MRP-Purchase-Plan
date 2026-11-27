"""사용자 수기 수정사항 보존.

'구매계획_통합' 결과 파일에서 사람이 직접 바꾼 것을 읽어 마스터에 저장하고(capture),
다음 결과 파일을 만들 때 같은 품목·같은 위치에 다시 적용한다(apply).

  - 값      : 출처·BPA·공급자·규격·구분·비고·MOQ → 마스터 기본 필드로 저장 (기존과 동일)
  - 메모    : 품목 두 행(MRP계획/과부족)의 모든 셀
  - 서식    : 채우기색·글자색·굵게/기울임/밑줄/취소선·표시형식 중 '출처별 기본 서식'과 다른 것
  - 추가입력 : 도구가 비워두는 칸에 사람이 넣은 값/수식 (예: MRP계획 행 V열의 =SUM(...))
               → 행 번호가 바뀌어도 상대참조를 맞춰 다시 넣음
  - 레이아웃 : 열 너비·숨김 열·틀 고정 (마스터의 '#LAYOUT' 행)

저장 형식: 품목별 JSON 문자열 {"s": {"0V": {...}}, "m": {"0G": {...}}, "x": {"0V": {...}}}
  키 = 행구분(0=MRP계획, 1=과부족) + 열문자
"""
import json
import re
import datetime as dt

from openpyxl.comments import Comment
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill, Font, Color
from openpyxl.utils import get_column_letter as CL

from . import config as C

LAYOUT_KEY = "#LAYOUT"
MASS_TAG = "양산품목X"
MASS_SEP = " / "

# 도구가 값을 쓰지 않는 칸 (사람이 채우면 '추가입력'으로 보존)
EMPTY_SLOTS = {
    0: {21, 22},                                              # MRP계획 행: U, V
    1: {4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 23},   # 과부족 행
}


# ── 서식 키 ─────────────────────────────────────────────────────
def _ck(color):
    if color is None:
        return None
    t = color.type
    if t == "rgb":
        return ["r", color.rgb]
    if t == "theme":
        return ["t", color.theme, round(color.tint or 0, 3)]
    if t == "indexed":
        return ["i", color.indexed]
    return None


def style_key(cell):
    f, ft = cell.fill, cell.font
    return {
        "fill": [f.fill_type, _ck(f.fgColor) if f.fill_type else None],
        "font": [_ck(ft.color), bool(ft.b), bool(ft.i), ft.u or None, bool(ft.strike)],
        "nf": cell.number_format,
    }


def _color(k):
    if not k:
        return None
    if k[0] == "r":
        return Color(rgb=k[1])
    if k[0] == "t":
        return Color(theme=k[1], tint=k[2])
    if k[0] == "i":
        return Color(indexed=k[1])
    return None


def proto_keys(template_wb):
    """템플릿의 출처별 견본 행 서식 키. (by_source, known) 반환."""
    ws = template_wb[C.PLAN_SHEET]
    by_src, known = {}, {}
    for i, k in enumerate(C.PROTO_KEYS):
        r = C.FIRST_DATA_ROW + 2 * i
        by_src[k] = [[style_key(ws.cell(r + d, c)) for c in range(1, C.N_PLAN_COLS + 1)] for d in (0, 1)]
        for d in (0, 1):
            for c in range(1, C.N_PLAN_COLS + 1):
                sk = by_src[k][d][c - 1]
                for comp, v in sk.items():
                    known.setdefault((d, c, comp), []).append(v)
    return by_src, known


# ── 캡처 ────────────────────────────────────────────────────────
def _jsonable(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return {"d": v.isoformat()}
    return v


def capture_item(ws, r, src, protos):
    """품목 하나(r=MRP계획 행)의 수기 수정사항 dict. 없으면 {}."""
    by_src, known = protos
    base = by_src.get(src) or by_src[C.NEW_SOURCE]
    out = {"s": {}, "m": {}, "x": {}}
    rows = [0, 1] if ws.cell(r + 1, 2).value == "과부족" else [0]
    for d in rows:
        for c in range(1, C.N_PLAN_COLS + 1):
            cell = ws.cell(r + d, c)
            key = f"{d}{CL(c)}"
            sk = style_key(cell)
            diff = {comp: v for comp, v in sk.items()
                    if v != base[d][c - 1][comp] and v not in known.get((d, c, comp), [])}
            if diff:
                out["s"][key] = diff
            if cell.comment is not None and (cell.comment.text or "").strip():
                out["m"][key] = {"t": cell.comment.text, "a": cell.comment.author or ""}
            if c in EMPTY_SLOTS[d] and cell.value not in (None, ""):
                out["x"][key] = {"v": _jsonable(cell.value), "o": cell.coordinate}
    out = {k: v for k, v in out.items() if v}
    return out


def capture_layout(ws):
    widths, hidden = {}, []
    for k, dim in ws.column_dimensions.items():
        if dim.width:
            widths[k] = round(dim.width, 2)
        if dim.hidden:
            hidden.append(k)
    return {"w": widths, "h": hidden, "f": ws.freeze_panes}


def strip_mass_tag(text):
    """비고에서 도구가 붙인 '양산품목X'를 떼어 사람이 쓴 내용만 남긴다."""
    if not isinstance(text, str):
        return text
    s = re.sub(r"\s*/?\s*" + MASS_TAG + r"\s*$", "", text)
    return s.strip()


def summarize(master):
    n = {"메모": 0, "서식": 0, "추가입력": 0, "품목": 0}
    for m in master:
        c = loads(m)
        if c:
            n["품목"] += 1
        n["메모"] += len(c.get("m", {}))
        n["서식"] += len(c.get("s", {}))
        n["추가입력"] += len(c.get("x", {}))
    return n


# ── 마스터 레코드 ↔ JSON ────────────────────────────────────────
def dumps(d):
    return json.dumps(d, ensure_ascii=False, separators=(",", ":")) if d else ""


def loads(rec):
    v = rec.get("수기수정") if isinstance(rec, dict) else rec
    if not v:
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except ValueError:
        return {}


def items_only(master):
    return [m for m in master if m.get("품번") != LAYOUT_KEY]


def get_layout(master):
    for m in master:
        if m.get("품번") == LAYOUT_KEY:
            return loads(m)
    return {}


def layout_record(layout):
    rec = {k: "" for k in C.MASTER_FIELDS}
    rec.update({"품번": LAYOUT_KEY, "규격": "열너비·틀고정 설정 (지우지 마세요)", "순번": 0, "수기수정": dumps(layout)})
    return rec


# ── 적용 (writer) ───────────────────────────────────────────────
def apply_item(ws, r, custom):
    if not custom:
        return
    for key, ov in custom.get("s", {}).items():
        cell = ws[f"{key[1:]}{r + int(key[0])}"]
        if "fill" in ov:
            t, ck = ov["fill"]
            cell.fill = PatternFill(fill_type=t, fgColor=_color(ck)) if t else PatternFill(fill_type=None)
        if "font" in ov:
            ck, b, i, u, s = ov["font"]
            f = cell.font
            cell.font = Font(name=f.name, sz=f.sz, b=b, i=i, u=u, strike=s, color=_color(ck), vertAlign=f.vertAlign)
        if "nf" in ov:
            cell.number_format = ov["nf"]
    for key, m in custom.get("m", {}).items():
        cell = ws[f"{key[1:]}{r + int(key[0])}"]
        cm = Comment(m["t"], m.get("a") or "")
        cm.width, cm.height = 220, 110
        cell.comment = cm
    for key, x in custom.get("x", {}).items():
        target = f"{key[1:]}{r + int(key[0])}"
        v = x["v"]
        if isinstance(v, dict) and "d" in v:
            v = dt.datetime.fromisoformat(v["d"])
        elif isinstance(v, str) and v.startswith("=") and x.get("o"):
            try:
                v = Translator(v, origin=x["o"]).translate_formula(target)
            except Exception:                     # 번역 실패 시 원문 유지
                pass
        ws[target] = v


def apply_layout(ws, layout):
    if not layout:
        return
    for k, w in layout.get("w", {}).items():
        ws.column_dimensions[k].width = w
    hidden = set(layout.get("h", []))
    for k, dim in list(ws.column_dimensions.items()):
        dim.hidden = k in hidden
    for k in hidden:
        ws.column_dimensions[k].hidden = True
    ws.freeze_panes = layout.get("f") or None
