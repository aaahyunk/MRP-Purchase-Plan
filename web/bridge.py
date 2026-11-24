# 웹(워커) ↔ mrp_plan 연결부. 상태(현재 MRP)는 이 모듈 전역에 보관한다.
import json, sys, site, zipfile, datetime as dt

_sp = site.getsitepackages()[0]
for _w in ("openpyxl-3.1.5-py2.py3-none-any.whl", "et_xmlfile-2.0.0-py3-none-any.whl"):
    zipfile.ZipFile("/tmp/" + _w).extractall(_sp)
sys.path.insert(0, "/app")

from mrp_plan import loader, master as M, engine as E, writer as W, config as C  # noqa: E402

STATE = {"mrp": None, "file": "", "hol": E.load_holidays(open("/app/holidays.csv", encoding="utf-8").read())}


def _j(o):
    return json.dumps(o, ensure_ascii=False, default=lambda x: x.isoformat() if isinstance(x, (dt.date, dt.datetime)) else str(x))


def _d(s):
    return dt.date.fromisoformat(s[:10])


def _master(a):
    return M.from_csv(a["master_csv"]) if a.get("master_csv") else []


def api_load_master(args):
    a = json.loads(args)
    if a["kind"] == "plan":
        ms, rep = M.extract_from_plan(open("/tmp/" + a["name"], "rb").read())
    else:
        ms, rep = M.from_csv(open("/tmp/" + a["name"], "rb").read()), {"duplicates": [], "cleared": []}
    return _j({"csv": M.to_csv(ms), "count": len(ms), "report": rep})


def api_load_mrp(args):
    a = json.loads(args)
    mrp = loader.load_mrp(open("/tmp/" + a["name"], "rb").read(), a["name"])
    STATE["mrp"], STATE["file"] = mrp, a["name"]
    ok = bool(mrp.rows) and not any(c.level == "error" for c in mrp.checks)
    s, e = E.default_period(mrp) if ok else (None, None)
    eff = E.effective_holidays(STATE["hol"], mrp) if ok else []
    near = [h for h in eff if ok and mrp.dates[0] - dt.timedelta(days=7) <= h["날짜"] <= mrp.dates[-1] + dt.timedelta(days=14)]
    return _j({"ok": ok, "checks": [c.__dict__ for c in mrp.checks], "count": len(mrp.rows),
               "dates": mrp.dates, "start": s, "end": e, "holidays": near,
               "owners": sorted({str(r[C.I_OWNER]) for r in mrp.rows})})


def api_review(args):
    a = json.loads(args)
    mrp, ms = STATE["mrp"], _master(a)
    items = M.find_new_items(mrp, ms, _d(a["start"]), _d(a["end"]))
    exc = M.exceptions(mrp, ms)
    return _j({"items": items, "exc": {"not_in_mrp": len(exc["not_in_mrp"]), "bpa_diff": exc["bpa_diff"],
                                       "owner_na": exc["owner_na"]}})


def _merged(a):
    ms = _master(a)
    take = [x for x in a.get("new_items", []) if x.get("포함", True)]
    for x in take:                         # 웹 입력값 숫자 변환
        for k in ("MOQ", "BPA"):
            v = x.get(k, "")
            if isinstance(v, str) and v.strip() != "":
                try:
                    f = float(v.replace(",", ""))
                    x[k] = int(f) if f.is_integer() else f
                except ValueError:
                    pass
    return M.add_items(ms, take, dt.date.today())


def _run(a):
    mrp = STATE["mrp"]
    ms = _merged(a)
    s, e = _d(a["start"]), _d(a["end"])
    eff = E.effective_holidays(STATE["hol"], mrp)
    res = E.compute(mrp, ms, s, e, [h for h in eff if h["적용"]])
    return mrp, ms, s, e, eff, res


def api_compute(args):
    a = json.loads(args)
    mrp, ms, s, e, eff, res = _run(a)
    keep = ["상태", "품번", "규격", "공급자", "담당자", "공정", "출처", "MOQ", "발주수량", "발주일",
            "기간중최저값", "발주잔량", "210", "재고", "기간말과부족"]
    orders = [{k: r.get(k) for k in keep} for r in res if r["발주수량"] > 0]
    return _j({"summary": E.summary(res), "orders": orders})


def api_build(args):
    a = json.loads(args)
    mrp, ms, s, e, eff, res = _run(a)
    summ = E.summary(res)
    params = {"start": s, "end": e, "파일명": STATE["file"], "생성시각": a.get("now", "")}
    new_log = a.get("new_items", [])
    b = W.build_workbook(open("/app/template.xlsx", "rb").read(), mrp, ms, params, eff, new_log, summ,
                         M.exceptions(mrp, ms))
    open("/tmp/out.xlsx", "wb").write(b)
    return _j({"master_csv": M.to_csv(ms), "count": len(ms), "summary": summ,
               "filename": f"구매계획_통합_{mrp.dates[0]:%y%m%d}.xlsx"})
