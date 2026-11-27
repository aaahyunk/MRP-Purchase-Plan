/* UI 로직 (HITL). 계산은 worker.js(Pyodide)에서. 마스터는 이 브라우저 localStorage에 저장. */
const $ = (s) => document.querySelector(s);
const LS_MASTER = "mrpPlan.master.v1", LS_PREV = "mrpPlan.master.prev.v1", LS_META = "mrpPlan.meta.v1";
const GUBUN = ["양산", "SVC", "개발", "단종", "자재단종", "미사용"];
const SOURCES = ["와이앤제이", "월발주", "Wafer", "R.CHIP", "Shoe,Tiiun", "가스켓", "Sheet", "SME", "신규"];
const PAGE = 50;

const S = { master: null, meta: {}, mrp: null, items: [], view: [], page: 0, sel: new Set(), computed: false };

/* ── 저장소 (localStorage는 실패할 수 있으므로 모두 try/catch) ── */
const store = {
  get(k) { try { return localStorage.getItem(k); } catch { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); return true; } catch { return false; } },
};
function saveMaster(csv, note, count) {
  const prev = store.get(LS_MASTER);
  if (prev && prev !== csv) store.set(LS_PREV, prev);
  S.master = csv;
  S.meta = { count: count ?? countCsv(csv), updated: new Date().toLocaleString("ko-KR"), note };
  const ok = store.set(LS_MASTER, csv) && store.set(LS_META, JSON.stringify(S.meta));
  if (!ok) toast("⚠ 브라우저 저장소에 저장하지 못했습니다. 마스터 CSV를 꼭 내려받아 보관하세요.", 6000);
  renderMaster();
}
const countCsv = (csv) => Math.max(0, csv.trim().split(/\r?\n/).filter((l) => !l.startsWith("#LAYOUT")).length - 1);

/* ── 워커 RPC ── */
const worker = new Worker("./worker.js");
let seq = 0; const pending = {};
worker.onmessage = (e) => {
  const m = e.data, p = pending[m.id]; if (!p) return;
  if (m.type === "progress") return p.init ? ($("#engineText").textContent = m.text) : overlay(m.text);
  delete pending[m.id];
  m.type === "error" ? p.reject(new Error(m.error)) : p.resolve(m);
};
function rpc(cmd, args = {}, file = null, msg = "처리 중…") {
  overlay(msg);
  return new Promise((resolve, reject) => {
    const id = ++seq; pending[id] = { resolve, reject };
    worker.postMessage({ id, cmd, args, file }, file ? [file.bytes] : []);
  }).finally(() => overlay(null));
}
async function readFile(f) { return { name: f.name, bytes: await f.arrayBuffer() }; }

/* ── 공통 UI ── */
function overlay(t) { $("#overlay").classList.toggle("hidden", !t); if (t) $("#overlayText").textContent = t; }
function toast(t, ms = 3500) { const x = $("#toast"); x.textContent = t; x.classList.remove("hidden"); clearTimeout(x._t); x._t = setTimeout(() => x.classList.add("hidden"), ms); }
function unlock(id, on = true) { $(id).classList.toggle("locked", !on); }
const esc = (v) => String(v ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const fmt = (v) => (typeof v === "number" ? v.toLocaleString("ko-KR", { maximumFractionDigits: 2 }) : esc(v));
function download(name, data, type) {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = Object.assign(document.createElement("a"), { href: url, download: name });
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 2000);
}
function showError(e) { console.error(e); toast("⛔ " + e.message, 8000); }
function invalidate() { S.computed = false; $("#kpis").innerHTML = ""; $("#orderTable").innerHTML = ""; unlock("#sec-download", false); }

/* ── 0. 마스터 ── */
function renderMaster() {
  const has = !!S.master;
  $("#masterEmpty").classList.toggle("hidden", has);
  $("#masterExport").disabled = !has;
  $("#masterUndo").disabled = !store.get(LS_PREV);
  $("#masterBadge").innerHTML = has
    ? `<span class="pill ok">${S.meta.count?.toLocaleString() ?? countCsv(S.master)}개</span><span class="small muted">최종 저장 ${esc(S.meta.updated || "-")}</span>`
    : `<span class="pill warn">마스터 없음</span>`;
}
$("#masterFile").onchange = async (e) => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  try {
    const kind = f.name.toLowerCase().endsWith(".csv") ? "csv" : "plan";
    const { result } = await rpc("loadMaster", { kind, name: f.name, old_csv: S.master || "" }, await readFile(f));
    saveMaster(result.csv, `${f.name} 에서 불러옴`, result.count);
    const r = result.report || {}, d = r.duplicates || [], c = r.cleared || [];
    $("#masterReport").innerHTML = `✅ ${result.count.toLocaleString()}개 불러옴` +
      (d.length ? ` · ⚠ 중복 품번 ${d.length}건은 첫 행 유지 (${d.slice(0, 6).map((x) => esc(x.품번)).join(", ")}${d.length > 6 ? "…" : ""})` : "") +
      (c.length ? ` · #N/A 값 ${c.length}건은 빈칸 처리` : "") +
      (r.custom ? `<br>✍ 수기 수정사항 보존: 품목 ${r.custom.품목.toLocaleString()}개 · 하이라이트/서식 ${r.custom.서식}칸 · 메모 ${r.custom.메모}개 · 추가 입력(수식 등) ${r.custom.추가입력}칸` : "") +
      (r.diff ? `<br>↔ 이전 마스터 대비: 추가 ${r.diff.추가} · 삭제 ${r.diff.삭제} · 값 변경 ${r.diff.값변경} · 메모/서식 변경 ${r.diff.수기수정변경} (잘못 올렸다면 '직전 상태로 되돌리기')` : "");
    if (S.mrp) await loadReview();
  } catch (err) { showError(err); }
};
$("#masterExport").onclick = () => download(`item_master_${new Date().toISOString().slice(0, 10)}.csv`, S.master, "text/csv;charset=utf-8");
$("#masterUndo").onclick = async () => {
  const prev = store.get(LS_PREV); if (!prev || !confirm("직전 승인 이전의 마스터로 되돌릴까요?")) return;
  S.master = prev; S.meta = { count: countCsv(prev), updated: new Date().toLocaleString("ko-KR"), note: "되돌림" };
  store.set(LS_MASTER, prev); store.set(LS_META, JSON.stringify(S.meta)); localStorage.removeItem(LS_PREV);
  renderMaster(); if (S.mrp) await loadReview(); toast("마스터를 되돌렸습니다.");
};

/* ── 1. MRP ── */
const drop = $("#drop");
["dragover", "dragenter"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach((t) => drop.addEventListener(t, () => drop.classList.remove("over")));
drop.addEventListener("drop", (e) => { e.preventDefault(); if (e.dataTransfer.files[0]) loadMrp(e.dataTransfer.files[0]); });
$("#mrpFile").onchange = (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) loadMrp(f); };

async function loadMrp(f) {
  try {
    const { result } = await rpc("loadMrp", { name: f.name }, await readFile(f), "MRP 파일 읽는 중…");
    const icon = { ok: "✅", info: "ℹ️", warn: "⚠️", error: "⛔" };
    $("#checks").innerHTML = result.checks.map((c) => `<li class="${c.level}"><span>${icon[c.level]}</span><b>${esc(c.title)}</b> <span class="muted">${esc(c.detail)}</span></li>`).join("");
    $("#dropText").textContent = `${f.name} · 품목 ${result.count.toLocaleString()}개`;
    if (!result.ok) { S.mrp = null; ["#sec-period", "#sec-review", "#sec-result", "#sec-download"].forEach((x) => unlock(x, false)); return; }
    S.mrp = result;
    const d0 = result.dates[0], dN = result.dates[result.dates.length - 1];
    for (const id of ["#start", "#end"]) { $(id).min = d0; $(id).max = dN; }
    $("#start").value = result.start; $("#end").value = result.end;
    $("#fOwner").innerHTML = `<option value="">담당자 전체</option>` + result.owners.map((o) => `<option>${esc(o)}</option>`).join("");
    renderHolidays(); renderPeriod();
    unlock("#sec-period"); unlock("#sec-review"); unlock("#sec-result");
    await loadReview();
  } catch (err) { showError(err); }
}

/* ── 2. 기간 ── */
function renderPeriod() {
  const s = $("#start").value, e = $("#end").value;
  const n = S.mrp ? S.mrp.dates.filter((d) => d >= s && d <= e).length : 0;
  $("#periodInfo").innerHTML = s > e ? `<span class="bad">⚠ 시작일 &gt; 종료일</span>` : `${n}일 반영 · MRP ${S.mrp.dates[0]} ~ ${S.mrp.dates.at(-1)}`;
}
function renderHolidays() {
  const h = S.mrp.holidays;
  $("#holTable").innerHTML = h.length
    ? `<tr><th>날짜</th><th>공휴일</th><th>그날 소요합계</th><th>판단</th></tr>` + h.map((x) => `<tr><td>${x.날짜}</td><td>${esc(x.명칭)}</td><td class="r">${x.소요합계 == null ? "-" : fmt(x.소요합계)}</td><td class="${x.적용 ? "" : "ok"}">${esc(x.사유)}</td></tr>`).join("")
    : `<tr><td class="muted">MRP 기간 근처에 공휴일이 없습니다.</td></tr>`;
}
for (const id of ["#start", "#end"]) $(id).onchange = async () => { renderPeriod(); invalidate(); await loadReview(true); };
$("#ignoreTail").onchange = () => invalidate();

/* ── 3. 신규 품목 (HITL) ── */
async function loadReview(keepEdits = false) {
  if (!S.mrp) return;
  if (!S.master) toast("마스터가 없어서 MRP 전 품목이 신규로 표시됩니다. 먼저 0단계에서 마스터를 불러오는 것을 권장합니다.", 6000);
  const { result } = await rpc("review", { master_csv: S.master || "", start: $("#start").value, end: $("#end").value }, null, "신규 품목 찾는 중…");
  const old = keepEdits ? Object.fromEntries(S.items.map((x) => [x.품번, x])) : {};
  S.items = result.items.map((x) => (old[x.품번] ? { ...x, ...pick(old[x.품번], EDITABLE), 포함: old[x.품번].포함, 기간최저: x.기간최저, 첫부족일: x.첫부족일 } : x));
  S.sel.clear(); S.page = 0; invalidate();
  renderExc(result.exc); applyFilter();
}
const EDITABLE = ["출처", "BPA", "공급자", "규격", "구분", "비고", "MOQ"];
const pick = (o, ks) => Object.fromEntries(ks.map((k) => [k, o[k]]));

function applyFilter() {
  const q = $("#q").value.trim().toLowerCase(), ow = $("#fOwner").value, sh = $("#fShort").checked, nm = $("#fNoMoq").checked;
  S.view = S.items.map((x, i) => i).filter((i) => {
    const x = S.items[i];
    if (ow && String(x.담당자) !== ow) return false;
    if (sh && !(x.기간최저 < 0)) return false;
    if (nm && String(x.MOQ ?? "") !== "") return false;
    if (q && !`${x.품번} ${x.규격} ${x.공급자}`.toLowerCase().includes(q)) return false;
    return true;
  });
  S.page = Math.min(S.page, Math.max(0, Math.ceil(S.view.length / PAGE) - 1));
  renderTable();
}
["#q", "#fOwner", "#fShort", "#fNoMoq"].forEach((id) => $(id).addEventListener("input", () => { S.page = 0; applyFilter(); }));

function renderTable() {
  const inc = S.items.filter((x) => x.포함).length;
  $("#reviewBadge").innerHTML = `<span class="pill">${S.items.length.toLocaleString()}개</span><span class="pill ok">포함 ${inc.toLocaleString()}</span>` +
    `<span class="pill warn">부족예상 ${S.items.filter((x) => x.기간최저 < 0).length.toLocaleString()}</span><span class="pill">MOQ 빈칸 ${S.items.filter((x) => String(x.MOQ ?? "") === "").length.toLocaleString()}</span>`;
  $("#selInfo").textContent = `필터 ${S.view.length.toLocaleString()}개 · 선택 ${S.sel.size.toLocaleString()}개`;
  const rows = S.view.slice(S.page * PAGE, (S.page + 1) * PAGE);
  const allSel = S.view.length > 0 && S.view.every((i) => S.sel.has(i));
  const head = `<thead><tr><th><input type="checkbox" id="selAll" ${allSel ? "checked" : ""} title="필터된 전체 선택"></th><th>포함</th><th>품번</th><th>규격</th><th>출처</th><th>BPA</th><th>공급자</th><th>구분</th><th>비고</th><th>MOQ</th><th>담당자</th><th>공정</th><th class="r">재고</th><th class="r">기간최저</th><th>첫부족일</th></tr></thead>`;
  const body = rows.map((i) => {
    const x = S.items[i];
    const inp = (k, w, type = "text") => `<input data-i="${i}" data-k="${k}" type="${type}" value="${esc(x[k])}" style="width:${w}px" ${k === "출처" ? 'list="sourceList"' : ""}>`;
    const gub = `<select data-i="${i}" data-k="구분"><option value=""></option>${GUBUN.map((g) => `<option ${x.구분 === g ? "selected" : ""}>${g}</option>`).join("")}</select>`;
    return `<tr class="${x.포함 ? "" : "off"}"><td><input type="checkbox" class="sel" data-i="${i}" ${S.sel.has(i) ? "checked" : ""}></td>
      <td><input type="checkbox" data-i="${i}" data-k="포함" ${x.포함 ? "checked" : ""}></td>
      <td class="mono">${esc(x.품번)}</td><td>${inp("규격", 230)}</td><td>${inp("출처", 80)}</td><td>${inp("BPA", 96)}${x.BPA출처 === "없음" ? '<span class="dotwarn" title="단가 시트에 BPA 없음"></span>' : ""}</td>
      <td>${inp("공급자", 110)}</td><td>${gub}</td><td>${inp("비고", 90)}</td><td>${inp("MOQ", 70, "number")}</td>
      <td>${esc(x.담당자)}</td><td>${esc(x.공정)}</td><td class="r">${fmt(x.재고)}</td><td class="r ${x.기간최저 < 0 ? "bad" : ""}">${fmt(x.기간최저)}</td><td>${esc(x.첫부족일)}</td></tr>`;
  }).join("");
  $("#newTable").innerHTML = head + `<tbody>${body || `<tr><td colspan="15" class="muted c">${S.items.length ? "필터 결과가 없습니다." : "✅ 신규 품목이 없습니다."}</td></tr>`}</tbody>`;
  const pages = Math.ceil(S.view.length / PAGE);
  $("#pager").innerHTML = pages > 1 ? `<button class="btn sm ghost" data-p="${S.page - 1}" ${S.page ? "" : "disabled"}>‹ 이전</button><span>${S.page + 1} / ${pages}</span><button class="btn sm ghost" data-p="${S.page + 1}" ${S.page < pages - 1 ? "" : "disabled"}>다음 ›</button>` : "";
}
$("#newTable").addEventListener("change", (e) => {
  const t = e.target, i = +t.dataset.i;
  if (t.id === "selAll") { S.view.forEach((j) => (t.checked ? S.sel.add(j) : S.sel.delete(j))); return renderTable(); }
  if (t.classList.contains("sel")) { t.checked ? S.sel.add(i) : S.sel.delete(i); return renderTable(); }
  const k = t.dataset.k; if (!k) return;
  S.items[i][k] = k === "포함" ? t.checked : t.value;
  invalidate(); if (k === "포함") renderTable();
});
$("#pager").addEventListener("click", (e) => { const p = e.target.dataset.p; if (p != null) { S.page = +p; renderTable(); } });
function bulk(fn) {
  if (!S.sel.size) return toast("먼저 왼쪽 체크박스로 품목을 선택하세요 (머리글 체크 = 필터된 전체).");
  S.sel.forEach((i) => fn(S.items[i])); invalidate(); renderTable(); toast(`${S.sel.size.toLocaleString()}개 적용`);
}
$("#bApply").onclick = () => {
  const g = $("#bGubun").value, s = $("#bSource").value.trim(), m = $("#bMoq").value;
  if (!g && !s && m === "") return toast("바꿀 값(구분/출처/MOQ)을 입력하세요.");
  bulk((x) => { if (g) x.구분 = g; if (s) x.출처 = s; if (m !== "") x.MOQ = m; });
};
$("#bInclude").onclick = () => bulk((x) => (x.포함 = true));
$("#bExclude").onclick = () => bulk((x) => (x.포함 = false));

function renderExc(exc) {
  $("#excBody").innerHTML = `
    <p>📭 마스터에는 있으나 오늘 MRP에 없는 품목: <b>${exc.not_in_mrp.toLocaleString()}개</b> → 결과에 0으로 표시됩니다.</p>
    <p>👤 담당자 #N/A (담당자 시트에 지점 없음): <b>${exc.owner_na.length}개</b> <span class="muted small">${exc.owner_na.slice(0, 30).map(esc).join(", ")}</span></p>
    <p>🔢 마스터 BPA ≠ 단가 시트 BPA: <b>${exc.bpa_diff.length}개</b></p>
    ${exc.bpa_diff.length ? `<table class="mini"><tr><th>품번</th><th>마스터</th><th>단가 시트</th></tr>${exc.bpa_diff.slice(0, 40).map((d) => `<tr><td>${esc(d.품번)}</td><td>${esc(d.마스터BPA)}</td><td>${esc(d.단가시트BPA)}</td></tr>`).join("")}</table>` : ""}`;
}

/* ── 4. 계산 ── */
function payload() {
  return { master_csv: S.master || "", start: $("#start").value, end: $("#end").value, ignore_tail: $("#ignoreTail").checked,
    new_items: S.items.map((x) => ({ ...x })), now: new Date().toLocaleString("ko-KR") };
}
$("#computeBtn").onclick = async () => {
  if ($("#start").value > $("#end").value) return toast("⚠ 시작일이 종료일보다 늦습니다.");
  try {
    const { result } = await rpc("compute", payload(), null, "계산 중…");
    const s = result.summary;
    $("#kpis").innerHTML = [["총 품목", s.총품목, ""], ["🔴 기간 중 부족 예상", s.부족예상, "warn"], ["🟡 발주 필요", s.발주필요, "accent"],
      ["잔량·210으로 커버", s.잔량커버, ""], ["MRP 미존재(0 표시)", s.MRP미존재, "muted"],
      ["🟥 대체품목 (규격 N/A)", s["대체품목(N/A)"], "red"], ["양산품목X 표시", s.양산품목X, "muted"]]
      .map(([t, v, c]) => `<div class="kpi ${c}"><div class="kv">${v.toLocaleString()}</div><div class="kt">${t}</div></div>`).join("");
    const o = result.orders.sort((a, b) => String(a.발주일).localeCompare(String(b.발주일)) || String(a.담당자).localeCompare(String(b.담당자)));
    $("#orderTable").innerHTML = `<thead><tr><th>발주일</th><th>품번</th><th>규격</th><th>공급자</th><th>담당자</th><th class="r">MOQ</th><th class="r">발주수량</th><th class="r">기간중최저</th><th class="r">발주잔량</th><th class="r">210</th><th class="r">기간말과부족</th><th>표시</th></tr></thead><tbody>` +
      o.map((r) => `<tr class="${r.대체품목 ? "na" : ""}"><td>${esc(r.발주일)}</td><td class="mono">${esc(r.품번)}</td><td class="ell">${esc(r.규격)}</td><td>${esc(r.공급자)}</td><td>${esc(r.담당자)}</td><td class="r">${fmt(r.MOQ)}</td><td class="r b">${fmt(r.발주수량)}</td><td class="r bad">${fmt(r.기간중최저값)}</td><td class="r">${fmt(r.발주잔량)}</td><td class="r">${fmt(r["210"])}</td><td class="r">${fmt(r.기간말과부족)}</td><td>${r.대체품목 ? '<span class="badge red">대체 N/A</span>' : ""}${r.양산품목X ? '<span class="badge">양산X</span>' : ""}</td></tr>`).join("") + "</tbody>";
    S.computed = true; unlock("#sec-download");
    $("#sec-result").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (err) { showError(err); }
};

/* ── 5. 승인 · 다운로드 ── */
$("#buildBtn").onclick = async () => {
  if (!S.computed) return toast("먼저 4단계에서 계산하기를 눌러 결과를 확인하세요.");
  const inc = S.items.filter((x) => x.포함).length;
  if (inc && !confirm(`신규 품목 ${inc.toLocaleString()}개를 마스터에 추가하고 엑셀을 만들까요?`)) return;
  try {
    const m = await rpc("build", payload(), null, "엑셀 생성 중…");
    download(m.result.filename, m.bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    saveMaster(m.result.master_csv, `${m.result.filename} 생성 시 승인`, m.result.count);
    if ($("#alsoCsv").checked) download(`item_master_${new Date().toISOString().slice(0, 10)}.csv`, m.result.master_csv, "text/csv;charset=utf-8");
    $("#buildInfo").innerHTML = `✅ <b>${esc(m.result.filename)}</b> 저장 · 마스터 ${m.result.count.toLocaleString()}개로 갱신`;
    await loadReview();   // 승인된 품목은 이제 신규가 아님
  } catch (err) { showError(err); }
};

/* ── 시작 ── */
(function init() {
  $("#bGubun").innerHTML += GUBUN.map((g) => `<option>${g}</option>`).join("");
  $("#sourceList").innerHTML = SOURCES.map((s) => `<option value="${s}">`).join("");
  S.master = store.get(LS_MASTER);
  try { S.meta = JSON.parse(store.get(LS_META) || "{}"); } catch { S.meta = {}; }
  renderMaster();
  const id = ++seq;
  pending[id] = {
    init: true,
    resolve: (m) => { $("#engine").innerHTML = `<span class="dot ok"></span>엔진 준비 완료${m.result.pyodide === "cdn" ? " (CDN)" : ""}`; unlock("#sec-mrp"); },
    reject: (e) => { $("#engine").innerHTML = `<span class="dot bad"></span>엔진 로드 실패`; showError(e); },
  };
  worker.postMessage({ id, cmd: "init" });
})();
