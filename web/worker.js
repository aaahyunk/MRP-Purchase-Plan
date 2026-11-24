/* Pyodide 워커: 브라우저 안에서 Python(mrp_plan)을 실행한다. 파일은 PC 밖으로 나가지 않는다. */
const LOCAL_PYODIDE = "./vendor/pyodide/";
const CDN_PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";
const WHEELS = ["openpyxl-3.1.5-py2.py3-none-any.whl", "et_xmlfile-2.0.0-py3-none-any.whl"];
const MODULES = ["__init__.py", "config.py", "loader.py", "master.py", "engine.py", "writer.py", "review.py"];

let py = null;
const post = (m) => self.postMessage(m);
const progress = (id, text) => post({ id, type: "progress", text });

async function fetchBytes(url) {
  const r = await fetch(url, { cache: "no-cache" });
  if (!r.ok) throw new Error(`${url} 를 불러오지 못했습니다 (${r.status})`);
  return new Uint8Array(await r.arrayBuffer());
}

async function boot(id) {
  progress(id, "Python 엔진 불러오는 중…");
  let base = LOCAL_PYODIDE;
  try {
    const h = await fetch(LOCAL_PYODIDE + "pyodide.js", { method: "HEAD" });
    if (!h.ok) throw 0;
  } catch { base = CDN_PYODIDE; }
  importScripts(base + "pyodide.js");
  py = await loadPyodide({ indexURL: base });

  progress(id, "엑셀 라이브러리(openpyxl) 설치 중…");
  for (const w of WHEELS) py.FS.writeFile("/tmp/" + w, await fetchBytes("./vendor/wheels/" + w));
  py.FS.mkdirTree("/app/mrp_plan");
  for (const f of MODULES) py.FS.writeFile("/app/mrp_plan/" + f, await fetchBytes("./py/mrp_plan/" + f));
  py.FS.writeFile("/app/template.xlsx", await fetchBytes("./assets/template.xlsx"));
  py.FS.writeFile("/app/holidays.csv", await fetchBytes("./assets/holidays.csv"));
  await py.runPythonAsync(await (await fetch("./bridge.py", { cache: "no-cache" })).text());
  return { pyodide: base === LOCAL_PYODIDE ? "local" : "cdn" };
}

async function call(fn, args, buffers = {}) {
  for (const [k, v] of Object.entries(buffers)) py.FS.writeFile("/tmp/" + k, v);
  py.globals.set("_args_json", JSON.stringify(args || {}));
  const out = await py.runPythonAsync(`${fn}(_args_json)`);
  return JSON.parse(out);
}

self.onmessage = async (ev) => {
  const { id, cmd, args, file } = ev.data;
  try {
    if (cmd === "init") return post({ id, type: "done", result: await boot(id) });
    if (!py) throw new Error("엔진이 아직 준비되지 않았습니다.");
    if (cmd === "build") {
      progress(id, "엑셀 생성 중… (30초~1분)");
      const res = await call("api_build", args);
      const bytes = py.FS.readFile("/tmp/out.xlsx");
      return post({ id, type: "done", result: res, bytes }, [bytes.buffer]);
    }
    const buffers = file ? { [file.name]: new Uint8Array(file.bytes) } : {};
    if (cmd === "loadMaster" && args.kind === "plan") progress(id, "기존 통합파일에서 마스터 추출 중… (최초 1회, 15~40초)");
    if (cmd === "loadMrp") progress(id, "MRP 파일 읽는 중…");
    const map = { loadMaster: "api_load_master", loadMrp: "api_load_mrp", review: "api_review", compute: "api_compute" };
    post({ id, type: "done", result: await call(map[cmd], args, buffers) });
  } catch (e) {
    post({ id, type: "error", error: String(e && e.message ? e.message : e).split("\n").slice(-6).join("\n") });
  }
};
