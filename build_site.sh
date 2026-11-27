#!/usr/bin/env bash
# GitHub Pages 용 정적 사이트 조립:  web/ + src/mrp_plan + Pyodide(자체 호스팅) + openpyxl 휠
# 사용: bash tools/build_site.sh [_site]
set -euo pipefail
cd "$(dirname "$0")/.."
OUT="${1:-_site}"
PYODIDE_VER="0.27.7"
rm -rf "$OUT"
mkdir -p "$OUT/py/mrp_plan" "$OUT/vendor/pyodide" "$OUT/vendor/wheels"
cp -r web/. "$OUT/"
cp src/mrp_plan/*.py "$OUT/py/mrp_plan/"

TMP="$(mktemp -d)"
( cd "$TMP" && npm pack "pyodide@${PYODIDE_VER}" --silent >/dev/null && tar xzf "pyodide-${PYODIDE_VER}.tgz" )
for f in pyodide.js pyodide.asm.js pyodide.asm.wasm python_stdlib.zip pyodide-lock.json; do
  cp "$TMP/package/$f" "$OUT/vendor/pyodide/"
done
python3 -m pip download --no-deps -q openpyxl==3.1.5 et_xmlfile==2.0.0 -d "$OUT/vendor/wheels"
touch "$OUT/.nojekyll"
echo "site built → $OUT ($(du -sh "$OUT" | cut -f1))"
