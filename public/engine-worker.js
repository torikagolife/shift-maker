// 計算エンジン（ブラウザの中で Python を動かす Web Worker）
// Pyodide ＋ numpy / scipy（HiGHS）/ openpyxl を読み込み、/py/ の Python ファイルを呼び出す
const PYODIDE = 'https://cdn.jsdelivr.net/pyodide/v0.27.2/full/';
const FILES = ['shift_engine.py', 'solver_highs.py', 'engine_web.py'];
importScripts(PYODIDE + 'pyodide.js');

const status = msg => postMessage({ type: 'status', msg });
let api = null;

const ready = (async () => {
  status('Python を読み込み中…');
  const py = await loadPyodide({ indexURL: PYODIDE });
  status('計算用ライブラリを読み込み中…（初回は少し時間がかかります）');
  try {
    await py.loadPackage(['numpy', 'scipy', 'openpyxl']);
  } catch (e) {
    // openpyxl が Pyodide に無い場合は PyPI から入れる
    await py.loadPackage(['numpy', 'scipy', 'micropip']);
    await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl')");
  }
  status('エンジンを準備中…');
  const v = Date.now();
  for (const f of FILES) {
    const r = await fetch('/py/' + f + '?v=' + v, { credentials: 'same-origin' });
    if (!r.ok) throw new Error(f + ' が読めません（' + r.status + '）');
    py.FS.writeFile('/home/pyodide/' + f, await r.text());
  }
  py.runPython("import sys\nif '/home/pyodide' not in sys.path: sys.path.insert(0, '/home/pyodide')");
  api = py.pyimport('engine_web');
  status('ready');
})().catch(e => { postMessage({ type: 'status', msg: 'error', error: String(e.message || e) }); throw e; });

onmessage = async ({ data }) => {
  const { id, fn, args } = data;
  try {
    await ready;
    const out = api[fn](...args);
    const res = typeof out === 'string' ? JSON.parse(out) : out;
    if (res && res.error) postMessage({ id, ok: false, error: res.error });
    else postMessage({ id, ok: true, result: res });
  } catch (e) {
    postMessage({ id, ok: false, error: String(e.message || e) });
  }
};
