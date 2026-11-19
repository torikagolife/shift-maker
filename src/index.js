// シフト作成 Web（Cloudflare Workers 無料プラン + D1）
// Worker はパスワード確認・データ保存・ファイル配信だけ。自動作成の計算はブラウザの中（public/engine-worker.js）で行う

// ---------------------------------------------------------------- D1
const SCHEMA = [
  `CREATE TABLE IF NOT EXISTS files (name TEXT PRIMARY KEY, data TEXT, meta TEXT, updated TEXT)`,
  `CREATE TABLE IF NOT EXISTS plan (month TEXT, staff_key TEXT, day INTEGER, value TEXT, PRIMARY KEY (month, staff_key, day))`,
  `CREATE TABLE IF NOT EXISTS special (id INTEGER PRIMARY KEY AUTOINCREMENT, month TEXT, date TEXT, target TEXT, min INTEGER, max INTEGER, note TEXT)`,
  `CREATE TABLE IF NOT EXISTS month_meta (month TEXT PRIMARY KEY, holidays_n INTEGER)`,
];
let schemaReady = false;
async function ensureSchema(env) {
  if (schemaReady) return;
  await env.DB.batch(SCHEMA.map(s => env.DB.prepare(s)));
  schemaReady = true;
}

const now = () => new Date().toLocaleString('sv-SE', { timeZone: 'Asia/Tokyo' });
const getFile = (env, name) => env.DB.prepare('SELECT data, meta, updated FROM files WHERE name = ?').bind(name).first();
const putFile = (env, name, data, meta) => env.DB.prepare('INSERT OR REPLACE INTO files (name, data, meta, updated) VALUES (?, ?, ?, ?)')
  .bind(name, data, meta ? JSON.stringify(meta) : null, now()).run();

// ---------------------------------------------------------------- 小物
const json = (obj, status = 200) => new Response(JSON.stringify(obj), {
  status, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' } });
const fail = (msg, status = 400) => json({ error: msg }, status);
const MONTH_RE = /^\d{4}-\d{2}$/;
const B64_RE = /^[A-Za-z0-9+/=]+$/;
const MAX_B64 = 1_500_000; // D1 の1行の上限（約2MB）より小さく

function xlsxResponse(b64, filename) {
  const s = atob(b64); const u = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i);
  return new Response(u, { headers: {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}` } });
}

// パスワード（Basic認証）。ユーザー名は何でもよい
function authorized(req, env) {
  const h = req.headers.get('Authorization') || '';
  if (!h.startsWith('Basic ')) return false;
  try {
    const bin = atob(h.slice(6)); const u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    const pass = new TextDecoder().decode(u).split(':').slice(1).join(':');
    return pass === env.APP_PASSWORD;
  } catch { return false; }
}

// ---------------------------------------------------------------- API
async function api(req, env, url) {
  const p = url.pathname, m = req.method;
  const month = url.searchParams.get('month');

  // 画面に必要なものをまとめて返す
  if (p === '/api/state' && m === 'GET') {
    if (!MONTH_RE.test(month || '')) return fail('month が必要です');
    const [set, plan, sp, meta, result] = await Promise.all([
      env.DB.prepare('SELECT meta, updated FROM files WHERE name = ?').bind('settings').first(),
      env.DB.prepare('SELECT staff_key AS key, day, value FROM plan WHERE month = ?').bind(month).all(),
      env.DB.prepare('SELECT * FROM special WHERE month = ? ORDER BY date, id').bind(month).all(),
      env.DB.prepare('SELECT holidays_n FROM month_meta WHERE month = ?').bind(month).first(),
      env.DB.prepare('SELECT meta, updated FROM files WHERE name = ?').bind('result:' + month).first(),
    ]);
    return json({
      info: set ? JSON.parse(set.meta) : null, settingsUpdated: set?.updated || null,
      plan: plan.results, specials: sp.results, holidays_n: meta?.holidays_n ?? null,
      result: result ? { ...JSON.parse(result.meta), created: result.updated } : null,
    });
  }

  // 設定ファイル（ブラウザで読み取った内容 info と、ファイル本体の base64 を受け取る）
  if (p === '/api/settings' && m === 'POST') {
    const { b64, info } = await req.json();
    if (!b64 || !B64_RE.test(b64) || b64.length > MAX_B64 || !info?.staff) return fail('設定ファイルが正しくありません');
    await putFile(env, 'settings', b64, info);
    return json({ ok: true });
  }
  if (p === '/api/settings' && m === 'GET') {
    const f = await getFile(env, 'settings');
    if (!f) return fail('設定ファイルがまだありません', 404);
    if (url.searchParams.get('raw')) return json({ b64: f.data });   // 計算エンジン用
    return xlsxResponse(f.data, '設定.xlsx');
  }

  // 予定（希望休など）の書き込み。value が空なら消す
  if (p === '/api/plan' && m === 'PUT') {
    const { month: mo, changes } = await req.json();
    if (!MONTH_RE.test(mo || '') || !Array.isArray(changes)) return fail('month と changes が必要です');
    const st = changes.slice(0, 3000).map(c => {
      const v = String(c.value ?? '').trim().slice(0, 30);
      return v
        ? env.DB.prepare('INSERT OR REPLACE INTO plan (month, staff_key, day, value) VALUES (?, ?, ?, ?)').bind(mo, String(c.key), Number(c.day), v)
        : env.DB.prepare('DELETE FROM plan WHERE month = ? AND staff_key = ? AND day = ?').bind(mo, String(c.key), Number(c.day));
    });
    for (let i = 0; i < st.length; i += 100) await env.DB.batch(st.slice(i, i + 100));
    return json({ ok: true, saved: st.length });
  }
  if (p === '/api/plan/clear' && m === 'POST') {
    const { month: mo } = await req.json();
    if (!MONTH_RE.test(mo || '')) return fail('month が必要です');
    await env.DB.prepare('DELETE FROM plan WHERE month = ?').bind(mo).run();
    return json({ ok: true });
  }
  if (p === '/api/meta' && m === 'PUT') {
    const { month: mo, holidays_n } = await req.json();
    if (!MONTH_RE.test(mo || '')) return fail('month が必要です');
    await env.DB.prepare('INSERT OR REPLACE INTO month_meta (month, holidays_n) VALUES (?, ?)')
      .bind(mo, holidays_n === '' || holidays_n == null ? null : Number(holidays_n)).run();
    return json({ ok: true });
  }

  // 特別日
  if (p === '/api/special' && m === 'POST') {
    const x = await req.json();
    if (!MONTH_RE.test(x.month || '') || !x.date || !x.target) return fail('日付と対象が必要です');
    const num = v => (v === '' || v == null ? null : Number(v));
    await env.DB.prepare('INSERT INTO special (month, date, target, min, max, note) VALUES (?, ?, ?, ?, ?, ?)')
      .bind(x.month, x.date, x.target, num(x.min), num(x.max), x.note || '').run();
    return json({ ok: true });
  }
  if (p === '/api/special' && m === 'DELETE') {
    await env.DB.prepare('DELETE FROM special WHERE id = ?').bind(Number(url.searchParams.get('id'))).run();
    return json({ ok: true });
  }

  // 自動作成の結果（計算はブラウザ。結果だけ保存する）
  if (p === '/api/result' && m === 'PUT') {
    const { month: mo, summary, xlsx } = await req.json();
    if (!MONTH_RE.test(mo || '') || !summary || !xlsx || !B64_RE.test(xlsx) || xlsx.length > MAX_B64) return fail('結果が正しくありません');
    await putFile(env, 'result:' + mo, xlsx, summary);
    return json({ ok: true, created: now() });
  }
  if (p === '/api/result' && m === 'DELETE') {
    if (!MONTH_RE.test(month || '')) return fail('month が必要です');
    await env.DB.prepare('DELETE FROM files WHERE name = ?').bind('result:' + month).run();
    return json({ ok: true });
  }
  if (p === '/api/result' && m === 'GET') {
    const f = await getFile(env, 'result:' + month);
    if (!f) return fail('まだ自動作成していません', 404);
    return xlsxResponse(f.data, `勤務表_${month}_自動作成.xlsx`);
  }

  return fail('not found', 404);
}

// パスワードなしで配ってよいもの（計算エンジンのプログラム。データは含まない）
const PUBLIC = /^\/(engine-worker\.js|py\/[\w.]+\.py)$/;

export default {
  async fetch(req, env) {
    const url = new URL(req.url);
    if (PUBLIC.test(url.pathname)) return env.ASSETS.fetch(req);
    if (!env.APP_PASSWORD) return new Response('APP_PASSWORD（シークレット）を設定してください', { status: 500 });
    if (!authorized(req, env)) {
      return new Response('パスワードが必要です', { status: 401, headers: { 'WWW-Authenticate': 'Basic realm="shift", charset="UTF-8"' } });
    }
    try {
      if (url.pathname.startsWith('/api/')) {
        await ensureSchema(env);
        return await api(req, env, url);
      }
      return env.ASSETS.fetch(req);
    } catch (e) {
      return fail(String(e.message || e), 500);
    }
  },
};
