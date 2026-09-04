/* DuckDB-WASM runner — fully vendored, zero CDN.
 * Loads ./vendor/duckdb/* (populated by `npm run vendor` from the
 * package.json lockfile, so Renovate can bump versions natively).
 * The MVP build needs no COOP/COEP headers, so it works on GitHub Pages.
 * Needs http(s) — file:// has no worker/module support, caller falls back
 * to the built-in engine there. Classic script; ESM loaded via dynamic import.
 */
(function () {
'use strict';

var VENDOR = './vendor/duckdb/';
var state = { status: 'idle', error: '', duckdb: null, arrow: null, db: null, conn: null, loading: null };

function isFileProto() {
  try { return window.location.protocol === 'file:'; } catch (e) { return true; }
}
function qIdent(s) { return '"' + String(s).replace(/"/g, '""') + '"'; }
function abs(p) {
  // workers resolve relative URLs against the WORKER script location, not the
  // document — so hand DuckDB absolute URLs (subpath-safe via document.baseURI).
  try { return new URL(p, document.baseURI).href; } catch (e) { return p; }
}

async function ensure() {
  if (state.status === 'ready') return state;
  if (state.loading) return state.loading;
  state.loading = (async function () {
    state.status = 'loading'; state.error = '';
    try {
      if (isFileProto()) throw new Error('DuckDB needs http(s) — serve this folder (python3 -m http.server) or use the hosted site. Built-in engine still works from file://.');
      if (typeof Worker === 'undefined') throw new Error('Web Workers unavailable in this browser.');
      var duckdb = await import(VENDOR + 'duckdb-browser.mjs');
      var bundle = await duckdb.selectBundle({
        mvp: { mainModule: abs(VENDOR + 'duckdb-mvp.wasm'), mainWorker: abs(VENDOR + 'duckdb-browser-mvp.worker.js') }
      });
      var worker = new Worker(bundle.mainWorker);
      var db = new duckdb.AsyncDuckDB(new duckdb.ConsoleLogger(), worker);
      await db.instantiate(bundle.mainModule);
      var conn = await db.connect();
      var arrow = await import('apache-arrow'); // resolved by the importmap in index.html
      state.duckdb = duckdb; state.arrow = arrow; state.db = db; state.conn = conn;
      state.status = 'ready';
    } catch (e) {
      state.status = 'error';
      state.error = (e && e.message ? e.message : String(e)) +
        (/Failed to fetch|failed to load|Load failed/i.test((e && e.message) || '') ? ' — was vendor/ built? Run `npm run vendor` locally (deploy builds it automatically).' : '');
      throw new Error(state.error);
    } finally {
      state.loading = null;
    }
    return state;
  })();
  return state.loading;
}

async function registerTable(name, csvText) {
  var fname = 'ldl__' + name + '.csv';
  await state.db.registerFileText(fname, csvText);
  await state.conn.query(
    'CREATE OR REPLACE TABLE ' + qIdent(name) +
    ' AS SELECT * FROM read_csv(' + "'" + fname + "'" + ', header=true)'
  );
}
function toPlain(v) {
  if (typeof v === 'bigint') {
    return (v <= BigInt(Number.MAX_SAFE_INTEGER) && v >= BigInt(-Number.MAX_SAFE_INTEGER)) ? Number(v) : v.toString();
  }
  if (v instanceof Date) return isNaN(v) ? null : v.toISOString().slice(0, 10);
  if (v instanceof Uint8Array) {
    try { return new TextDecoder().decode(v); } catch (e) { return '[binary ' + v.length + 'B]'; }
  }
  if (v !== null && typeof v === 'object' && !Array.isArray(v) && v.toString && Object.prototype.toString.call(v) !== '[object Object]') {
    return String(v);
  }
  return v;
}
function arrowToResult(table) {
  var columns = table.schema.fields.map(function (f) { return f.name; });
  var rows = table.toArray().map(function (r) {
    var o = r.toJSON(), out = {};
    for (var k in o) out[k] = toPlain(o[k]);
    return out;
  });
  return { columns: columns, rows: rows };
}

async function runQuery(sql, tables) {
  await ensure();
  var names = Object.keys(tables);
  if (!names.length) throw new Error('No tables to query.');
  for (var i = 0; i < names.length; i++) {
    var t = tables[names[i]];
    if (!t.columns.length) continue;
    await registerTable(names[i], t.csv);
  }
  var res = await state.conn.query(sql);
  return arrowToResult(res);
}

window.DuckDBRunner = {
  status: function () { return { status: state.status, error: state.error }; },
  ensure: ensure,
  runQuery: runQuery
};
})();
