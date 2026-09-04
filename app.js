/* Local Data Lab UI glue. Requires engine.js (window.DataLabEngine). */
(function () {
'use strict';
var E = window.DataLabEngine;
function $(id) { return document.getElementById(id); }

var tables = [], seq = 0, activeId = null, page = 0, queryResult = null, lastChartSVG = '';
var PAGE = 100, CAP = 100000;

/* ---------- helpers ---------- */
function stamp() {
  var d = new Date(), p = function (n) { return String(n).padStart(2, '0'); };
  return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '-' + p(d.getHours()) + p(d.getMinutes());
}
function download(blob, name) {
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 600);
}
function log(msg, cls) {
  var box = $('importLog');
  if (box.querySelector('.hint')) box.innerHTML = '';
  var d = document.createElement('div');
  d.className = 'msg' + (cls ? ' ' + cls : '');
  d.textContent = msg;
  box.prepend(d);
  while (box.children.length > 30) box.lastChild.remove();
}
function uniqueName(base) {
  base = (base || 'table').replace(/\.[a-z0-9]+$/i, '').replace(/[^\w\-]+/g, '_').slice(0, 40) || 'table';
  var names = tables.map(function (t) { return t.name.toLowerCase(); });
  if (names.indexOf(base.toLowerCase()) < 0) return base;
  var n = 2;
  while (names.indexOf((base + '_' + n).toLowerCase()) >= 0) n++;
  return base + '_' + n;
}
function getActive() {
  for (var i = 0; i < tables.length; i++) if (tables[i].id === activeId) return tables[i];
  return tables[0] || null;
}
function tablesMap() {
  var m = {};
  tables.forEach(function (t) { m[t.name] = { columns: t.columns, rows: t.rows }; });
  return m;
}
function flattenOpts() {
  return {
    separator: ($('sepInput').value || '.').slice(0, 4),
    arrayMode: $('arrayMode').value,
    joinSeparator: ', ',
    maxDepth: Math.min(20, Math.max(1, parseInt($('depthInput').value || '10', 10))),
    trimStrings: true,
    explode: $('explodeCheck').checked
  };
}
function addTable(name, columns, rows, meta) {
  name = uniqueName(name);
  if (rows.length > CAP) { log(name + ': capped at ' + CAP.toLocaleString() + ' rows (had ' + rows.length.toLocaleString() + ').'); rows = rows.slice(0, CAP); }
  var t = { id: ++seq, name: name, columns: columns, rows: rows, meta: meta || '' };
  tables.push(t);
  activeId = t.id;
  page = 0;
  renderAll();
  return t;
}

/* ---------- importers ---------- */
function importDelimited(name, text, delimOpt) {
  var r = E.parseDelimited(text, { delimiter: delimOpt || 'auto', header: $('firstHeaderCheck').checked });
  if (!r.columns.length) throw new Error(name + ': no data parsed.');
  addTable(name, r.columns, r.rows, 'delimited(' + JSON.stringify(r.delimiter) + ')');
  log('✓ ' + name + ' — ' + r.rows.length.toLocaleString() + ' rows × ' + r.columns.length + ' cols (delimited).', 'okk');
}
function importJSON(name, text) {
  var res = E.tryParseJSON(text);
  if (!res.ok) throw new Error(name + ': ' + res.error);
  var cands = E.findArrayCandidates(res.value);
  var root = $('rootInput').value.trim();
  if (!root && cands.length) {
    var s = cands.slice().sort(function (a, b) { return b.objCount - a.objCount || b.length - a.length; });
    if (s[0].objCount > 0) root = s[0].path;
  }
  var c = E.coerceToRows(res.value, root);
  var t = E.rowsToTable(c.rows, flattenOpts());
  if (!t.columns.length) throw new Error(name + ': nothing tabular found.');
  addTable(name, t.columns, t.rows, res.format + ' · ' + c.format + (t.truncated ? ' · capped' : ''));
  log('✓ ' + name + ' — ' + res.format + ' → ' + t.rows.length.toLocaleString() + ' rows × ' + t.columns.length + ' cols.', 'okk');
}
function importXML(name, text) {
  var t = E.xmlToTable(text, {
    recordTag: $('xmlInput').value.trim() || undefined,
    separator: ($('sepInput').value || '.'), arrayMode: $('arrayMode').value,
    maxDepth: parseInt($('depthInput').value || '10', 10), explode: $('explodeCheck').checked
  });
  addTable(name, t.columns, t.rows, 'xml · <' + t.recordTag + '>');
  log('✓ ' + name + ' — XML <' + t.recordTag + '> → ' + t.rows.length.toLocaleString() + ' rows × ' + t.columns.length + ' cols.', 'okk');
}
function importXlsxBuf(name, buf) {
  return E.importXlsx(buf, { header: $('firstHeaderCheck').checked, dateConvert: $('xlsxDatesCheck').checked }).then(function (sheets) {
    sheets.forEach(function (sh, i) {
      addTable(sheets.length > 1 ? name + '_' + sh.sheet : name, sh.columns, sh.rows, 'xlsx · ' + sh.sheet);
    });
    log('✓ ' + name + ' — XLSX ' + sheets.length + ' sheet(s) imported.', 'okk');
  });
}
function importZipBuf(name, buf) {
  return E.unzipAll(buf).then(function (files) {
    var keys = Object.keys(files).filter(function (k) { return !/\/$/.test(k) && k.indexOf('__MACOSX') < 0; });
    if (!keys.length) throw new Error(name + ': zip is empty.');
    var n = 0;
    var chain = Promise.resolve();
    keys.forEach(function (k) {
      chain = chain.then(function () {
        var low = k.toLowerCase();
        try {
          if (/\.xlsx?$/.test(low) || /\.xlsm$/.test(low)) return importXlsxBuf(k.split('/').pop(), files[k].buffer.slice(0));
          var txt = new TextDecoder().decode(files[k]);
          if (/\.jsonl?$/.test(low) || /\.ndjson$/.test(low)) importJSON(k.split('/').pop(), txt);
          else if (/\.xml$/.test(low)) importXML(k.split('/').pop(), txt);
          else importDelimited(k.split('/').pop(), txt, 'auto');
          n++;
        } catch (err) { log('✕ ' + k + ': ' + err.message, 'err'); }
      });
    });
    return chain.then(function () { log('✓ ' + name + ' — zip unpacked: ' + n + ' table(s).', 'okk'); });
  });
}
function importTextAuto(name, text) {
  var t = text.trim();
  var fmt = ($('pasteFormat') && document.querySelector('#spane-paste.active')) ? $('pasteFormat').value : 'auto';
  if (fmt === 'auto') {
    if (t[0] === '<') fmt = 'xml';
    else {
      var r = E.tryParseJSON(t);
      if (r.ok && (t[0] === '{' || t[0] === '[' || r.format === 'JSON Lines')) fmt = 'json';
      else fmt = 'csv';
    }
  }
  if (fmt === 'json') importJSON(name, text);
  else if (fmt === 'xml') importXML(name, text);
  else {
    var d = fmt === 'tsv' ? '\t' : fmt === 'csv' ? ($('pasteDelim') ? ($('pasteDelim').value === 'auto' ? 'auto' : $('pasteDelim').value) : 'auto') : 'auto';
    var hdr = $('pasteHeader') ? $('pasteHeader').checked : $('firstHeaderCheck').checked;
    var rr = E.parseDelimited(text, { delimiter: d, header: hdr });
    if (!rr.columns.length) throw new Error(name + ': no data parsed.');
    addTable(name, rr.columns, rr.rows, 'pasted ' + fmt);
    log('✓ ' + name + ' — ' + rr.rows.length.toLocaleString() + ' rows × ' + rr.columns.length + ' cols.', 'okk');
  }
}
function readFiles(list) {
  var files = Array.prototype.slice.call(list);
  if (!files.length) return;
  $('fileInfo').textContent = files.length + ' file(s): ' + files.map(function (f) { return f.name; }).join(', ');
  var chain = Promise.resolve();
  files.forEach(function (f) {
    chain = chain.then(function () {
      var low = f.name.toLowerCase();
      var base = f.name.replace(/\.[a-z0-9]+$/i, '');
      if (/\.xlsx?$/.test(low) || /\.xlsm$/.test(low)) return f.arrayBuffer().then(function (b) { return importXlsxBuf(base, b); }).catch(function (e) { log('✕ ' + f.name + ': ' + e.message, 'err'); });
      if (/\.zip$/.test(low)) return f.arrayBuffer().then(function (b) { return importZipBuf(base, b); }).catch(function (e) { log('✕ ' + f.name + ': ' + e.message, 'err'); });
      return f.text().then(function (t) {
        try {
          if (/\.jsonl?$/.test(low) || /\.ndjson$/.test(low)) importJSON(base, t);
          else if (/\.xml$/.test(low)) importXML(base, t);
          else importDelimited(base, t, 'auto');
        } catch (e) { log('✕ ' + f.name + ': ' + e.message, 'err'); }
      });
    });
  });
  chain.then(renderAll);
}

/* ---------- rendering ---------- */
function renderAll() {
  $('srcPill').textContent = tables.length + ' table' + (tables.length === 1 ? '' : 's');
  renderTableList(); renderSchemaPreview(); renderJoin(); renderSqlSchema(); renderChartSources();
}
function renderTableList() {
  var box = $('tableList');
  box.innerHTML = '';
  if (!tables.length) { box.innerHTML = '<p class="hint">Load data to begin.</p>'; $('tblPill').textContent = 'no tables'; return; }
  tables.forEach(function (t) {
    var b = document.createElement('button');
    b.type = 'button';
    b.className = t.id === (getActive() || {}).id ? 'active' : '';
    b.innerHTML = '';
    var s = document.createElement('span'); s.textContent = t.name;
    var n = document.createElement('span'); n.className = 'n'; n.textContent = '  ' + t.rows.length.toLocaleString() + '×' + t.columns.length;
    b.appendChild(s); b.appendChild(n);
    b.addEventListener('click', function () { activeId = t.id; page = 0; renderAll(); });
    box.appendChild(b);
  });
  var a = getActive();
  $('tblPill').textContent = a ? a.name + ' · ' + a.rows.length.toLocaleString() + ' rows × ' + a.columns.length + ' cols' : 'no tables';
}
function renderSchemaPreview() {
  var t = getActive();
  var sb = $('schemaBox'), thead = $('prevTable').querySelector('thead'), tbody = $('prevTable').querySelector('tbody');
  thead.innerHTML = ''; tbody.innerHTML = '';
  if (!t) { sb.innerHTML = '<p class="hint">—</p>'; $('previewInfo').textContent = 'No preview.'; return; }
  var schema = E.inferSchema(t.rows, t.columns);
  var tbl = document.createElement('table');
  var h = document.createElement('thead'), hr = document.createElement('tr');
  ['column', 'type', 'distinct', 'nulls', 'rename →', 'cast', 'hide'].forEach(function (x) {
    var th = document.createElement('th'); th.textContent = x; hr.appendChild(th);
  });
  h.appendChild(hr); tbl.appendChild(h);
  var tb = document.createElement('tbody');
  t.columns.forEach(function (c) {
    var tr = document.createElement('tr');
    function td(txt) { var d = document.createElement('td'); d.textContent = txt; tr.appendChild(d); return d; }
    td(c);
    td(schema[c].type + (schema[c].min != null ? ' [' + schema[c].min + '…' + schema[c].max + ']' : ''));
    td(String(schema[c].distinct)); td(String(schema[c].nulls));
    var rtd = document.createElement('td');
    var ri = document.createElement('input'); ri.value = c; ri.title = 'New name, Enter to apply';
    ri.addEventListener('change', function () {
      if (!ri.value.trim() || ri.value === c) return;
      var nt = E.normalizeTable(t, [{ op: 'rename', from: c, to: ri.value.trim() }]);
      t.columns = nt.columns; t.rows = nt.rows; renderAll();
    });
    rtd.appendChild(ri); tr.appendChild(rtd);
    var ctd = document.createElement('td');
    var sel = document.createElement('select');
    ['string', 'integer', 'number', 'boolean', 'date'].forEach(function (ty) {
      var o = document.createElement('option'); o.value = ty; o.textContent = ty;
      if (schema[c].type === ty) o.selected = true;
      sel.appendChild(o);
    });
    sel.addEventListener('change', function () {
      var nt = E.normalizeTable(t, [{ op: 'cast', col: c, type: sel.value }]);
      t.rows = nt.rows; renderAll();
    });
    ctd.appendChild(sel); tr.appendChild(ctd);
    var htd = document.createElement('td');
    var hb = document.createElement('button'); hb.className = 'btn ghost sm'; hb.textContent = 'hide';
    hb.addEventListener('click', function () {
      var nt = E.normalizeTable(t, [{ op: 'hide', col: c }]);
      t.columns = nt.columns; t.rows = nt.rows; renderAll();
    });
    htd.appendChild(hb); tr.appendChild(htd);
    tb.appendChild(tr);
  });
  tbl.appendChild(tb);
  sb.innerHTML = ''; sb.appendChild(tbl);
  // preview
  var trh = document.createElement('tr');
  t.columns.forEach(function (c) { var th = document.createElement('th'); th.textContent = c; trh.appendChild(th); });
  thead.appendChild(trh);
  var pages = Math.max(1, Math.ceil(t.rows.length / PAGE));
  page = Math.min(page, pages - 1);
  t.rows.slice(page * PAGE, page * PAGE + PAGE).forEach(function (r) {
    var tr = document.createElement('tr');
    t.columns.forEach(function (c) {
      var td = document.createElement('td'), v = E.cellText(r[c]);
      td.textContent = v.length > 120 ? v.slice(0, 120) + '…' : v;
      if (v.length > 120) td.title = v;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  $('previewInfo').textContent = t.rows.length.toLocaleString() + ' rows × ' + t.columns.length + ' cols' + (t.meta ? ' · ' + t.meta : '');
  $('pageInfo').textContent = 'page ' + (page + 1) + ' / ' + pages;
}
function fillSelect(sel, items, keep) {
  var cur = keep ? sel.value : null;
  sel.innerHTML = '';
  items.forEach(function (x) {
    var o = document.createElement('option'); o.value = x; o.textContent = x;
    sel.appendChild(o);
  });
  if (cur && items.indexOf(cur) >= 0) sel.value = cur;
}
function renderJoin() {
  var names = tables.map(function (t) { return t.name; });
  fillSelect($('joinLeft'), names, true);
  fillSelect($('joinRight'), names, true);
  if (names.length > 1 && $('joinLeft').value === $('joinRight').value) $('joinRight').selectedIndex = 1;
  function tbl(nm) { return tables.filter(function (t) { return t.name === nm; })[0]; }
  var L = tbl($('joinLeft').value), R = tbl($('joinRight').value);
  fillSelect($('joinLeftKey'), L ? L.columns : [], true);
  fillSelect($('joinRightKey'), R ? R.columns : [], true);
}
function renderSqlSchema() {
  var box = $('sqlSchema');
  box.innerHTML = '';
  if (!tables.length) { box.innerHTML = '<p class="hint">No tables.</p>'; return; }
  tables.forEach(function (t) {
    var d = document.createElement('details');
    var s = document.createElement('summary'); s.textContent = t.name + ' (' + t.rows.length.toLocaleString() + ')';
    var u = document.createElement('ul');
    t.columns.forEach(function (c) { var li = document.createElement('li'); li.textContent = c; u.appendChild(li); });
    d.appendChild(s); d.appendChild(u); box.appendChild(d);
  });
}
function renderChartSources() {
  var names = tables.map(function (t) { return t.name; });
  if (queryResult) names.push('$query');
  var cur = $('chartSource').value;
  fillSelect($('chartSource'), names.length ? names : ['—'], false);
  if (cur && names.indexOf(cur) >= 0) $('chartSource').value = cur;
  refreshChartCols();
}
function chartRows() {
  if ($('chartSource').value === '$query' && queryResult) return queryResult;
  var t = tables.filter(function (x) { return x.name === $('chartSource').value; })[0] || getActive();
  return t ? { columns: t.columns, rows: t.rows } : { columns: [], rows: [] };
}
function refreshChartCols() {
  var r = chartRows();
  fillSelect($('chartX'), r.columns, true);
  fillSelect($('chartY'), r.columns, true);
}
function renderSQLResult(res) {
  var thead = $('sqlTable').querySelector('thead'), tbody = $('sqlTable').querySelector('tbody');
  thead.innerHTML = ''; tbody.innerHTML = '';
  var trh = document.createElement('tr');
  res.columns.forEach(function (c) { var th = document.createElement('th'); th.textContent = c; trh.appendChild(th); });
  thead.appendChild(trh);
  res.rows.slice(0, 200).forEach(function (r) {
    var tr = document.createElement('tr');
    res.columns.forEach(function (c) {
      var td = document.createElement('td'), v = E.cellText(r[c]);
      td.textContent = v.length > 120 ? v.slice(0, 120) + '…' : v;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  $('sqlMsg').textContent = res.rows.length.toLocaleString() + ' row(s) × ' + res.columns.length + ' col(s)' + (res.rows.length > 200 ? ' — showing first 200' : '');
}

/* ---------- events ---------- */
document.addEventListener('DOMContentLoaded', function () {
  var root = document.documentElement;
  try { if (localStorage.getItem('ldl-theme')) root.setAttribute('data-theme', localStorage.getItem('ldl-theme')); } catch (e) {}
  $('themeBtn').addEventListener('click', function () {
    var cur = root.getAttribute('data-theme') === 'light' ? '' : 'light';
    if (cur) root.setAttribute('data-theme', cur); else root.removeAttribute('data-theme');
    try { localStorage.setItem('ldl-theme', cur); } catch (e) {}
  });
  $('resetAllBtn').addEventListener('click', function () {
    if (!confirm('Delete all tables, results and charts?')) return;
    tables = []; activeId = null; page = 0; queryResult = null; lastChartSVG = '';
    $('importLog').innerHTML = '<p class="hint">Import messages appear here.</p>';
    $('chartBox').innerHTML = '<p class="hint">Configure and render — chart appears here.</p>';
    var st = $('sqlTable'); st.querySelector('thead').innerHTML = ''; st.querySelector('tbody').innerHTML = '';
    renderAll();
  });
  document.querySelectorAll('.tab').forEach(function (t) {
    t.addEventListener('click', function () {
      document.querySelectorAll('.tab').forEach(function (x) { x.classList.remove('active'); });
      document.querySelectorAll('.tabpane').forEach(function (x) { x.classList.remove('active'); });
      t.classList.add('active');
      $('spane-' + t.dataset.stab).classList.add('active');
    });
  });
  // upload
  var dz = $('dropzone'), fi = $('fileInput');
  dz.addEventListener('click', function () { fi.click(); });
  dz.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') fi.click(); });
  ['dragover', 'dragenter'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('over'); }); });
  dz.addEventListener('drop', function (e) { if (e.dataTransfer && e.dataTransfer.files.length) readFiles(e.dataTransfer.files); });
  fi.addEventListener('change', function () { readFiles(fi.files); fi.value = ''; });
  // paste
  $('addPasteBtn').addEventListener('click', function () {
    var txt = $('pasteText').value;
    if (!txt.trim()) { log('Paste is empty.', 'err'); return; }
    try { importTextAuto($('pasteName').value.trim() || 'pasted_data', txt); }
    catch (e) { log('✕ paste: ' + e.message, 'err'); }
  });
  // samples
  document.querySelectorAll('.sample').forEach(function (b) {
    b.addEventListener('click', function () {
      var k = b.dataset.sample;
      try {
        if (k === 'sales') {
          var cols = ['region', 'rep', 'month', 'revenue', 'units'];
          var rows = [
            { region: 'EMEA', rep: 'Ali', month: '2026-01', revenue: 12000, units: 130 },
            { region: 'EMEA', rep: 'Bea', month: '2026-01', revenue: 9500, units: 98 },
            { region: 'NA', rep: 'Cid', month: '2026-01', revenue: 17100, units: 160 },
            { region: 'NA', rep: 'Ali', month: '2026-02', revenue: 14300, units: 141 },
            { region: 'APAC', rep: 'Dee', month: '2026-02', revenue: 8200, units: 88 }
          ];
          var bytes = E.buildXlsx(cols, rows, 'Sales');
          E.importXlsx(bytes.buffer.slice ? bytes.buffer.slice(0) : bytes, { header: true, dateConvert: true }).then(function (sheets) {
            sheets.forEach(function (sh) { addTable('sales', sh.columns, sh.rows, 'generated XLSX round-trip'); });
            log('✓ sales — generated XLSX built + re-imported (round-trip proof).', 'okk');
          }).catch(function (e) { log('✕ sales xlsx: ' + e.message, 'err'); });
        }
        else if (k === 'employees') importDelimited('employees', E.SAMPLES.employees, ',');
        else if (k === 'events') importDelimited('events', E.SAMPLES.events, '\t');
        else if (k === 'orders') { $('rootInput').value = 'orders'; importJSON('orders', E.SAMPLES.orders); }
        else if (k === 'catalog') importXML('catalog', E.SAMPLES.catalog);
      } catch (e) { log('✕ sample ' + k + ': ' + e.message, 'err'); }
    });
  });
  // table ops
  $('renameBtn').addEventListener('click', function () {
    var t = getActive(); if (!t) return;
    var n = $('renameInput').value.trim(); if (!n) return;
    t.name = uniqueName(n); $('renameInput').value = ''; renderAll();
  });
  $('dupBtn').addEventListener('click', function () {
    var t = getActive(); if (!t) return;
    addTable(t.name + '_copy', t.columns.slice(), t.rows.map(function (r) { return Object.assign({}, r); }), t.meta);
  });
  $('delBtn').addEventListener('click', function () {
    var t = getActive(); if (!t) return;
    tables = tables.filter(function (x) { return x.id !== t.id; });
    activeId = tables.length ? tables[0].id : null; page = 0; renderAll();
  });
  $('trimBtn').addEventListener('click', function () {
    var t = getActive(); if (!t) return;
    var nt = E.normalizeTable(t, [{ op: 'trim' }]); t.rows = nt.rows; renderAll();
  });
  $('dropRowsBtn').addEventListener('click', function () {
    var t = getActive(); if (!t) return;
    var before = t.rows.length;
    var nt = E.normalizeTable(t, [{ op: 'dropEmptyRows' }]); t.rows = nt.rows; renderAll();
    log(t.name + ': dropped ' + (before - t.rows.length) + ' empty row(s).');
  });
  $('dropColsBtn').addEventListener('click', function () {
    var t = getActive(); if (!t) return;
    var nt = E.normalizeTable(t, [{ op: 'dropEmptyCols' }]); t.columns = nt.columns; t.rows = nt.rows; renderAll();
  });
  $('dedupeBtn').addEventListener('click', function () {
    var t = getActive(); if (!t) return;
    var before = t.rows.length;
    var nt = E.normalizeTable(t, [{ op: 'dedupe' }]); t.rows = nt.rows; renderAll();
    log(t.name + ': removed ' + (before - t.rows.length) + ' duplicate(s).');
  });
  $('prevPage').addEventListener('click', function () { page = Math.max(0, page - 1); renderSchemaPreview(); });
  $('nextPage').addEventListener('click', function () { page++; renderSchemaPreview(); });
  function expActive(kind) {
    var t = getActive(); if (!t) return;
    if (kind === 'csv') download(new Blob([E.buildDelimited(t.columns, t.rows, ',')], { type: 'text/csv' }), t.name + '-' + stamp() + '.csv');
    if (kind === 'tsv') download(new Blob([E.buildDelimited(t.columns, t.rows, '\t')], { type: 'text/tab-separated-values' }), t.name + '-' + stamp() + '.tsv');
    if (kind === 'json') download(new Blob([JSON.stringify(t.rows, null, 2)], { type: 'application/json' }), t.name + '-' + stamp() + '.json');
    if (kind === 'xlsx') download(new Blob([E.buildXlsx(t.columns, t.rows, t.name)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), t.name + '-' + stamp() + '.xlsx');
  }
  $('expCsvBtn').addEventListener('click', function () { expActive('csv'); });
  $('expTsvBtn').addEventListener('click', function () { expActive('tsv'); });
  $('expJsonBtn').addEventListener('click', function () { expActive('json'); });
  $('expXlsxBtn').addEventListener('click', function () { expActive('xlsx'); });
  // join
  $('joinLeft').addEventListener('change', renderJoin);
  $('joinRight').addEventListener('change', renderJoin);
  $('joinRunBtn').addEventListener('click', function () {
    var L = tables.filter(function (t) { return t.name === $('joinLeft').value; })[0];
    var R = tables.filter(function (t) { return t.name === $('joinRight').value; })[0];
    if (!L || !R) { $('joinInfo').textContent = 'Need two tables.'; return; }
    try {
      var r = E.joinTables(L, R, { leftKey: $('joinLeftKey').value, rightKey: $('joinRightKey').value, type: $('joinType').value });
      addTable($('joinOutName').value.trim() || 'joined', r.columns, r.rows, L.name + ' ' + $('joinType').value + ' ⨝ ' + R.name);
      $('joinInfo').textContent = '→ ' + r.rows.length.toLocaleString() + ' rows × ' + r.columns.length + ' cols';
    } catch (e) { $('joinInfo').textContent = 'Join failed: ' + e.message; }
  });
  // sql
  function firstName(fb) {
    return tables.length ? tables[0].name : fb;
  }
  document.querySelectorAll('#sqlSamples button').forEach(function (b) {
    b.addEventListener('click', function () {
      var a = firstName('employees');
      var q = {
        select: 'SELECT * FROM ' + a + ' LIMIT 20',
        agg: 'SELECT dept, COUNT(*) AS n, AVG(salary) AS avg_salary FROM ' + a + ' GROUP BY dept ORDER BY n DESC',
        join: tables.length > 1 ? 'SELECT a.name, a.dept, b.budget FROM ' + tables[0].name + ' a JOIN ' + tables[1].name + ' b ON a.dept = b.dept LIMIT 20' : 'SELECT a.name, b.orderId FROM ' + a + ' a JOIN orders b ON 1=1 LIMIT 5  -- pick two real tables first',
        like: "SELECT name, salary FROM " + a + " WHERE name LIKE '%a%' ORDER BY salary DESC LIMIT 10"
      }[b.dataset.q];
      $('sqlInput').value = q;
    });
  });
  $('sqlRunBtn').addEventListener('click', function () {
    var q = $('sqlInput').value;
    if (!q.trim()) return;
    try {
      var r = E.runSQL(q, tablesMap());
      queryResult = { columns: r.result.columns, rows: r.result.rows };
      renderSQLResult(queryResult);
      renderChartSources();
    } catch (e) { $('sqlMsg').textContent = 'Error: ' + e.message; }
  });
  $('sqlSaveBtn').addEventListener('click', function () {
    if (!queryResult) return;
    addTable($('sqlSaveName').value.trim() || 'query_result', queryResult.columns.slice(), queryResult.rows.map(function (r) { return Object.assign({}, r); }), 'sql result');
  });
  $('sqlExpCsvBtn').addEventListener('click', function () {
    if (!queryResult) return;
    download(new Blob([E.buildDelimited(queryResult.columns, queryResult.rows, ',')], { type: 'text/csv' }), 'query-' + stamp() + '.csv');
  });
  $('sqlExpXlsxBtn').addEventListener('click', function () {
    if (!queryResult) return;
    download(new Blob([E.buildXlsx(queryResult.columns, queryResult.rows, 'query')], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'query-' + stamp() + '.xlsx');
  });
  // charts
  $('chartSource').addEventListener('change', refreshChartCols);
  $('chartRenderBtn').addEventListener('click', function () {
    var src = chartRows();
    if (!src.rows.length) { $('chartInfo').textContent = 'No data.'; return; }
    var type = $('chartType').value, x = $('chartX').value, y = $('chartY').value;
    var lim = Math.min(60, Math.max(2, parseInt($('chartLimit').value || '15', 10)));
    var title = ($('chartSource').value || 'data') + ' · ' + type;
    try {
      if (type === 'bar' || type === 'pie') {
        var items = E.prepareBar(src.rows, x, y, $('chartAgg').value, type === 'pie' ? 12 : lim);
        lastChartSVG = E.buildChartSVG({ type: type, items: items, title: title + ' — ' + $('chartAgg').value + '(' + y + ') by ' + x });
        $('chartInfo').textContent = items.length + ' group(s) from ' + src.rows.length.toLocaleString() + ' rows.';
      } else {
        var pts = E.prepareXY(src.rows, x, y, 300);
        if (!pts.length) { $('chartInfo').textContent = 'Need numeric X and Y with values.'; return; }
        lastChartSVG = E.buildChartSVG({ type: type, points: pts, title: title + ' — ' + y + ' vs ' + x });
        $('chartInfo').textContent = pts.length + ' point(s).';
      }
      $('chartBox').innerHTML = lastChartSVG;
    } catch (e) { $('chartInfo').textContent = 'Chart failed: ' + e.message; }
  });
  $('chartSvgBtn').addEventListener('click', function () {
    if (!lastChartSVG) return;
    download(new Blob([lastChartSVG], { type: 'image/svg+xml' }), 'chart-' + stamp() + '.svg');
  });
  $('chartPngBtn').addEventListener('click', function () {
    if (!lastChartSVG) return;
    var img = new Image();
    var svg64 = btoa(unescape(encodeURIComponent(lastChartSVG)));
    img.onload = function () {
      var c = document.createElement('canvas');
      c.width = 1360; c.height = 760;
      var ctx = c.getContext('2d');
      var bg = getComputedStyle(document.body).backgroundColor;
      ctx.fillStyle = root.getAttribute('data-theme') === 'light' ? '#ffffff' : '#171e28';
      ctx.fillRect(0, 0, c.width, c.height);
      ctx.drawImage(img, 0, 0, c.width, c.height);
      c.toBlob(function (b) { download(b, 'chart-' + stamp() + '.png'); });
    };
    img.src = 'data:image/svg+xml;base64,' + svg64;
  });
  // export all
  $('zipBtn').addEventListener('click', function () {
    if (!tables.length) return;
    var files = tables.map(function (t) {
      return { name: t.name + '.csv', data: new TextEncoder().encode(E.buildDelimited(t.columns, t.rows, ',')) };
    });
    download(new Blob([E.buildZipStored(files)], { type: 'application/zip' }), 'data-lab-' + stamp() + '.zip');
  });

  // seed: useful workspace out of the box
  try {
    importDelimited('employees', E.SAMPLES.employees, ',');
    var deps = E.parseDelimited('dept,budget,head\nEng,500000,Ada\nResearch,300000,Alan\nSales,200000,Adaeze', { delimiter: ',', header: true });
    addTable('departments', deps.columns, deps.rows, 'seed');
    $('rootInput').value = 'orders';
    importJSON('orders', E.SAMPLES.orders);
    $('rootInput').value = '';
    $('sqlInput').value = 'SELECT dept, COUNT(*) AS n, AVG(salary) AS avg_salary FROM employees GROUP BY dept ORDER BY n DESC';
  } catch (e) { log('seed: ' + e.message, 'err'); }
  renderAll();
});
})();
