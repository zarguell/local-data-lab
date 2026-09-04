/* Local Data Lab engine — pure logic, zero deps, no network.
 * CSV/TSV + JSON(+lines) + XML + XLSX/ZIP import, flatten/normalize,
 * hash joins, mini-SQL, SVG charts, CSV/XLSX/ZIP export.
 * Browser-safe (no DOM at load). Node-exported for tests. */
(function () {
'use strict';

/* ================= utils ================= */
function stripBOM(s) { return s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s; }
// Decode raw file bytes: UTF-8 (strict) → UTF-16 via BOM → windows-1252
// fallback (classic Excel-exported CSVs are often latin1-encoded).
function decodeBytes(u8) {
  var b = u8 instanceof Uint8Array ? u8 : new Uint8Array(u8);
  if (b.length >= 2 && b[0] === 0xFF && b[1] === 0xFE) return new TextDecoder('utf-16le').decode(b.subarray(2));
  if (b.length >= 2 && b[0] === 0xFE && b[1] === 0xFF) return new TextDecoder('utf-16be').decode(b.subarray(2));
  try { return new TextDecoder('utf-8', { fatal: true }).decode(b); }
  catch (e) { return new TextDecoder('windows-1252').decode(b); }
}
function sanitizeCol(name, i, used) {
  var base = String(name == null ? '' : name).trim();
  if (!base) base = 'col_' + (i + 1);
  base = base.replace(/[\r\n\t]+/g, ' ').slice(0, 120);
  var out = base, n = 2;
  while (used.has(out)) out = base + '_' + (n++);
  used.add(out);
  return out;
}
function sanitizeColumns(names) {
  var used = new Set(), out = [];
  for (var i = 0; i < names.length; i++) out.push(sanitizeCol(names[i], i, used));
  return out;
}
function escXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') { try { return JSON.stringify(v); } catch (e) { return String(v); } }
  return String(v);
}
function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }
function clone(o) { try { return JSON.parse(JSON.stringify(o)); } catch (e) { return Array.isArray(o) ? o.slice() : Object.assign({}, o); } }
function safeString(v) { try { return JSON.stringify(v); } catch (e) { return String(v); } }

/* ================= delimited (CSV/TSV) ================= */
var DELIMS = [',', '\t', ';', '|'];
function detectDelimiter(text) {
  var lines = stripBOM(text).split(/\r?\n/).filter(function (l) { return l.trim(); }).slice(0, 5);
  if (!lines.length) return ',';
  var best = ',', bestScore = -1;
  DELIMS.forEach(function (d) {
    var counts = lines.map(function (l) { return splitRecord(l, d).length; });
    var sum = counts.reduce(function (a, b) { return a + b; }, 0);
    var avg = sum / counts.length;
    var variance = counts.reduce(function (a, c) { return a + Math.abs(c - avg); }, 0);
    var score = avg > 1 ? avg * 2 - variance : -1;
    if (score > bestScore) { bestScore = score; best = d; }
  });
  return best;
}
// split one physical line respecting quotes (no embedded newlines here)
function splitRecord(line, d) {
  var out = [], cur = '', q = false;
  for (var i = 0; i < line.length; i++) {
    var c = line[i];
    if (q) {
      if (c === '"') {
        if (line[i + 1] === '"') { cur += '"'; i++; } else q = false;
      } else cur += c;
    } else if (c === '"') q = true;
    else if (c === d) { out.push(cur); cur = ''; }
    else cur += c;
  }
  out.push(cur);
  return out;
}
function parseDelimited(text, opts) {
  opts = opts || {};
  text = stripBOM(String(text == null ? '' : text));
  var delim = opts.delimiter && opts.delimiter !== 'auto' ? opts.delimiter : detectDelimiter(text);
  var header = opts.header !== false;
  // full parse with embedded newlines inside quotes
  var records = [], cur = [], field = '', q = false;
  for (var i = 0; i < text.length; i++) {
    var c = text[i];
    if (q) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else q = false;
      } else field += c;
    } else if (c === '"') q = true;
    else if (c === delim) { cur.push(field); field = ''; }
    else if (c === '\r') { /* wait for \n */ }
    else if (c === '\n') { cur.push(field); field = ''; records.push(cur); cur = []; }
    else field += c;
  }
  cur.push(field);
  // trailing empty record from final newline?
  if (!(cur.length === 1 && cur[0] === '' && records.length)) records.push(cur);
  records = records.filter(function (r) { return !(r.length === 1 && String(r[0]).trim() === ''); });
  if (!records.length) return { columns: [], rows: [], delimiter: delim };
  var columns, startIdx = 0;
  if (header) {
    columns = sanitizeColumns(records[0].map(function (h) { return String(h).trim(); }));
    startIdx = 1;
  } else {
    var w = Math.max.apply(null, records.map(function (r) { return r.length; }));
    columns = sanitizeColumns(Array.apply(null, { length: w }).map(function (_, k) { return 'col_' + (k + 1); }));
  }
  var rows = [];
  for (var r = startIdx; r < records.length; r++) {
    var o = {};
    for (var k = 0; k < columns.length; k++) o[columns[k]] = k < records[r].length ? records[r][k] : '';
    rows.push(o);
  }
  return { columns: columns, rows: rows, delimiter: delim };
}
function csvEscape(v) {
  if (v === null || v === undefined) return '';
  var s = typeof v === 'object' ? safeString(v) : String(v);
  return /["\n\r,\t]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function buildDelimited(columns, rows, delim) {
  var lines = [columns.map(csvEscape).join(delim)];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    lines.push(columns.map(function (c) { return csvEscape(r[c]); }).join(delim));
  }
  return '﻿' + lines.join('\r\n');
}

/* ================= JSON (+lines) + flatten ================= */
function tryParseJSON(text) {
  text = stripBOM(String(text == null ? '' : text)).trim();
  if (!text) return { ok: false, error: 'Empty input.' };
  try { return { ok: true, value: JSON.parse(text), format: 'JSON' }; } catch (e1) {}
  var lines = text.split(/\r?\n/).map(function (l) { return l.trim(); }).filter(function (l) { return l && l.indexOf('//') !== 0; });
  if (lines.length > 1) {
    try { return { ok: true, value: lines.map(function (l) { return JSON.parse(l); }), format: 'JSON Lines' }; } catch (e2) {}
  }
  try { return { ok: true, value: JSON.parse(text.replace(/,\s*([}\]])/g, '$1')), format: 'JSON (trailing commas fixed)' }; } catch (e3) {}
  var msg = 'Invalid JSON.';
  try { JSON.parse(text); } catch (e) { msg = 'Invalid JSON: ' + (e && e.message ? e.message : e); }
  return { ok: false, error: msg };
}
function findArrayCandidates(root, maxDepth) {
  maxDepth = maxDepth || 4;
  var out = [];
  (function walk(node, path, depth) {
    if (depth > maxDepth || !isObj(node)) return;
    Object.keys(node).forEach(function (k) {
      var v = node[k], p = path.concat([k]);
      if (Array.isArray(v)) {
        out.push({ path: p.join('.'), length: v.length, objCount: v.filter(isObj).length });
        if (depth < maxDepth && v.length && isObj(v[0])) walk(v[0], p.concat(['[]']), depth + 1);
      } else if (isObj(v)) walk(v, p, depth + 1);
    });
  })(root, [], 0);
  return out;
}
function getByPath(root, dotted) {
  if (!dotted) return root;
  var parts = String(dotted).split('.').filter(function (p) { return p !== '[]'; });
  var cur = root;
  for (var i = 0; i < parts.length; i++) { if (cur == null) return undefined; cur = cur[parts[i]]; }
  return cur;
}
function coerceToRows(parsed, rootPath) {
  if (Array.isArray(parsed)) return { rows: parsed, format: 'array' };
  if (isObj(parsed)) {
    if (!rootPath) return { rows: [parsed], format: 'single object' };
    var v = getByPath(parsed, rootPath);
    if (Array.isArray(v)) return { rows: v, format: 'object → ' + rootPath + '[]' };
    return { rows: [parsed], format: 'single object' };
  }
  return { rows: [{ value: parsed }], format: 'primitive' };
}
function flattenValue(out, key, val, opts, depth) {
  var sep = opts.separator;
  if (depth > opts.maxDepth) { out[key] = safeString(val); return; }
  if (val === undefined) return;
  if (val === null) { out[key] = null; return; }
  var t = typeof val;
  if (t === 'string') { out[key] = opts.trimStrings === false ? val : val.trim(); return; }
  if (t === 'number' || t === 'boolean' || t === 'bigint') { out[key] = val; return; }
  if (Array.isArray(val)) {
    if (!val.length) { out[key] = null; return; }
    if (opts.arrayMode === 'json') { out[key] = JSON.stringify(val); return; }
    var hasObj = val.some(function (x) { return isObj(x) || Array.isArray(x); });
    if (!hasObj) {
      if (opts.arrayMode === 'indexed') val.forEach(function (el, i) { flattenValue(out, key + sep + i, el, opts, depth + 1); });
      else out[key] = val.map(function (x) { return x == null ? '' : String(x); }).join(opts.joinSeparator);
      return;
    }
    if (opts.arrayMode === 'indexed') {
      val.forEach(function (el, i) {
        if (isObj(el)) flattenObjectInto(out, el, key + sep + i, opts, depth + 1);
        else if (Array.isArray(el)) flattenValue(out, key + sep + i, el, opts, depth + 1);
        else out[key + sep + i] = el;
      });
    } else out[key] = val.map(function (x) { return (isObj(x) || Array.isArray(x)) ? JSON.stringify(x) : (x == null ? '' : String(x)); }).join(opts.joinSeparator);
    return;
  }
  if (isObj(val)) {
    var keys = Object.keys(val);
    if (!keys.length) { out[key] = null; return; }
    flattenObjectInto(out, val, key, opts, depth + 1);
    return;
  }
  out[key] = String(val);
}
function flattenObjectInto(out, obj, prefix, opts, depth) {
  Object.keys(obj).forEach(function (k) {
    flattenValue(out, prefix ? prefix + opts.separator + k : k, obj[k], opts, depth);
  });
}
function flattenRow(raw, opts) {
  var out = {};
  if (isObj(raw)) flattenObjectInto(out, raw, '', opts, 0);
  else if (Array.isArray(raw)) raw.forEach(function (el, i) { flattenValue(out, String(i), el, opts, 0); });
  else if (raw == null) out.value = null;
  else out.value = raw;
  return out;
}
function findExplodablePath(node, base, depth, maxDepth) {
  if (depth > maxDepth || node == null) return null;
  if (Array.isArray(node)) {
    if (node.length > 0 && node.length <= 10000 && node.some(isObj)) return base;
    return null;
  }
  if (isObj(node)) {
    var keys = Object.keys(node);
    for (var i = 0; i < keys.length; i++) {
      var v = node[keys[i]];
      if (Array.isArray(v) && v.length > 0 && v.some(isObj)) return base.concat([keys[i]]);
      var r = findExplodablePath(v, base.concat([keys[i]]), depth + 1, maxDepth);
      if (r) return r;
    }
  }
  return null;
}
function getAtPath(o, p) { var c = o; for (var i = 0; i < p.length; i++) { if (c == null) return undefined; c = c[p[i]]; } return c; }
function setAtPath(o, p, v) { var c = o; for (var i = 0; i < p.length - 1; i++) c = c[p[i]]; c[p[p.length - 1]] = v; }
function explodeRows(rawRows, enabled, cap) {
  cap = cap || 50000;
  if (!enabled) return { rows: rawRows, truncated: false };
  var out = [], truncated = false;
  for (var ri = 0; ri < rawRows.length; ri++) {
    var exploded = [rawRows[ri]], guard = 0, changed = true;
    while (changed && guard++ < 8) {
      changed = false;
      var next = [];
      for (var k = 0; k < exploded.length; k++) {
        var row = exploded[k];
        var p = (isObj(row) || Array.isArray(row)) ? findExplodablePath(row, [], 0, 6) : null;
        if (!p) { next.push(row); continue; }
        var arr = getAtPath(row, p);
        if (!Array.isArray(arr) || !arr.length || arr.length > 5000) { next.push(row); continue; }
        changed = true;
        for (var j = 0; j < arr.length; j++) {
          var c = clone(row);
          setAtPath(c, p, arr[j]);
          next.push(c);
          if (out.length + next.length >= cap) break;
        }
        if (out.length + next.length >= cap) break;
      }
      exploded = next;
      if (out.length + exploded.length >= cap) break;
    }
    for (var m = 0; m < exploded.length; m++) {
      if (out.length >= cap) { truncated = true; break; }
      out.push(exploded[m]);
    }
    if (out.length >= cap) { truncated = true; break; }
  }
  return { rows: out, truncated: truncated };
}
function defaultFlattenOpts() {
  return { separator: '.', arrayMode: 'join', joinSeparator: ', ', maxDepth: 10, trimStrings: true, explode: true };
}
function rowsToTable(rawRows, opts) {
  var ex = explodeRows(rawRows, opts.explode);
  var flat = ex.rows.map(function (r) { return flattenRow(r, opts); });
  var cols = [], seen = {};
  flat.forEach(function (row) {
    Object.keys(row).forEach(function (k) {
      if (!seen[k]) { seen[k] = 1; cols.push(k); }
    });
  });
  return { columns: cols, rows: flat, truncated: ex.truncated };
}

/* ================= XML (mini parser, portable) ================= */
var XML_ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: '\'', nbsp: '\u00a0', copy: '\u00a9', reg: '\u00ae', trade: '\u2122', euro: '\u20ac', pound: '\u00a3', yen: '\u00a5', deg: '\u00b0', plusmn: '\u00b1', middot: '\u00b7', bull: '\u2022', hellip: '\u2026', mdash: '\u2014', ndash: '\u2013', laquo: '\u00ab', raquo: '\u00bb', ldquo: '\u201c', rdquo: '\u201d', lsquo: '\u2018', rsquo: '\u2019', frasl: '\u2044' };
function decodeXmlEntities(s) {
  return String(s).replace(/&(#\d+|#x[0-9a-fA-F]+|[A-Za-z][A-Za-z0-9]+);/g, function (m, e) {
    if (e.charAt(0) === '#') {
      var cp = e.charAt(1) === 'x' || e.charAt(1) === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(cp); } catch (err) { return m; }
    }
    return XML_ENTITIES[e] !== undefined ? XML_ENTITIES[e] : m;
  });
}
function parseXmlTree(text) {
  text = stripBOM(String(text || '')).replace(/^\s*<\?xml[\s\S]*?\?>\s*/, '');
  var root = { name: '#root', attrs: {}, children: [] }, stack = [root];
  var re = /<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>|<\s*(\/?)\s*([A-Za-z_][\w\-.:]*)([^>]*?)(\/?)\s*>/g;
  var last = 0, m;
  function pushText(t) {
    if (!t || !t.trim()) return;
    var top = stack[stack.length - 1];
    top.children.push({ name: '#text', text: t });
  }
  while ((m = re.exec(text))) {
    pushText(text.slice(last, m.index));
    last = re.lastIndex;
    if (m[0].indexOf('<!--') === 0) continue;
    if (m[1] !== undefined) { // CDATA — literal, never entity-decoded
      var top2 = stack[stack.length - 1];
      if (m[1].trim()) top2.children.push({ name: '#text', text: m[1], cdata: true });
      continue;
    }
    var closing = !!m[2], name = m[3], attrS = m[4] || '', selfClose = !!m[5];
    if (closing) {
      for (var i = stack.length - 1; i > 0; i--) {
        if (stack[i].name === name) { stack.length = i; break; }
      }
      continue;
    }
    var attrs = {};
    var am = /([A-Za-z_][\w\-.:]*)\s*=\s*("([^"]*)"|'([^']*)')/g, a2;
    while ((a2 = am.exec(attrS))) attrs[a2[1]] = decodeXmlEntities(a2[3] !== undefined ? a2[3] : a2[4]);
    var el = { name: name, attrs: attrs, children: [] };
    stack[stack.length - 1].children.push(el);
    if (!selfClose) stack.push(el);
  }
  pushText(text.slice(last));
  // first element child is the document element
  for (var k = 0; k < root.children.length; k++) if (root.children[k].name !== '#text') return root.children[k];
  return null;
}
function elementToJS(el) {
  var obj = {};
  Object.keys(el.attrs || {}).forEach(function (k) { obj[k] = el.attrs[k]; });
  var groups = {}, order = [];
  (el.children || []).forEach(function (ch) {
    if (ch.name === '#text') return;
    if (!groups[ch.name]) { groups[ch.name] = []; order.push(ch.name); }
    groups[ch.name].push(ch);
  });
  order.forEach(function (tag) {
    var g = groups[tag];
    if (g.length === 1 && (!g[0].children || !g[0].children.length) && Object.keys(g[0].attrs || {}).length === 0) {
      obj[stripNs(tag)] = leafText(g[0]);
    } else if (g.length === 1) obj[stripNs(tag)] = elementToJS(g[0]);
    else obj[stripNs(tag)] = g.map(function (e) {
      return (!e.children || !e.children.length) && !Object.keys(e.attrs || {}).length ? leafText(e) : elementToJS(e);
    });
  });
  var t = [];
  (el.children || []).forEach(function (ch) { if (ch.name === '#text') t.push(ch.cdata ? ch.text : decodeXmlEntities(ch.text)); });
  var tJoined = t.join(' ').trim();
  if (!order.length && !Object.keys(obj).length) return tJoined;
  if (tJoined && order.length) obj._text = tJoined;
  return obj;
}
function stripNs(tag) { var i = tag.indexOf(':'); return i >= 0 ? tag.slice(i + 1) : tag; }
function leafText(el) {
  var t = [];
  (el.children || []).forEach(function (c) { if (c.name === '#text') t.push(c.cdata ? c.text : decodeXmlEntities(c.text)); });
  return t.join(' ').trim();
}
function xmlCandidates(tree) {
  var counts = {};
  (function walk(el, path) {
    if (!el || !el.children) return;
    var sibs = {};
    el.children.forEach(function (c) {
      if (c.name === '#text') return;
      var tag = stripNs(c.name);
      sibs[tag] = (sibs[tag] || 0) + 1;
    });
    Object.keys(sibs).forEach(function (tag) {
      if (sibs[tag] >= 2) {
        var p = path ? path + '/' + tag : tag;
        counts[p] = Math.max(counts[p] || 0, sibs[tag]);
      }
    });
    el.children.forEach(function (c) { if (c.name !== '#text') walk(c, path ? path + '/' + stripNs(c.name) : stripNs(c.name)); });
  })(tree, '');
  return Object.keys(counts).map(function (p) {
    var parts = p.split('/');
    return { path: p, tag: parts[parts.length - 1], count: counts[p] };
  }).sort(function (a, b) { return b.count - a.count; });
}
function findElementsByTag(root, tag) {
  var out = [];
  (function walk(el) {
    if (!el || !el.children) return;
    el.children.forEach(function (c) {
      if (c.name === '#text') return;
      if (stripNs(c.name) === tag) out.push(c);
      walk(c);
    });
  })(root);
  return out;
}
function xmlToTable(text, opts) {
  opts = opts || {};
  var tree = parseXmlTree(text);
  if (!tree) throw new Error('No XML elements found.');
  var cands = xmlCandidates(tree);
  var tag = opts.recordTag;
  if (!tag && cands.length) tag = cands[0].tag;
  var raws;
  if (tag) {
    var els = findElementsByTag(tree, tag);
    raws = els.map(elementToJS);
  } else {
    raws = [elementToJS(tree)];
  }
  var f = { separator: opts.separator || '.', arrayMode: opts.arrayMode || 'join', joinSeparator: ', ', maxDepth: opts.maxDepth || 10, trimStrings: true, explode: opts.explode !== false };
  var t = rowsToTable(raws, f);
  return { columns: t.columns, rows: t.rows, candidates: cands, recordTag: tag || '(whole document)', truncated: t.truncated };
}

/* ================= types / normalize ================= */
var DATE_RE = /^(\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?)?|\d{1,2}\/\d{1,2}\/\d{2,4}([ T]\d{1,2}:\d{2}(:\d{2})?)?|\d{1,2}-[A-Za-z]{3}-\d{2,4})$/;
function inferTypeOf(values) {
  var vals = values.filter(function (v) { return v !== null && v !== undefined && String(v).trim() !== ''; }).map(function (v) { return String(v).trim(); });
  if (!vals.length) return 'empty';
  function all(fn) { return vals.every(fn); }
  if (all(function (v) { return /^-?\d+$/.test(v); })) return 'integer';
  if (all(function (v) { return /^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(v); })) return 'number';
  if (all(function (v) { return /^(true|false)$/i.test(v); })) return 'boolean';
  if (all(function (v) { return DATE_RE.test(v) && !isNaN(Date.parse(v)); })) return 'date';
  return 'string';
}
function inferSchema(rows, columns) {
  var out = {};
  columns.forEach(function (c) {
    var vals = rows.map(function (r) { return r[c]; });
    var nonEmpty = vals.filter(function (v) { return v !== null && v !== undefined && String(v) !== ''; });
    var dist = {};
    nonEmpty.forEach(function (v) { var k = cellText(v); dist[k] = 1; });
    var nums = nonEmpty.map(Number).filter(function (n) { return isFinite(n); });
    out[c] = {
      type: inferTypeOf(vals),
      nulls: vals.length - nonEmpty.length,
      distinct: Object.keys(dist).length,
      min: nums.length === nonEmpty.length && nums.length ? Math.min.apply(null, nums) : null,
      max: nums.length === nonEmpty.length && nums.length ? Math.max.apply(null, nums) : null
    };
  });
  return out;
}
function castValue(v, type) {
  if (v === null || v === undefined || v === '') return v === '' ? '' : v;
  var s = String(v).trim();
  if (type === 'string') return String(v);
  if (type === 'integer') { var i = parseInt(s, 10); return isNaN(i) ? v : i; }
  if (type === 'number') { var n = Number(s); return isNaN(n) ? v : n; }
  if (type === 'boolean') {
    if (/^(true|1|yes|y)$/i.test(s)) return true;
    if (/^(false|0|no|n)$/i.test(s)) return false;
    return v;
  }
  if (type === 'date') { var t = Date.parse(s); return isNaN(t) ? v : new Date(t).toISOString().slice(0, 10); }
  return v;
}
function normalizeTable(table, actions) {
  // actions: [{op:'rename',from,to},{op:'cast',col,type},{op:'hide',col},{op:'trim'},{op:'dropEmptyRows'},{op:'dropEmptyCols'},{op:'dedupe'}]
  var cols = table.columns.slice(), rows = table.rows.map(function (r) { return Object.assign({}, r); });
  var hidden = {};
  (actions || []).forEach(function (a) {
    if (a.op === 'rename' && cols.indexOf(a.from) >= 0) {
      cols[cols.indexOf(a.from)] = a.to;
      rows.forEach(function (r) { r[a.to] = r[a.from]; delete r[a.from]; });
    } else if (a.op === 'cast') {
      rows.forEach(function (r) { r[a.col] = castValue(r[a.col], a.type); });
    } else if (a.op === 'hide') hidden[a.col] = 1;
    else if (a.op === 'trim') {
      rows.forEach(function (r) {
        Object.keys(r).forEach(function (k) { if (typeof r[k] === 'string') r[k] = r[k].trim(); });
      });
    } else if (a.op === 'dropEmptyRows') {
      rows = rows.filter(function (r) { return cols.some(function (c) { return r[c] !== null && r[c] !== undefined && String(r[c]).trim() !== ''; }); });
    } else if (a.op === 'dropEmptyCols') {
      cols = cols.filter(function (c) { return rows.some(function (r) { return r[c] !== null && r[c] !== undefined && String(r[c]).trim() !== ''; }); });
    } else if (a.op === 'dedupe') {
      var seen = {}, out = [];
      rows.forEach(function (r) {
        var k = JSON.stringify(cols.map(function (c) { return r[c] == null ? null : String(r[c]); }));
        if (!seen[k]) { seen[k] = 1; out.push(r); }
      });
      rows = out;
    }
  });
  cols = cols.filter(function (c) { return !hidden[c]; });
  rows.forEach(function (r) { Object.keys(r).forEach(function (k) { if (hidden[k]) delete r[k]; }); });
  return { name: table.name, columns: cols, rows: rows };
}

/* ================= joins ================= */
function joinTables(left, right, opts) {
  var lk = opts.leftKey, rk = opts.rightKey, type = opts.type || 'inner';
  function key(v) { return v === null || v === undefined ? '∅NULL∅' : String(v); }
  var idx = {};
  right.rows.forEach(function (r) {
    var k = key(r[rk]);
    (idx[k] = idx[k] || []).push(r);
  });
  var cols = left.columns.concat(right.columns.filter(function (c) { return c !== rk || rk !== lk || left.columns.indexOf(c) < 0; }));
  // disambiguate duplicate names (other than merged keys)
  var counts = {};
  cols.forEach(function (c) { counts[c] = (counts[c] || 0) + 1; });
  var colMapR = {};
  right.columns.forEach(function (c) {
    if (c === rk && left.columns.indexOf(lk) >= 0 && rk === lk) colMapR[c] = null; // merged key
    else if ((counts[c] || 0) > 1 && left.columns.indexOf(c) >= 0) colMapR[c] = c + '_right';
    else colMapR[c] = c;
  });
  var finalCols = left.columns.slice();
  right.columns.forEach(function (c) {
    if (colMapR[c] === null) return;
    finalCols.push(colMapR[c]);
  });
  var rows = [];
  left.rows.forEach(function (lr) {
    var hits = idx[key(lr[lk])] || [];
    if (!hits.length) {
      if (type === 'left' || type === 'full') {
        var o = Object.assign({}, lr);
        right.columns.forEach(function (c) { if (colMapR[c]) o[colMapR[c]] = null; });
        rows.push(o);
      }
    } else hits.forEach(function (rr) {
      var o2 = Object.assign({}, lr);
      right.columns.forEach(function (c) {
        if (colMapR[c] === null) return;
        o2[colMapR[c]] = rr[c];
      });
      rows.push(o2);
    });
  });
  // track matched right rows properly
  var matchedIdx = {};
  left.rows.forEach(function (lr) {
    (idx[key(lr[lk])] || []).forEach(function (rr) { matchedIdx[right.rows.indexOf(rr)] = 1; });
  });
  if (type === 'right' || type === 'full') {
    right.rows.forEach(function (rr, i) {
      if (matchedIdx[i]) return;
      var o = {};
      left.columns.forEach(function (c) { o[c] = null; });
      right.columns.forEach(function (c) { if (colMapR[c]) o[colMapR[c]] = rr[c]; });
      if (colMapR[rk] === null) o[lk] = rr[rk];
      rows.push(o);
    });
  }
  return { columns: finalCols, rows: rows };
}

/* ================= zip (stored) + xlsx export ================= */
function crcTable() {
  if (crcTable.t) return crcTable.t;
  var t = new Uint32Array(256);
  for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
  crcTable.t = t; return t;
}
function crc32(b) { var t = crcTable(), c = 0xFFFFFFFF; for (var i = 0; i < b.length; i++) c = t[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function s2b(s) { return new TextEncoder().encode(s); }
function cat(parts) { var l = 0, i; for (i = 0; i < parts.length; i++) l += parts[i].length; var o = new Uint8Array(l), p = 0; for (i = 0; i < parts.length; i++) { o.set(parts[i], p); p += parts[i].length; } return o; }
function u16(n) { return new Uint8Array([n & 0xFF, (n >>> 8) & 0xFF]); }
function u32(n) { return new Uint8Array([n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF]); }
function buildZipStored(files) {
  var enc = new TextEncoder(), local = [], central = [], off = 0;
  var d = new Date();
  var tm = ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF;
  var dt = ((((d.getFullYear() - 1980) & 0x7F) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF;
  files.forEach(function (f) {
    var nb = enc.encode(f.name), crc = crc32(f.data);
    var lh = cat([s2b('PK\x03\x04'), u16(20), u16(0x0800), u16(0), u16(tm), u16(dt), u32(crc), u32(f.data.length), u32(f.data.length), u16(nb.length), u16(0), nb]);
    local.push(lh, f.data);
    central.push(cat([s2b('PK\x01\x02'), u16(63), u16(20), u16(0x0800), u16(0), u16(tm), u16(dt), u32(crc), u32(f.data.length), u32(f.data.length), u16(nb.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(off), nb]));
    off += lh.length + f.data.length;
  });
  var c = cat(central), l = cat(local);
  return cat([l, c, cat([s2b('PK\x05\x06'), u16(0), u16(0), u16(files.length), u16(files.length), u32(c.length), u32(l.length), u16(0)])]);
}
function colLetter(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
function buildXlsx(columns, rows, sheetName) {
  sheetName = (sheetName || 'Sheet1').slice(0, 31) || 'Sheet1';
  var cols = columns.length ? columns : ['value'];
  var widths = cols.map(function (c) { return Math.min(50, Math.max(String(c).length, 12)); });
  rows.slice(0, 200).forEach(function (r) {
    cols.forEach(function (c, i) {
      var v = r[c], l = v == null ? 0 : String(v).length;
      if (l > widths[i]) widths[i] = Math.min(50, l);
    });
  });
  var colsXml = '<cols>' + cols.map(function (c, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + (widths[i] + 2).toFixed(1) + '" customWidth="1"/>'; }).join('') + '</cols>';
  var rx = '<row r="1" s="1">';
  cols.forEach(function (c, i) { rx += '<c r="' + colLetter(i) + '1" t="inlineStr" s="1"><is><t xml:space="preserve">' + escXml(c) + '</t></is></c>'; });
  rx += '</row>';
  rows.forEach(function (r, ri) {
    var rn = ri + 2;
    rx += '<row r="' + rn + '">';
    cols.forEach(function (c, ci) {
      var ref = colLetter(ci) + rn, v = r[c];
      if (v === undefined || v === null || v === '') { rx += '<c r="' + ref + '"/>'; return; }
      if (typeof v === 'bigint') v = v.toString();
      if (typeof v === 'number') {
        rx += isFinite(v) ? '<c r="' + ref + '" t="n"><v>' + v + '</v></c>' : '<c r="' + ref + '" t="inlineStr"><is><t>' + escXml(String(v)) + '</t></is></c>';
      } else if (typeof v === 'boolean') rx += '<c r="' + ref + '" t="b"><v>' + (v ? '1' : '0') + '</v></c>';
      else {
        var s = typeof v === 'object' ? safeString(v) : String(v);
        if (s.length > 32760) s = s.slice(0, 32760);
        s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
        rx += '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + escXml(s) + '</t></is></c>';
      }
    });
    rx += '</row>';
  });
  var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + colsXml +
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
    '<sheetFormat defaultRowHeight="15"/><sheetData>' + rx + '</sheetData>' +
    '<autoFilter ref="A1:' + colLetter(cols.length - 1) + (rows.length + 1) + '"/></worksheet>';
  var ct = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
  var rels = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>';
  var wbr = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../styles.xml"/></Relationships>';
  var wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + escXml(sheetName) + '" sheetId="1" r:id="rId1"/></sheets></workbook>';
  var st = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F6FEB"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFill="1" applyFont="1"/></cellXfs></styleSheet>';
  var core = '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>export</dc:title><dc:creator>local-data-lab</dc:creator></cp:coreProperties>';
  var appP = '<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>local-data-lab</Application></Properties>';
  return buildZipStored([
    { name: '[Content_Types].xml', data: s2b(ct) },
    { name: '_rels/.rels', data: s2b(rels) },
    { name: 'xl/workbook.xml', data: s2b(wb) },
    { name: 'xl/_rels/workbook.xml.rels', data: s2b(wbr) },
    { name: 'xl/worksheets/sheet1.xml', data: s2b(sheetXml) },
    { name: 'xl/styles.xml', data: s2b(st) },
    { name: 'docProps/core.xml', data: s2b(core) },
    { name: 'docProps/app.xml', data: s2b(appP) }
  ]);
}

/* ================= zip read + xlsx import ================= */
function dvGetStr(dv, off, len) { var s = ''; for (var i = 0; i < len; i++) s += String.fromCharCode(dv.getUint8(off + i)); return s; }
function parseZipEntries(buf) {
  var data = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  var dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  var entries = [];
  for (var i = data.length - 22; i >= 0; i--) {
    if (dvGetStr(dv, i, 4) === 'PK\x05\x06') break;
  }
  if (i < 0) throw new Error('Not a ZIP file.');
  var count = dv.getUint16(i + 10, true), cdOff = dv.getUint32(i + 16, true);
  var p = cdOff;
  var dec = new TextDecoder();
  for (var k = 0; k < count; k++) {
    if (dvGetStr(dv, p, 4) !== 'PK\x01\x02') throw new Error('Corrupt ZIP central directory.');
    var method = dv.getUint16(p + 10, true), compSize = dv.getUint32(p + 20, true), uncomp = dv.getUint32(p + 24, true);
    // central header: 46 bytes fixed; name/extra/comment lens at +28/+30/+32, local offset at +42
    var nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    var lho = dv.getUint32(p + 42, true);
    var name = dec.decode(data.subarray(p + 46, p + 46 + nlen));
    entries.push({ name: name, method: method, compSize: compSize, uncompSize: uncomp, localOffset: lho });
    p += 46 + nlen + elen + clen;
  }
  return { data: data, dv: dv, entries: entries };
}
function localDataSlice(zip, e) {
  var dv = zip.dv;
  if (dvGetStr(dv, e.localOffset, 4) !== 'PK\x03\x04') throw new Error('Corrupt ZIP local header for ' + e.name);
  var nl = dv.getUint16(e.localOffset + 26, true), el = dv.getUint16(e.localOffset + 28, true);
  var start = e.localOffset + 30 + nl + el;
  return zip.data.subarray(start, start + e.compSize);
}
function inflateRawAsync(comp, expectedSize) {
  // returns Promise<Uint8Array>
  try {
    var DS = typeof DecompressionStream !== 'undefined' ? DecompressionStream : null;
    if (DS) {
      var blob = new Blob([comp]);
      var stream = blob.stream().pipeThrough(new DS('deflate-raw'));
      return new Response(stream).arrayBuffer().then(function (ab) { return new Uint8Array(ab); });
    }
  } catch (e) { /* fall through */ }
  try {
    if (typeof require !== 'undefined') {
      var z = require('zlib');
      return Promise.resolve(new Uint8Array(z.inflateRawSync(Buffer.from(comp))));
    }
  } catch (e2) {}
  return Promise.reject(new Error('This ZIP uses deflate compression but the browser has no DecompressionStream support. Use a modern Chrome/Edge/Firefox/Safari.'));
}
function unzipAll(buf) {
  var zip = parseZipEntries(buf);
  var out = {};
  var chain = Promise.resolve();
  zip.entries.forEach(function (e) {
    if (/\/$/.test(e.name)) return;
    chain = chain.then(function () {
      var slice = localDataSlice(zip, e);
      if (e.method === 0) { out[e.name] = slice; return; }
      if (e.method === 8) {
        return inflateRawAsync(slice, e.uncompSize).then(function (d) { out[e.name] = d; });
      }
      throw new Error('Unsupported ZIP method ' + e.method + ' for ' + e.name);
    });
  });
  return chain.then(function () { return out; });
}
function xmlTagText(xml, tag) {
  var m = xml.match(new RegExp('<' + tag + '\\b[^>]*>([\\s\\S]*?)<\\/' + tag + '>'));
  return m ? m[1] : '';
}
function parseSharedStrings(xml) {
  var out = [], re = /<si>([\s\S]*?)<\/si>/g, m;
  while ((m = re.exec(xml))) {
    var inner = m[1], t = [], tm = /<t\b[^>]*>([\s\S]*?)<\/t>/g, t2;
    while ((t2 = tm.exec(inner))) t.push(t2[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"));
    out.push(t.join(''));
  }
  return out;
}
function colLettersToIdx(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64); return n - 1; }
function splitRef(ref) { var m = String(ref).match(/^([A-Z]+)(\d+)$/); return m ? { c: colLettersToIdx(m[1]), r: parseInt(m[2], 10) } : null; }
function parseXlsxStyles(xml) {
  // returns Set of xf indexes that look like dates
  var dateFmtIds = {}, isDateXf = {};
  var nf = /<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g, m;
  while ((m = nf.exec(xml))) {
    var code = m[2].toLowerCase();
    if (/[ymd]/.test(code) && /[ymdhs]/.test(code) && code.indexOf('red') < 0 && code !== 'general' && code !== '@') dateFmtIds[m[1]] = 1;
  }
  // builtin date ids
  [14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 57].forEach(function (id) { dateFmtIds[id] = 1; });
  var xfs = xmlTagText(xml, 'cellXfs');
  var xr = /<xf\b[^>]*>/g, idx = 0, xm;
  while ((xm = xr.exec(xfs))) {
    var idm = xm[0].match(/numFmtId="(\d+)"/);
    if (idm && dateFmtIds[idm[1]]) isDateXf[idx] = 1;
    idx++;
  }
  return isDateXf;
}
function excelSerialToISO(n) {
  var ms = Math.round((n - 25569) * 86400 * 1000);
  var d = new Date(ms);
  if (isNaN(d)) return n;
  var hasTime = Math.abs(n % 1) > 1e-9;
  var iso = d.toISOString();
  return hasTime ? iso.slice(0, 19).replace('T', ' ') : iso.slice(0, 10);
}
function parseSheetData(sheetXml, shared, isDateXf, dateConvert) {
  var re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, m;
  var cells = {};
  var maxR = 0, maxC = 0;
  while ((m = re.exec(sheetXml))) {
    var attrs = m[1] || '', inner = m[2] || '';
    var rm = attrs.match(/\br="([^"]+)"/), tm = attrs.match(/\bt="([^"]+)"/), sm = attrs.match(/\bs="(\d+)"/);
    if (!rm) continue;
    var ref = splitRef(rm[1]);
    if (!ref) continue;
    var t = tm ? tm[1] : 'n', s = sm ? parseInt(sm[1], 10) : 0;
    var v = null;
    var vm = inner.match(/<v>([\s\S]*?)<\/v>/);
    if (t === 'inlineStr') {
      var im = inner.match(/<t\b[^>]*>([\s\S]*?)<\/t>/);
      v = im ? im[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"') : '';
    } else if (t === 's' && vm) v = shared[parseInt(vm[1], 10)] != null ? shared[parseInt(vm[1], 10)] : '';
    else if (t === 'str' && vm) v = vm[1];
    else if (t === 'b' && vm) v = vm[1] === '1';
    else if (t === 'e') v = '';
    else if (vm) {
      var num = Number(vm[1]);
      if (isNaN(num)) v = vm[1];
      else if (dateConvert && isDateXf[s] && num > 20000 && num < 60000) v = excelSerialToISO(num);
      else v = num;
    } else v = '';
    cells[ref.r + ':' + ref.c] = v;
    if (ref.r > maxR) maxR = ref.r;
    if (ref.c > maxC) maxC = ref.c;
  }
  var grid = [];
  for (var r = 1; r <= maxR; r++) {
    var row = [];
    for (var c = 0; c <= maxC; c++) row.push(cells[r + ':' + c] !== undefined ? cells[r + ':' + c] : '');
    grid.push(row);
  }
  return grid;
}
function importXlsx(buf, opts) {
  opts = opts || {};
  var header = opts.header !== false, dateConvert = opts.dateConvert !== false;
  return unzipAll(buf).then(function (files) {
    var wbKeys = Object.keys(files).filter(function (k) { return /xl\/workbook\.xml$/i.test(k); });
    if (!wbKeys.length) throw new Error('workbook.xml not found — not an XLSX file?');
    var dec = new TextDecoder();
    var wb = dec.decode(files[wbKeys[0]]);
    var sheets = [], sr = /<sheet\b[^>]*>/g, sm;
    while ((sm = sr.exec(wb))) {
      var nm = sm[0].match(/\bname="([^"]*)"/), idm = sm[0].match(/r:id="([^"]*)"/);
      sheets.push({ name: nm ? nm[1] : 'Sheet', rid: idm ? idm[1] : '' });
    }
    var relsKey = Object.keys(files).filter(function (k) { return /xl\/_rels\/workbook\.xml\.rels$/i.test(k); })[0];
    var targets = {};
    if (relsKey) {
      var rels = dec.decode(files[relsKey]);
      var rr = /<Relationship\b[^>]*>/g, rm;
      while ((rm = rr.exec(rels))) {
        var im = rm[0].match(/\bId="([^"]*)"/), tm2 = rm[0].match(/\bTarget="([^"]*)"/);
        if (im && tm2) targets[im[1]] = tm2[1];
      }
    }
    var shared = [];
    var ssKey = Object.keys(files).filter(function (k) { return /sharedStrings\.xml$/i.test(k); })[0];
    if (ssKey) shared = parseSharedStrings(dec.decode(files[ssKey]));
    var stylesKey = Object.keys(files).filter(function (k) { return /xl\/styles\.xml$/i.test(k); })[0];
    var isDateXf = stylesKey ? parseXlsxStyles(dec.decode(files[stylesKey])) : {};
    var sheetFiles = Object.keys(files).filter(function (k) { return /xl\/worksheets\/sheet\d+\.xml$/i.test(k); }).sort();
    var out = [];
    sheets.forEach(function (sh, i) {
      var target = targets[sh.rid] ? 'xl/' + targets[sh.rid].replace(/^\//, '') : (sheetFiles[i] || sheetFiles[0]);
      target = target.replace('xl/xl/', 'xl/');
      var key = Object.keys(files).filter(function (k) { return k.toLowerCase() === target.toLowerCase(); })[0] || sheetFiles[i];
      if (!key) return;
      var grid = parseSheetData(dec.decode(files[key]), shared, isDateXf, dateConvert);
      if (!grid.length) return;
      var cols, start = 0;
      if (header) { cols = sanitizeColumns(grid[0].map(function (h) { return String(h == null ? '' : h); })); start = 1; }
      else cols = sanitizeColumns(grid[0].map(function (_, k) { return 'col_' + (k + 1); }));
      var rows = [];
      for (var r = start; r < grid.length; r++) {
        var o = {}, empty = true;
        for (var c = 0; c < cols.length; c++) {
          var v = c < grid[r].length ? grid[r][c] : '';
          if (v === undefined) v = '';
          if (v !== '' && v !== null) empty = false;
          o[cols[c]] = v;
        }
        if (!empty) rows.push(o);
      }
      out.push({ sheet: sh.name, columns: cols, rows: rows });
    });
    if (!out.length) throw new Error('No readable worksheets found.');
    return out;
  });
}

/* ================= mini SQL ================= */
var SQL_KW = { SELECT: 1, DISTINCT: 1, FROM: 1, WHERE: 1, GROUP: 1, BY: 1, ORDER: 1, LIMIT: 1, OFFSET: 1, AS: 1, AND: 1, OR: 1, NOT: 1, LIKE: 1, IN: 1, IS: 1, NULL: 1, TRUE: 1, FALSE: 1, JOIN: 1, INNER: 1, LEFT: 1, RIGHT: 1, FULL: 1, OUTER: 1, CROSS: 1, ON: 1, COUNT: 1, SUM: 1, AVG: 1, MIN: 1, MAX: 1, ASC: 1, DESC: 1 };
function sqlTokenize(s) {
  var toks = [], i = 0;
  function peek() { return s[i]; }
  while (i < s.length) {
    var c = s[i];
    if (/\s/.test(c)) { i++; continue; }
    if (c === '-' && s[i + 1] === '-') { while (i < s.length && s[i] !== '\n') i++; continue; }
    if (c === ',') { toks.push({ t: ',' }); i++; continue; }
    if (c === '(' || c === ')') { toks.push({ t: c }); i++; continue; }
    if (c === '*') { toks.push({ t: '*' }); i++; continue; }
    if (c === ';') { i++; continue; }
    if (c === "'" ) {
      var j = i + 1, v = '';
      while (j < s.length) {
        if (s[j] === "'" && s[j + 1] === "'") { v += "'"; j += 2; }
        else if (s[j] === "'") { j++; break; }
        else { v += s[j]; j++; }
      }
      toks.push({ t: 'str', v: v }); i = j; continue;
    }
    if (c === '"' || c === '`') {
      var q = c, k = i + 1, vv = '';
      while (k < s.length && s[k] !== q) { vv += s[k]; k++; }
      k++;
      toks.push({ t: 'ident', v: vv }); i = k; continue;
    }
    if (c === '[') {
      var k2 = s.indexOf(']', i);
      if (k2 < 0) throw new Error('Unclosed [identifier]');
      toks.push({ t: 'ident', v: s.slice(i + 1, k2) }); i = k2 + 1; continue;
    }
    if (/[0-9.]/.test(c) && /[0-9]/.test(c)) {
      var m = s.slice(i).match(/^\d+(\.\d+)?([eE][+-]?\d+)?/);
      toks.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue;
    }
    if (/[<>=!]/.test(c)) {
      var two = s.substr(i, 2);
      if (['<=', '>=', '!=', '<>'].indexOf(two) >= 0) { toks.push({ t: 'op', v: two === '<>' ? '!=' : two }); i += 2; continue; }
      if (c === '=') { toks.push({ t: 'op', v: '=' }); i++; continue; }
      if (c === '<' || c === '>') { toks.push({ t: 'op', v: c }); i++; continue; }
      throw new Error('Unexpected character: ' + c);
    }
    if (/[+\-%/]/.test(c)) { toks.push({ t: 'op', v: c }); i++; continue; }
    if (c === '.') { toks.push({ t: '.' }); i++; continue; }
    var wm = s.slice(i).match(/^[A-Za-z_][A-Za-z0-9_]*/);
    if (wm) {
      var w = wm[0], up = w.toUpperCase();
      toks.push(SQL_KW[up] ? { t: up } : { t: 'ident', v: w });
      i += w.length; continue;
    }
    throw new Error('Unexpected character in SQL: ' + c);
  }
  toks.push({ t: 'EOF' });
  return toks;
}
function sqlParse(query) {
  var toks = sqlTokenize(query), pos = 0;
  function tk() { return toks[pos]; }
  function eat(t) { if (toks[pos].t === t) { pos++; return true; } return false; }
  function expect(t) { if (toks[pos].t !== t) throw new Error('Expected ' + t + ' but found ' + toks[pos].t); pos++; }
  function parseExpr() { return parseOr(); }
  function parseOr() {
    var l = parseAnd();
    while (tk().t === 'OR') { pos++; l = { t: 'bin', op: 'OR', l: l, r: parseAnd() }; }
    return l;
  }
  function parseAnd() {
    var l = parseNot();
    while (tk().t === 'AND') { pos++; l = { t: 'bin', op: 'AND', l: l, r: parseNot() }; }
    return l;
  }
  function parseNot() {
    if (tk().t === 'NOT') { pos++; return { t: 'un', op: 'NOT', e: parseNot() }; }
    return parseCmp();
  }
  function parseCmp() {
    var l = parseAdd();
    var t = tk();
    if (t.t === 'op' && ['=', '!=', '<', '>', '<=', '>='].indexOf(t.v) >= 0) { pos++; return { t: 'bin', op: t.v, l: l, r: parseAdd() }; }
    if (t.t === 'NOT' && toks[pos + 1] && (toks[pos + 1].t === 'LIKE' || toks[pos + 1].t === 'IN')) {
      pos++;
      var k = toks[pos++].t;
      if (k === 'LIKE') return { t: 'like', e: l, pat: parseAdd(), not: true };
      expect('('); var list = parseLitList(); expect(')'); return { t: 'in', e: l, list: list, not: true };
    }
    if (t.t === 'LIKE') { pos++; return { t: 'like', e: l, pat: parseAdd(), not: false }; }
    if (t.t === 'IN') { pos++; expect('('); var l2 = parseLitList(); expect(')'); return { t: 'in', e: l, list: l2, not: false }; }
    if (t.t === 'IS') {
      pos++;
      var not = false;
      if (tk().t === 'NOT') { pos++; not = true; }
      expect('NULL');
      return { t: 'isnull', e: l, not: not };
    }
    return l;
  }
  function parseLitList() {
    var out = [];
    if (tk().t === ')') return out;
    out.push(parseExpr());
    while (eat(',')) out.push(parseExpr());
    return out;
  }
  function parseAdd() {
    var l = parseMul();
    while (tk().t === 'op' && (tk().v === '+' || tk().v === '-')) { var o = tk().v; pos++; l = { t: 'bin', op: o, l: l, r: parseMul() }; }
    return l;
  }
  function parseMul() {
    var l = parsePrim();
    while (tk().t === 'op' && (tk().v === '*' || tk().v === '/' || tk().v === '%')) { var o = tk().v; pos++; l = { t: 'bin', op: o, l: l, r: parsePrim() }; }
    return l;
  }
  function parsePrim() {
    var t = tk();
    if (t.t === '(') { pos++; var e = parseExpr(); expect(')'); return e; }
    if (t.t === '*') { pos++; return { t: 'star' }; }
    if (t.t === 'num') { pos++; return { t: 'lit', value: t.v }; }
    if (t.t === 'str') { pos++; return { t: 'lit', value: t.v }; }
    if (t.t === 'NULL') { pos++; return { t: 'lit', value: null }; }
    if (t.t === 'TRUE') { pos++; return { t: 'lit', value: true }; }
    if (t.t === 'FALSE') { pos++; return { t: 'lit', value: false }; }
    if (t.t === 'COUNT' || t.t === 'SUM' || t.t === 'AVG' || t.t === 'MIN' || t.t === 'MAX') {
      var fn = t.t; pos++; expect('(');
      var star = false, arg = null;
      if (tk().t === '*') { pos++; star = true; }
      else arg = parseExpr();
      expect(')');
      return { t: 'call', fn: fn, arg: arg, star: star };
    }
    if (t.t === 'ident') {
      var name = t.v; pos++;
      if (eat('.')) {
        if (tk().t === '*') { pos++; return { t: 'star', table: name }; }
        var c = tk();
        if (c.t !== 'ident' && SQL_KW[c.t]) { name = name; }
        else if (c.t === 'ident') { var cn = c.v; pos++; return { t: 'col', table: name, name: cn }; }
        else throw new Error('Expected column after .');
        return { t: 'col', table: name, name: c.t };
      }
      return { t: 'col', table: null, name: name };
    }
    if (SQL_KW[t.t]) return { t: 'col', table: null, name: t.t };
    throw new Error('Unexpected token in expression: ' + t.t);
  }
  function parseTableRef() {
    var t = tk(), name;
    if (t.t === 'ident') { name = t.v; pos++; }
    else if (SQL_KW[t.t]) { name = t.t; pos++; }
    else throw new Error('Expected table name, found ' + t.t);
    var alias = null;
    if (eat('AS')) {
      var a = tk();
      if (a.t === 'ident') { alias = a.v; pos++; } else if (SQL_KW[a.t]) { alias = a.t; pos++; }
      else throw new Error('Expected alias');
    } else if (tk().t === 'ident') { alias = tk().v; pos++; }
    return { table: name, alias: alias || name };
  }
  // SELECT ...
  expect('SELECT');
  var distinct = eat('DISTINCT');
  var select = [];
  if (tk().t === '*') { pos++; select.push({ expr: { t: 'star' }, alias: null }); }
  else {
    for (;;) {
      var e = parseExpr();
      var al = null;
      if (eat('AS')) {
        var at = tk();
        if (at.t === 'ident') { al = at.v; pos++; } else if (SQL_KW[at.t]) { al = at.t; pos++; }
        else throw new Error('Expected alias after AS');
      } else if (tk().t === 'ident' && toks[pos + 1] && [',', 'FROM', 'WHERE', 'GROUP', 'ORDER', 'LIMIT', 'OFFSET', 'EOF', ')'].indexOf(toks[pos + 1].t) >= 0) {
        al = tk().v; pos++;
      }
      select.push({ expr: e, alias: al });
      if (!eat(',')) break;
    }
  }
  expect('FROM');
  var from = parseTableRef();
  var joins = [];
  for (;;) {
    var jt = 'INNER', hasJoin = false;
    var t0 = tk().t;
    if (t0 === 'JOIN') { pos++; hasJoin = true; }
    else if (t0 === 'INNER' || t0 === 'LEFT' || t0 === 'RIGHT' || t0 === 'FULL' || t0 === 'CROSS') {
      pos++; jt = t0;
      if (tk().t === 'OUTER') pos++;
      if (tk().t === 'JOIN') { pos++; hasJoin = true; }
      else throw new Error('Expected JOIN');
    } else break;
    if (!hasJoin) break;
    var jr = parseTableRef();
    var on = null;
    if (jt !== 'CROSS') {
      if (tk().t === 'ON') { pos++; on = parseExpr(); }
      else if (jt !== 'INNER') { /* allow without ON => cross */ }
      else throw new Error('JOIN needs ON condition (or use CROSS JOIN)');
    }
    joins.push({ type: jt, table: jr.table, alias: jr.alias, on: on });
  }
  var where = null, groupBy = null, orderBy = null, limit = null, offset = null;
  if (eat('WHERE')) where = parseExpr();
  if (tk().t === 'GROUP') {
    pos++; expect('BY');
    groupBy = [];
    groupBy.push(parseExpr());
    while (eat(',')) groupBy.push(parseExpr());
  }
  if (tk().t === 'ORDER') {
    pos++; expect('BY');
    orderBy = [];
    for (;;) {
      var oe = parseExpr(), dir = 'ASC';
      if (tk().t === 'ASC') pos++;
      else if (tk().t === 'DESC') { pos++; dir = 'DESC'; }
      orderBy.push({ expr: oe, dir: dir });
      if (!eat(',')) break;
    }
  }
  if (eat('LIMIT')) {
    if (tk().t !== 'num') throw new Error('LIMIT needs a number');
    limit = Math.max(0, Math.floor(tk().v)); pos++;
  }
  if (eat('OFFSET')) {
    if (tk().t !== 'num') throw new Error('OFFSET needs a number');
    offset = Math.max(0, Math.floor(tk().v)); pos++;
  }
  return { distinct: distinct, select: select, from: from, joins: joins, where: where, groupBy: groupBy, orderBy: orderBy, limit: limit, offset: offset };
}
function sqlResolveTable(tables, name) {
  if (tables[name]) return { name: name, def: tables[name] };
  var keys = Object.keys(tables);
  for (var i = 0; i < keys.length; i++) if (keys[i].toLowerCase() === String(name).toLowerCase()) return { name: keys[i], def: tables[keys[i]] };
  throw new Error('Unknown table: ' + name);
}
function sqlGetCol(env, col) {
  if (col.table) {
    var ns = env[col.table] || env[col.table.toLowerCase()];
    if (!ns) {
      var ks = Object.keys(env);
      for (var i = 0; i < ks.length; i++) if (ks[i].toLowerCase() === String(col.table).toLowerCase()) { ns = env[ks[i]]; break; }
    }
    if (!ns) throw new Error('Unknown table alias: ' + col.table);
    var v = ns[col.name];
    if (v === undefined) {
      var kk = Object.keys(ns);
      for (var j = 0; j < kk.length; j++) if (kk[j].toLowerCase() === String(col.name).toLowerCase()) return ns[kk[j]];
      return null;
    }
    return v === undefined ? null : v;
  }
  // unqualified: search namespaces
  var found = 0, val = null;
  Object.keys(env).forEach(function (a) {
    var row = env[a];
    Object.keys(row).forEach(function (k) {
      if (k.toLowerCase() === String(col.name).toLowerCase()) { found++; val = row[k]; }
    });
  });
  if (found > 1) throw new Error('Ambiguous column: ' + col.name);
  if (!found) {
    if (col.name === '*') return null;
    return null;
  }
  return val;
}
function sqlToBool(v) { return !(v === null || v === undefined || v === false || v === 0 || v === ''); }
function sqlEq(a, b) {
  if (a === null || a === undefined || b === null || b === undefined || a === '' || b === '') {
    if ((a === '' || a == null) && (b === '' || b == null)) return true;
    return false;
  }
  if (typeof a === 'number' && typeof b === 'string' && String(b).trim() !== '' && !isNaN(Number(b))) b = Number(b);
  if (typeof b === 'number' && typeof a === 'string' && String(a).trim() !== '' && !isNaN(Number(a))) a = Number(a);
  return String(a).toLowerCase() === String(b).toLowerCase() && a == b ? true : String(a) === String(b) || a == b;
}
function sqlEval(e, env) {
  switch (e.t) {
    case 'lit': return e.value;
    case 'col': return sqlGetCol(env, e);
    case 'bin': {
      var l = sqlEval(e.l, env), r = sqlEval(e.r, env);
      switch (e.op) {
        case 'AND': return sqlToBool(l) && sqlToBool(r);
        case 'OR': return sqlToBool(l) || sqlToBool(r);
        case '=': return (l == null || r == null) && !(l == null && r == null && l === r) ? ((l == null || l === '') && (r == null || r === '') ? true : false) : sqlEq(l, r);
        case '!=': return !sqlEq(l, r);
        case '<': case '>': case '<=': case '>=': {
          if (l == null || r == null || l === '' || r === '') return false;
          var ln = Number(l), rn = Number(r), useN = isFinite(ln) && isFinite(rn) && String(l).trim() !== '' && String(r).trim() !== '';
          var a = useN ? ln : String(l), b = useN ? rn : String(r);
          if (e.op === '<') return a < b;
          if (e.op === '>') return a > b;
          if (e.op === '<=') return a <= b;
          return a >= b;
        }
        case '+': {
          if (l == null) l = 0; if (r == null) r = 0;
          if (typeof l === 'number' && typeof r === 'number') return l + r;
          var x = Number(l), y = Number(r);
          if (l !== '' && r !== '' && isFinite(x) && isFinite(y) && (typeof l === 'number' || typeof r === 'number' || String(l).match(/^-?[\d.]/))) return x + y;
          return String(l) + String(r);
        }
        case '-': return (Number(l) || 0) - (Number(r) || 0);
        case '*': return (Number(l) || 0) * (Number(r) || 0);
        case '/': return Number(r) ? (Number(l) || 0) / Number(r) : null;
        case '%': return Number(r) ? (Number(l) || 0) % Number(r) : null;
      }
      return null;
    }
    case 'un': return !sqlToBool(sqlEval(e.e, env));
    case 'like': {
      var v = sqlGetCol ? String(sqlEval(e.e, env) == null ? '' : sqlEval(e.e, env)) : '';
      var p = String(sqlEval(e.pat, env) == null ? '' : sqlEval(e.pat, env));
      var rx = '^' + p.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/%/g, '.*').replace(/_/g, '.') + '$';
      var ok = new RegExp(rx, 'i').test(v);
      return e.not ? !ok : ok;
    }
    case 'in': {
      var vv = sqlEval(e.e, env);
      var hit = e.list.some(function (x) { return sqlEq(vv, sqlEval(x, env)); });
      return e.not ? !hit : hit;
    }
    case 'isnull': {
      var z = sqlEval(e.e, env);
      var isN = z === null || z === undefined || z === '';
      return e.not ? !isN : isN;
    }
  }
  return null;
}
function sqlHasAgg(select, groupBy) {
  if (groupBy) return true;
  var found = false;
  function walk(e) {
    if (!e || found) return;
    if (e.t === 'call') { found = true; return; }
    ['l', 'r', 'e', 'pat', 'arg'].forEach(function (k) { if (e[k] && typeof e[k] === 'object') walk(e[k]); });
    (e.list || []).forEach(walk);
  }
  select.forEach(function (s) { walk(s.expr); });
  return found;
}
function sqlAgg(fn, vals, star) {
  var nn = vals.filter(function (v) { return v !== null && v !== undefined && v !== ''; });
  if (fn === 'COUNT') return star ? vals.length : nn.length;
  var nums = nn.map(Number).filter(isFinite);
  if (fn === 'SUM') return nums.reduce(function (a, b) { return a + b; }, 0);
  if (fn === 'AVG') return nums.length ? nums.reduce(function (a, b) { return a + b; }, 0) / nums.length : null;
  if (fn === 'MIN') {
    if (!nn.length) return null;
    return nn.reduce(function (a, b) { return (a < b ? a : b); });
  }
  if (fn === 'MAX') {
    if (!nn.length) return null;
    return nn.reduce(function (a, b) { return (a > b ? a : b); });
  }
  return null;
}
function exprLabel(e) {
  if (e.t === 'col') return e.table ? e.table + '.' + e.name : e.name;
  if (e.t === 'lit') return e.value === null ? 'NULL' : String(e.value);
  if (e.t === 'star') return '*';
  if (e.t === 'call') return e.fn + '(' + (e.star ? '*' : exprLabel(e.arg)) + ')';
  if (e.t === 'bin') return '(' + exprLabel(e.l) + ' ' + e.op + ' ' + exprLabel(e.r) + ')';
  if (e.t === 'un') return 'NOT(' + exprLabel(e.e) + ')';
  return 'expr';
}
function sqlExecute(ast, tables) {
  var from = sqlResolveTable(tables, ast.from.table);
  var aliases = {};
  aliases[ast.from.alias] = from.def;
  ast.joins.forEach(function (j) {
    var r = sqlResolveTable(tables, j.table);
    aliases[j.alias] = r.def;
  });
  // base envs
  var envs = from.def.rows.map(function (r) { var o = {}; o[ast.from.alias] = r; return o; });
  ast.joins.forEach(function (j) {
    var right = aliases[j.alias];
    var next = [];
    if (!j.on && (j.type === 'CROSS' || j.type === 'INNER')) {
      envs.forEach(function (e) { right.rows.forEach(function (rr) { var o = Object.assign({}, e); o[j.alias] = rr; next.push(o); }); });
    } else if (j.on && j.on.t === 'bin' && j.on.op === '=' && j.on.l.t === 'col' && j.on.r.t === 'col') {
      // hash path
      var idx = {};
      right.rows.forEach(function (rr) {
        var o = {}; o[j.alias] = rr;
        var k = keyOf(sqlEval(j.on.r, Object.assign({}, o)));
        (idx[k] = idx[k] || []).push(rr);
      });
      envs.forEach(function (e) {
        var k2 = keyOf(sqlEval(j.on.l, e));
        var hits = idx[k2] || [];
        if (!hits.length) {
          if (j.type === 'LEFT' || j.type === 'FULL') { var o = Object.assign({}, e); o[j.alias] = nullRow(right.columns); next.push(o); }
        } else hits.forEach(function (rr) { var o2 = Object.assign({}, e); o2[j.alias] = rr; next.push(o2); });
      });
      if (j.type === 'RIGHT' || j.type === 'FULL') {
        var matched = {};
        envs.forEach(function (e) {
          var k3 = keyOf(sqlEval(j.on.l, e));
          (idx[k3] || []).forEach(function (rr) { matched[right.rows.indexOf(rr)] = 1; });
        });
        right.rows.forEach(function (rr, i) {
          if (matched[i]) return;
          var o = {};
          Object.keys(aliases).forEach(function (a) { o[a] = a === j.alias ? rr : nullRow(aliases[a].columns); });
          next.push(o);
        });
      }
    } else {
      envs.forEach(function (e) {
        var any = false;
        right.rows.forEach(function (rr) {
          var o = Object.assign({}, e); o[j.alias] = rr;
          var ok = j.on ? sqlToBool(sqlEval(j.on, o)) : true;
          if (ok) { any = true; next.push(o); }
        });
        if (!any && (j.type === 'LEFT' || j.type === 'FULL')) { var o2 = Object.assign({}, e); o2[j.alias] = nullRow(right.columns); next.push(o2); }
      });
      if (j.type === 'RIGHT' || j.type === 'FULL') {
        // approximate: rows that match nothing
        right.rows.forEach(function (rr) {
          var found = false;
          envs.forEach(function (e) {
            var o = Object.assign({}, e); o[j.alias] = rr;
            if (j.on ? sqlToBool(sqlEval(j.on, o)) : true) found = true;
          });
          if (!found) {
            var o3 = {};
            Object.keys(aliases).forEach(function (a) { o3[a] = a === j.alias ? rr : nullRow(aliases[a].columns); });
            next.push(o3);
          }
        });
      }
    }
    envs = next;
  });
  if (ast.where) envs = envs.filter(function (e) { return sqlToBool(sqlEval(ast.where, e)); });
  var aggMode = sqlHasAgg(ast.select, ast.groupBy);
  var outCols = [], outRows = [];
  function outName(s) { return s.alias || exprLabel(s.expr); }
  if (!aggMode) {
    // expand stars
    ast.select.forEach(function (s) {
      if (s.expr.t === 'star' && !s.expr.table) {
        Object.keys(aliases).forEach(function (a) { aliases[a].columns.forEach(function (c) { outCols.push(s.alias || (Object.keys(aliases).length > 1 ? a + '.' + c : c)); }); });
      } else if (s.expr.t === 'star' && s.expr.table) {
        var al = Object.keys(aliases).filter(function (a) { return a.toLowerCase() === String(s.expr.table).toLowerCase(); })[0] || s.expr.table;
        aliases[al].columns.forEach(function (c) { outCols.push(s.alias || c); });
      } else outCols.push(outName(s));
    });
    envs.forEach(function (e) {
      var row = {};
      var ci = 0;
      ast.select.forEach(function (s) {
        if (s.expr.t === 'star' && !s.expr.table) {
          Object.keys(aliases).forEach(function (a) {
            aliases[a].columns.forEach(function (c) { row[outCols[ci++]] = e[a] ? e[a][c] : null; });
          });
        } else if (s.expr.t === 'star' && s.expr.table) {
          var al = Object.keys(aliases).filter(function (a) { return a.toLowerCase() === String(s.expr.table).toLowerCase(); })[0] || s.expr.table;
          aliases[al].columns.forEach(function (c) { row[outCols[ci++]] = e[al] ? e[al][c] : null; });
        } else row[outCols[ci++]] = sqlEval(s.expr, e);
      });
      outRows.push(row);
    });
  } else {
    ast.select.forEach(function (s) { outCols.push(outName(s)); });
    var groups = {};
    var order = [];
    if (!ast.groupBy) {
      groups.__all__ = envs; order.push('__all__');
    } else {
      envs.forEach(function (e) {
        var k = JSON.stringify(ast.groupBy.map(function (g) { var v = sqlEval(g, e); return v === undefined ? null : v; }));
        if (!groups[k]) { groups[k] = []; order.push(k); }
        groups[k].push(e);
      });
    }
    order.forEach(function (k) {
      var g = groups[k], row = {};
      ast.select.forEach(function (s, si) {
        row[outCols[si]] = evalSelect(s.expr, g);
      });
      outRows.push(row);
    });
  }
  function evalSelect(expr, genvs) {
    if (expr.t === 'call') {
      var vals = expr.star ? genvs.map(function () { return 1; }) : genvs.map(function (e) { return sqlEval(expr.arg, e); });
      return sqlAgg(expr.fn, vals, expr.star);
    }
    if (expr.t === 'star') return null;
    // bare column in grouped query: take from first row
    if (expr.t === 'col') return sqlEval(expr, genvs[0] || {});
    // mixed expr with agg inside? evaluate per-row then? Simplify: if contains agg, unsupported nesting -> eval on first
    return sqlEval(expr, genvs[0] || {});
  }
  if (ast.distinct) {
    var seen = {}, uni = [];
    outRows.forEach(function (r) {
      var k = JSON.stringify(outCols.map(function (c) { return r[c]; }));
      if (!seen[k]) { seen[k] = 1; uni.push(r); }
    });
    outRows = uni;
  }
  if (ast.orderBy) {
    outRows.sort(function (ra, rb) {
      for (var i = 0; i < ast.orderBy.length; i++) {
        var o = ast.orderBy[i];
        var va = orderVal(o.expr, ra), vb = orderVal(o.expr, rb);
        var c = cmpVals(va, vb);
        if (c) return o.dir === 'DESC' ? -c : c;
      }
      return 0;
    });
  }
  function orderVal(expr, row) {
    var env = {};
    Object.keys(row).forEach(function (k) { env.__out__ = env.__out__ || {}; env.__out__[k] = row[k]; });
    // allow ordering by output alias
    if (expr.t === 'col' && !expr.table) {
      var hit = null;
      outCols.forEach(function (c) { if (c.toLowerCase() === String(expr.name).toLowerCase()) hit = c; });
      if (hit) return row[hit];
    }
    try { return sqlEval(expr, { __out__: row }); } catch (e) { return null; }
  }
  function cmpVals(a, b) {
    if (a == null || a === '') {
      if (b == null || b === '') return 0;
      return -1;
    }
    if (b == null || b === '') return 1;
    var an = Number(a), bn = Number(b);
    if (isFinite(an) && isFinite(bn) && String(a).trim() !== '' && String(b).trim() !== '') return an < bn ? -1 : an > bn ? 1 : 0;
    var as = String(a), bs = String(b);
    return as < bs ? -1 : as > bs ? 1 : 0;
  }
  var off = ast.offset || 0;
  if (off) outRows = outRows.slice(off);
  if (ast.limit != null) outRows = outRows.slice(0, ast.limit);
  return { columns: outCols, rows: outRows };
}
function keyOf(v) { return v === null || v === undefined ? '∅' : String(v); }
function nullRow(cols) { var o = {}; cols.forEach(function (c) { o[c] = null; }); return o; }
function runSQL(query, tables) {
  var ast = sqlParse(query);
  return { ast: ast, result: sqlExecute(ast, tables) };
}

/* ================= charts ================= */
var PALETTE = ['#4da3ff', '#7c5cff', '#2ecc71', '#ff9f43', '#ff6b6b', '#48dbfb', '#f368e0', '#1dd1a1', '#feca57', '#5f27cd', '#ff7f50', '#00d2d3'];
function aggFn(name, vals) {
  var nn = vals.filter(function (v) { return v !== null && v !== undefined && v !== '' && ! (typeof v === 'number' && !isFinite(v)); });
  if (name === 'count') return vals.length;
  var nums = nn.map(Number).filter(isFinite);
  if (!nums.length) return 0;
  if (name === 'sum') return nums.reduce(function (a, b) { return a + b; }, 0);
  if (name === 'avg') return nums.reduce(function (a, b) { return a + b; }, 0) / nums.length;
  if (name === 'min') return Math.min.apply(null, nums);
  if (name === 'max') return Math.max.apply(null, nums);
  return nums.reduce(function (a, b) { return a + b; }, 0);
}
function prepareBar(rows, x, y, agg, limit) {
  var groups = {}, order = [];
  rows.forEach(function (r) {
    var k = cellText(r[x]);
    if (k === '') k = '(blank)';
    if (!groups[k]) { groups[k] = []; order.push(k); }
    groups[k].push(y ? r[y] : 1);
  });
  var items = order.map(function (k) { return { label: k, value: aggFn(agg, groups[k]), count: groups[k].length }; });
  items.sort(function (a, b) { return b.value - a.value; });
  return items.slice(0, limit || 15);
}
function prepareXY(rows, x, y, limit) {
  var pts = [];
  rows.forEach(function (r) {
    var xv = r[x], yv = r[y];
    if (xv == null || xv === '' || yv == null || yv === '') return;
    pts.push({ x: xv, y: Number(yv), label: cellText(xv), rawX: xv });
  });
  pts = pts.filter(function (p) { return isFinite(p.y); });
  pts.sort(function (a, b) {
    var an = Number(a.rawX), bn = Number(b.rawX);
    if (isFinite(an) && isFinite(bn)) return an - bn;
    return String(a.rawX) < String(b.rawX) ? -1 : 1;
  });
  return pts.slice(0, limit || 200);
}
function niceTicks(max, n) {
  if (!(max > 0)) return [0, 1];
  var step = max / n, mag = Math.pow(10, Math.floor(Math.log10(step))), norm = step / mag, s;
  s = norm >= 5 ? 10 : norm >= 2 ? 5 : norm >= 1 ? 2 : 1;
  step = s * mag;
  var ticks = [];
  for (var v = 0; v <= max * 1.02; v += step) ticks.push(Math.round(v * 100) / 100);
  return ticks;
}
function buildChartSVG(cfg) {
  var W = 680, H = 380, P = { l: 56, r: 14, t: 30, b: 64 };
  var iw = W - P.l - P.r, ih = H - P.t - P.b;
  var out = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + W + ' ' + H + '" role="img">';
  out += '<rect x="0" y="0" width="' + W + '" height="' + H + '" fill="transparent"/>';
  function grid(max) {
    var s = '';
    niceTicks(max, 5).forEach(function (t) {
      var y = P.t + ih - (t / (max || 1)) * ih;
      s += '<line x1="' + P.l + '" y1="' + y + '" x2="' + (W - P.r) + '" y2="' + y + '" stroke="currentColor" stroke-opacity="0.15"/>';
      s += '<text x="' + (P.l - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="currentColor" opacity="0.7">' + escXml(fmtNum(t)) + '</text>';
    });
    return s;
  }
  function fmtNum(n) {
    if (Math.abs(n) >= 1000000) return (n / 1000000).toFixed(1) + 'M';
    if (Math.abs(n) >= 1000) return (n / 1000).toFixed(1) + 'k';
    return String(Math.round(n * 100) / 100);
  }
  if (cfg.type === 'bar') {
    var items = cfg.items || [];
    var max = Math.max.apply(null, [0].concat(items.map(function (d) { return d.value; })));
    out += grid(max);
    var bw = items.length ? iw / items.length : iw;
    items.forEach(function (d, i) {
      var h = max ? (d.value / max) * ih : 0;
      var x = P.l + i * bw + bw * 0.15, w = bw * 0.7, y = P.t + ih - h;
      out += '<rect x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" width="' + w.toFixed(1) + '" height="' + Math.max(1, h).toFixed(1) + '" rx="3" fill="' + PALETTE[i % PALETTE.length] + '"><title>' + escXml(d.label) + ': ' + escXml(String(Math.round(d.value * 100) / 100)) + ' (n=' + d.count + ')</title></rect>';
      var lb = d.label.length > 12 ? d.label.slice(0, 11) + '…' : d.label;
      out += '<text x="' + (x + w / 2).toFixed(1) + '" y="' + (P.t + ih + 14) + '" text-anchor="' + (items.length > 8 ? 'end' : 'middle') + '" font-size="10" fill="currentColor" transform="' + (items.length > 8 ? 'rotate(-30 ' + (x + w / 2).toFixed(1) + ' ' + (P.t + ih + 14) + ')' : '') + '">' + escXml(lb) + '</text>';
      if (h > 14) out += '<text x="' + (x + w / 2).toFixed(1) + '" y="' + (y + 13) + '" text-anchor="middle" font-size="10" fill="#fff">' + escXml(fmtNum(d.value)) + '</text>';
    });
    out += '<text x="' + (P.l) + '" y="16" font-size="13" font-weight="bold" fill="currentColor">' + escXml(cfg.title || '') + '</text>';
  } else if (cfg.type === 'line') {
    var pts = cfg.points || [];
    var ys = pts.map(function (p) { return p.y; });
    var max2 = Math.max.apply(null, [0].concat(ys)), min2 = Math.min.apply(null, [0].concat(ys));
    var span = (max2 - min2) || 1;
    out += grid(max2);
    var step = pts.length > 1 ? iw / (pts.length - 1) : 0;
    var dAttr = pts.map(function (p, i) {
      var x = P.l + i * step, y = P.t + ih - ((p.y - min2) / span) * ih;
      return (i ? 'L' : 'M') + x.toFixed(1) + ' ' + y.toFixed(1);
    }).join(' ');
    if (pts.length) {
      out += '<path d="' + dAttr + '" fill="none" stroke="' + PALETTE[0] + '" stroke-width="2.5"/>';
      pts.forEach(function (p, i) {
        var x = P.l + i * step, y = P.t + ih - ((p.y - min2) / span) * ih;
        out += '<circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="3.5" fill="' + PALETTE[0] + '"><title>' + escXml(p.label) + ': ' + escXml(String(p.y)) + '</title></circle>';
        if (pts.length <= 12) out += '<text x="' + x.toFixed(1) + '" y="' + (P.t + ih + 14) + '" text-anchor="middle" font-size="9" fill="currentColor">' + escXml(String(p.label).slice(0, 10)) + '</text>';
      });
    }
    out += '<text x="' + P.l + '" y="16" font-size="13" font-weight="bold" fill="currentColor">' + escXml(cfg.title || '') + '</text>';
  } else if (cfg.type === 'scatter') {
    var ps = cfg.points || [];
    var xs = ps.map(function (p) { return Number(p.rawX); }).filter(isFinite);
    var x0 = Math.min.apply(null, [0].concat(xs)), x1 = Math.max.apply(null, [1].concat(xs));
    var y0 = 0, y1 = Math.max.apply(null, [1].concat(ps.map(function (p) { return p.y; })));
    out += grid(y1);
    ps.forEach(function (p, i) {
      var xn = Number(p.rawX);
      if (!isFinite(xn)) return;
      var x = P.l + ((xn - x0) / ((x1 - x0) || 1)) * iw, y = P.t + ih - (p.y / (y1 || 1)) * ih;
      out += '<circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="4" fill="' + PALETTE[i % PALETTE.length] + '" fill-opacity="0.8"><title>' + escXml(p.label + ', ' + p.y) + '</title></circle>';
    });
    out += '<text x="' + P.l + '" y="16" font-size="13" font-weight="bold" fill="currentColor">' + escXml(cfg.title || '') + '</text>';
  } else if (cfg.type === 'pie') {
    var it = (cfg.items || []).slice(0, 12);
    var tot = it.reduce(function (a, d) { return a + d.value; }, 0) || 1;
    var cx = P.l + iw / 2 - 60, cy = P.t + ih / 2, R = Math.min(iw, ih) / 2 - 6, a0 = -Math.PI / 2;
    it.forEach(function (d, i) {
      var a1 = a0 + (d.value / tot) * Math.PI * 2;
      var big = (a1 - a0) > Math.PI ? 1 : 0;
      var x0 = cx + R * Math.cos(a0), y0 = cy + R * Math.sin(a0), x1 = cx + R * Math.cos(a1), y1 = cy + R * Math.sin(a1);
      out += '<path d="M' + cx + ' ' + cy + ' L' + x0.toFixed(1) + ' ' + y0.toFixed(1) + ' A' + R + ' ' + R + ' 0 ' + big + ' 1 ' + x1.toFixed(1) + ' ' + y1.toFixed(1) + ' Z" fill="' + PALETTE[i % PALETTE.length] + '" stroke="#fff" stroke-width="1"><title>' + escXml(d.label + ': ' + (Math.round(d.value / tot * 1000) / 10) + '%') + '</title></path>';
      a0 = a1;
    });
    it.forEach(function (d, i) {
      var ly = P.t + i * 18;
      out += '<rect x="' + (W - P.r - 150) + '" y="' + ly + '" width="11" height="11" fill="' + PALETTE[i % PALETTE.length] + '"/>';
      out += '<text x="' + (W - P.r - 135) + '" y="' + (ly + 10) + '" font-size="10" fill="currentColor">' + escXml((d.label.length > 16 ? d.label.slice(0, 15) + '…' : d.label) + ' ' + (Math.round(d.value / tot * 1000) / 10) + '%') + '</text>';
    });
    out += '<text x="' + P.l + '" y="16" font-size="13" font-weight="bold" fill="currentColor">' + escXml(cfg.title || '') + '</text>';
  }
  out += '</svg>';
  return out;
}

/* ================= samples ================= */
var SAMPLES = {
  employees: 'id,name,dept,city,salary,active\n1,Ada Lovelace,Eng,London,95000,true\n2,Grace Hopper,Eng,New York,110000,true\n3,Alan Turing,Research,Manchester,88000,false\n4,Katherine Johnson,Research,Virginia,92000,true\n5,Linus Torvalds,Eng,Helsinki,99000,true\n6,Adaeze Nwosu,Sales,Lagos,61000,true\n7,Jo Silva,Sales,Lisbon,58000,false\n8,Kim Park,Eng,Seoul,87000,true',
  events: 'ts\tuser\taction\tms\n2026-01-01T10:00:00Z\tali\tlogin\t120\n2026-01-01T10:01:00Z\tbea\tpurchase\t340\n2026-01-01T10:02:00Z\tali\tpurchase\t210\n2026-01-01T10:03:00Z\tcid\tlogin\t95\n2026-01-01T10:04:00Z\tbea\tlogout\t40\n2026-01-01T10:05:00Z\tali\tlogout\t55',
  orders: '{"shop":"demo","orders":[{"orderId":"A1","customer":{"name":"Kim","vip":true},"items":[{"sku":"p1","qty":2,"price":9.99},{"sku":"p2","qty":1,"price":24.5}],"paid":true},{"orderId":"A2","customer":{"name":"Jo"},"items":[{"sku":"p3","qty":5,"price":3.2}],"paid":false},{"orderId":"A3","customer":{"name":"Ali","vip":false},"items":[],"paid":true}]}',
  catalog: '<?xml version="1.0"?><catalog><product sku="p1" stock="12"><name>Keyboard</name><price>49.99</price><tags><tag>peripheral</tag><tag>wired</tag></tags></product><product sku="p2" stock="0"><name>Mouse</name><price>25.5</price><tags><tag>peripheral</tag></tags></product><product sku="p3" stock="7"><name>Monitor</name><price>189.0</price></product></catalog>'
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    stripBOM: stripBOM, decodeBytes: decodeBytes, decodeXmlEntities: decodeXmlEntities, sanitizeColumns: sanitizeColumns, cellText: cellText,
    detectDelimiter: detectDelimiter, parseDelimited: parseDelimited, buildDelimited: buildDelimited,
    tryParseJSON: tryParseJSON, findArrayCandidates: findArrayCandidates, coerceToRows: coerceToRows,
    flattenRow: flattenRow, explodeRows: explodeRows, defaultFlattenOpts: defaultFlattenOpts, rowsToTable: rowsToTable,
    parseXmlTree: parseXmlTree, elementToJS: elementToJS, xmlCandidates: xmlCandidates, xmlToTable: xmlToTable,
    inferTypeOf: inferTypeOf, inferSchema: inferSchema, castValue: castValue, normalizeTable: normalizeTable,
    joinTables: joinTables,
    buildZipStored: buildZipStored, buildXlsx: buildXlsx,
    parseZipEntries: parseZipEntries, unzipAll: unzipAll, importXlsx: importXlsx, excelSerialToISO: excelSerialToISO,
    sqlTokenize: sqlTokenize, sqlParse: sqlParse, sqlExecute: sqlExecute, runSQL: runSQL,
    prepareBar: prepareBar, prepareXY: prepareXY, buildChartSVG: buildChartSVG,
    SAMPLES: SAMPLES, escXml: escXml
  };
} else if (typeof window !== 'undefined') window.DataLab = null; // attached below
if (typeof window !== 'undefined') {
  window.DataLabEngine = {
    stripBOM: stripBOM, decodeBytes: decodeBytes, decodeXmlEntities: decodeXmlEntities, sanitizeColumns: sanitizeColumns, cellText: cellText,
    detectDelimiter: detectDelimiter, parseDelimited: parseDelimited, buildDelimited: buildDelimited,
    tryParseJSON: tryParseJSON, findArrayCandidates: findArrayCandidates, coerceToRows: coerceToRows,
    flattenRow: flattenRow, explodeRows: explodeRows, defaultFlattenOpts: defaultFlattenOpts, rowsToTable: rowsToTable,
    parseXmlTree: parseXmlTree, elementToJS: elementToJS, xmlCandidates: xmlCandidates, xmlToTable: xmlToTable,
    inferTypeOf: inferTypeOf, inferSchema: inferSchema, castValue: castValue, normalizeTable: normalizeTable,
    joinTables: joinTables,
    buildZipStored: buildZipStored, buildXlsx: buildXlsx,
    parseZipEntries: parseZipEntries, unzipAll: unzipAll, importXlsx: importXlsx, excelSerialToISO: excelSerialToISO,
    sqlTokenize: sqlTokenize, sqlParse: sqlParse, sqlExecute: sqlExecute, runSQL: runSQL,
    prepareBar: prepareBar, prepareXY: prepareXY, buildChartSVG: buildChartSVG,
    SAMPLES: SAMPLES, escXml: escXml
  };
}
})();
