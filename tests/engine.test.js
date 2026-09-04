// Committed engine tests — run with `npm test`. Pure node, no browser needed.
const E = require('../engine.js');
let failures = 0;
const assert = (c, m) => { if (!c) { failures++; console.error('FAIL:', m); } else console.log('ok:', m); };
(async () => {
  const d = E.parseDelimited('a,b\n1,"x,y"\n2,3', { delimiter: ',', header: true });
  assert(d.columns.join() === 'a,b' && d.rows[0].b === 'x,y', 'csv quotes');
  assert(E.detectDelimiter('a\tb\n1\t2') === '\t', 'tsv detect');

  const j = E.tryParseJSON(E.SAMPLES.orders);
  assert(j.ok, 'orders json parses');
  const c = E.coerceToRows(j.value, 'orders');
  const t = E.rowsToTable(c.rows, E.defaultFlattenOpts());
  assert(t.rows.length === 4 && t.columns.includes('items.sku'), 'explode line items');

  const xt = E.xmlToTable(E.SAMPLES.catalog, {});
  assert(xt.recordTag === 'product' && xt.rows.length === 3, 'xml records');
  assert(xt.columns.includes('price'), 'xml cols');

  const s = E.inferSchema([{ a: '1' }, { a: '2' }], ['a']);
  assert(s.a.type === 'integer', 'infer integer');
  const nt = E.normalizeTable({ name: 't', columns: ['a'], rows: [{ a: ' 1 ' }, { a: '1' }] }, [{ op: 'trim' }, { op: 'cast', col: 'a', type: 'integer' }, { op: 'dedupe' }]);
  assert(nt.rows.length === 1 && nt.rows[0].a === 1, 'trim+cast+dedupe');

  const L = { columns: ['id', 'v'], rows: [{ id: '1', v: 'a' }, { id: '2', v: 'b' }] };
  const R = { columns: ['id', 'w'], rows: [{ id: '2', w: 'x' }, { id: '3', w: 'y' }] };
  assert(E.joinTables(L, R, { leftKey: 'id', rightKey: 'id', type: 'inner' }).rows.length === 1, 'inner join');
  assert(E.joinTables(L, R, { leftKey: 'id', rightKey: 'id', type: 'full' }).rows.length === 3, 'full join');

  const T = { employees: { columns: ['name', 'dept', 'salary'], rows: [
    { name: 'a', dept: 'Eng', salary: 100 }, { name: 'b', dept: 'Eng', salary: 200 }, { name: 'c', dept: 'Sales', salary: 50 }] } };
  const r1 = E.runSQL('SELECT * FROM employees WHERE salary > 60 ORDER BY salary DESC LIMIT 2', T).result;
  assert(r1.rows.length === 2 && r1.rows[0].name === 'b', 'sql where/order/limit');
  const r2 = E.runSQL('SELECT dept, COUNT(*) AS n FROM employees GROUP BY dept ORDER BY n DESC', T).result;
  assert(r2.rows.length === 2 && r2.rows[0].n === 2, 'sql groupby');

  const items = E.prepareBar(T.employees.rows, 'dept', 'salary', 'sum', 10);
  assert(E.buildChartSVG({ type: 'bar', items, title: 't' }).startsWith('<svg'), 'svg bar');

  // encodings: utf-8 BOM, latin1 fallback, utf-16le BOM
  assert(E.decodeBytes(new Uint8Array([0xEF, 0xBB, 0xBF, 104, 105])) === 'hi', 'decode utf8 bom');
  assert(E.decodeBytes(new Uint8Array([99, 97, 102, 0xE9])) === 'café', 'decode latin1');
  assert(E.decodeBytes(new Uint8Array([0xFF, 0xFE, 104, 0, 105, 0])) === 'hi', 'decode utf16le');

  // xml entities: numeric + named; CDATA stays literal
  const xe = E.xmlToTable('<r><i a="x&amp;y">A&#65; &euro;<c><![CDATA[a &amp; b]]></c></i></r>', { recordTag: 'i' });
  assert(xe.rows[0].a === 'x&y' && xe.rows[0].c === 'a &amp; b' && xe.rows[0]._text === 'AA €', 'xml entities ' + JSON.stringify(xe.rows[0]));

  // SheetJS (vendored): real xlsx + legacy .xls round-trips
  let sheetjs = null;
  try { sheetjs = require('../vendor/xlsx/xlsx.full.min.js'); } catch (e) { console.log('skip: SheetJS (run `npm run vendor` first): ' + e.message); }
  if (sheetjs) {
    const grid = [['d', 'n'], [new Date(Date.UTC(2024, 2, 5)), 42], ['x', true]];
    const wb = sheetjs.utils.book_new();
    sheetjs.utils.book_append_sheet(wb, sheetjs.utils.aoa_to_sheet(grid), 'S');
    for (const bookType of ['xlsx', 'biff8']) {
      const buf = sheetjs.write(wb, { bookType, type: 'array' });
      const back = sheetjs.utils.sheet_to_json(sheetjs.read(buf, { type: 'array', cellDates: true }).Sheets.S, { header: 1, defval: '' });
      assert(back.length === 3 && back[1][1] === 42, 'sheetjs ' + bookType + ' roundtrip');
    }
  }

  const bytes = E.buildXlsx(['a'], [{ a: 'hi' }], 'S1');
  const sheets = await E.importXlsx(bytes.buffer.slice(0), { header: true });
  assert(sheets[0].rows.length === 1 && String(sheets[0].rows[0].a) === 'hi', 'xlsx roundtrip');

  console.log(failures ? `${failures} FAILURES` : 'ALL ENGINE TESTS DONE');
  process.exit(failures ? 1 : 0);
})().catch((e) => { console.error('FATAL', e); process.exit(1); });
