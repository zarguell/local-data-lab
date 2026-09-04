/* Vendors DuckDB-WASM + Apache Arrow ESM into vendor/duckdb/ from node_modules.
 * Source of truth is package.json + package-lock.json, so Mend Renovate can
 * discover and bump these deps natively (npm manager). Run: npm run vendor
 * Deploy (.github/workflows/pages.yml) re-runs this from the lockfile, so
 * committed vendor/ output is never required to be in sync by hand.
 *
 * Vendored (browser MVP build only — works on GitHub Pages without
 * COOP/COEP headers, unlike the coi/eh threaded builds):
 *   duckdb-browser.mjs, duckdb-browser-mvp.worker.js, duckdb-mvp.wasm,
 *   arrow/*.mjs (apache-arrow ESM tree for the "apache-arrow" specifier),
 *   flatbuffers/*.mjs, tslib.mjs (their transitive bare imports),
 *   manifest.json {versions, files, importmap}
 * Also asserts every bare import in vendor/ is covered by the importmap,
 * so a future Renovate bump that adds a dependency fails here, not in users'
 * browsers.
 */
import { cp, mkdir, readFile, writeFile, stat, readdir, rm } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const outDir = path.join(root, 'vendor', 'duckdb');
const nm = (p) => path.join(root, 'node_modules', p);

const duckdbPkg = JSON.parse(await readFile(nm('@duckdb/duckdb-wasm/package.json'), 'utf8'));
const arrowPkg = JSON.parse(await readFile(nm('apache-arrow/package.json'), 'utf8'));
const flatPkg = JSON.parse(await readFile(nm('flatbuffers/package.json'), 'utf8'));
const tslibPkg = JSON.parse(await readFile(nm('tslib/package.json'), 'utf8'));
const duckdbRoot = nm('@duckdb/duckdb-wasm');
const arrowRoot = nm('apache-arrow');

const FILES = [
  [path.join(duckdbRoot, 'dist/duckdb-browser.mjs'), 'duckdb-browser.mjs'],
  [path.join(duckdbRoot, 'dist/duckdb-browser-mvp.worker.js'), 'duckdb-browser-mvp.worker.js'],
  [path.join(duckdbRoot, 'dist/duckdb-mvp.wasm'), 'duckdb-mvp.wasm'],
];

async function copyMjsTree(srcDir, destDir, excludeDirs, exts) {
  exts = exts || ['.mjs'];
  // apache-arrow's Arrow.mjs is a facade with ~140 relative sibling imports;
  // vendor the whole .mjs tree so they resolve same-origin.
  let count = 0, bytes = 0;
  async function walk(src, dest) {
    await mkdir(dest, { recursive: true });
    for (const e of await readdir(src, { withFileTypes: true })) {
      if (e.name === 'node_modules') continue;
      if (e.isDirectory()) {
        if ((excludeDirs || []).includes(e.name)) continue;
        await walk(path.join(src, e.name), path.join(dest, e.name));
      } else if (e.isFile() && exts.some((x) => e.name.endsWith(x)) && !e.name.endsWith('.d.ts')) {
        await cp(path.join(src, e.name), path.join(dest, e.name));
        const s = await stat(path.join(dest, e.name));
        bytes += s.size; count++;
      }
    }
  }
  await walk(srcDir, destDir);
  return { count, bytes };
}

async function checkBareImports(dir, allowed) {
  const bare = new Set();
  const re = /(?:from|import)\s*["']([^"'./][^"']*)["']/g;
  async function walk(d) {
    for (const e of await readdir(d, { withFileTypes: true })) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) await walk(p);
      else if (/\.(mjs|js)$/.test(e.name) && !e.name.endsWith('.d.ts')) {
        const src = (await readFile(p, 'utf8')).split('\n').filter((l) => !/^\s*\*/.test(l)).join('\n');
        let m;
        while ((m = re.exec(src))) bare.add(m[1]);
      }
    }
  }
  await walk(dir);
  const unmapped = [...bare].filter((b) => !allowed.has(b));
  if (unmapped.length) throw new Error('unmapped bare imports: ' + unmapped.join(', ') + ' — vendor them + extend the importmap');
  console.log('bare imports ok: ' + [...bare].sort().join(', '));
}

async function main() {
  await rm(outDir, { recursive: true, force: true }); // clean slate: no stale files across version bumps
  await mkdir(outDir, { recursive: true });
  const manifest = {
    generatedBy: 'npm run vendor',
    duckdbWasm: duckdbPkg.version,
    apacheArrow: arrowPkg.version,
    flatbuffers: flatPkg.version,
    tslib: tslibPkg.version,
    files: {},
  };
  for (const [src, dest] of FILES) {
    await cp(src, path.join(outDir, dest));
    const s = await stat(path.join(outDir, dest));
    manifest.files[dest] = s.size;
    console.log(`vendored ${dest} (${(s.size / 1048576).toFixed(1)} MiB)`);
  }
  const tree = await copyMjsTree(arrowRoot, path.join(outDir, 'arrow'), ['bin']);
  // prune node-only entry points (import 'node:*' — never loaded in browsers,
  // but keep them out so a stray reference fails at build scan, not runtime)
  for (const f of ['Arrow.node.mjs', 'Arrow.node.mjs.map']) {
    await rm(path.join(outDir, 'arrow', f), { force: true });
  }
  await rm(path.join(outDir, 'arrow', 'io', 'node'), { recursive: true, force: true });
  manifest.files['arrow/*.mjs'] = `${tree.count} files, ${tree.bytes} bytes`;
  console.log(`vendored arrow/*.mjs (${tree.count} files, ${(tree.bytes / 1048576).toFixed(1)} MiB)`);
  const fb = await copyMjsTree(nm('flatbuffers/mjs'), path.join(outDir, 'flatbuffers'), [], ['.js']);
  console.log(`vendored flatbuffers ESM (${fb.count} files)`);
  await cp(nm('tslib/tslib.es6.mjs'), path.join(outDir, 'tslib.mjs'));
  console.log('vendored tslib.mjs');
  // SheetJS: single self-contained classic script (no imports), loaded on demand
  // only when a spreadsheet is imported — keeps first paint light.
  await mkdir(path.join(root, 'vendor', 'xlsx'), { recursive: true });
  await cp(nm('xlsx/dist/xlsx.full.min.js'), path.join(root, 'vendor', 'xlsx', 'xlsx.full.min.js'));
  const xs = await stat(path.join(root, 'vendor', 'xlsx', 'xlsx.full.min.js'));
  manifest.files['../xlsx/xlsx.full.min.js'] = xs.size;
  manifest.sheetjs = JSON.parse(await readFile(nm('xlsx/package.json'), 'utf8')).version;
  console.log(`vendored xlsx.full.min.js (${(xs.size / 1048576).toFixed(1)} MiB)`);
  manifest.importmap = {
    'apache-arrow': './vendor/duckdb/arrow/Arrow.mjs',
    flatbuffers: './vendor/duckdb/flatbuffers/flatbuffers.js',
    tslib: './vendor/duckdb/tslib.mjs',
  };
  // every bare specifier in the vendored tree must be served by the importmap;
  // anything else is a future Renovate surprise — fail loudly instead.
  await checkBareImports(path.join(root, 'vendor'), new Set(Object.keys(manifest.importmap)));
  // sanity: browser bundle must reference the bare specifier our importmap provides
  const mjs = await readFile(path.join(outDir, 'duckdb-browser.mjs'), 'utf8');
  if (!mjs.includes('apache-arrow')) throw new Error('duckdb-browser.mjs no longer imports apache-arrow — update importmap wiring');
  await writeFile(path.join(outDir, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
  console.log(`manifest: duckdb-wasm@${manifest.duckdbWasm} + arrow@${manifest.apacheArrow} + flatbuffers@${manifest.flatbuffers} + tslib@${manifest.tslib}`);
}

main().catch((e) => { console.error('vendor failed:', e.message); process.exit(1); });
