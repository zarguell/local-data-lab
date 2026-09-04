# Local Data Lab

Offline web data platform. Static site, fully self-contained, no CDN, no server. Open `index.html` directly or host anywhere.

## Pipeline

**Upload → Parse → Normalize → Join → SQL → Visualize → Export** — everything runs in your browser; files never leave your machine.

## Formats in

- **CSV / TSV / TXT** — auto delimiter detection (`,` `⇥` `;` `|`), quoted fields, embedded newlines, header toggle
- **JSON / NDJSON** — objects, arrays, primitives, JSON Lines; root-path picker; nested flattening + array explode
- **XML** — auto-detects the repeating record tag (or set it manually), attributes + nested elements flattened
- **XLS/XLSX** — primary importer is vendored SheetJS (`xlsx`), so real-world files work: shared strings, styles/dates, multi-sheet, **legacy `.xls` (BIFF) and `.xlsb`** — formats the old hand parser never could. Loaded on demand only. If `vendor/` is absent (fresh clone over `file://`), xlsx-family gracefully falls back to the built-in dependency-free parser.
- **ZIP** — unpacks bundles of the above, one table per inner file
- **Encodings** — text files decode as UTF-8 strict → UTF-16 via BOM → windows-1252 fallback, so Excel-exported latin1 CSVs stop mojibaking.
- Paste box (auto-detect) + one-click samples, incl. a generated-XLSX round-trip proof

## Normalize

Per-table schema with inferred types (`integer · number · boolean · date · string`), null/distinct stats, rename, cast, hide, trim, drop empty rows/cols, dedupe.

## Join

Inner / left / right / full-outer hash joins on any two tables + key columns, with automatic `*_right` disambiguation → new table.

## SQL — two engines

- **Built-in mini-SQL** (default): instant, zero-load, works even from `file://`. `SELECT [DISTINCT]`, `*`, `t.*`, aliases, `WHERE` (`AND/OR/NOT`, comparisons, `LIKE`, `IN`, `IS [NOT] NULL`), all four joins + cross, `GROUP BY` + `COUNT/SUM/AVG/MIN/MAX`, `ORDER BY`, `LIMIT/OFFSET`.
- **DuckDB 1.32 (vendored WASM)**: full SQL — subqueries, window functions, `UNION`, `PIVOT`, rich types/functions. Pick it in the §4 engine dropdown; the ~38MB WASM loads once from the *same origin* (no CDN) and is cached afterwards. Needs `http(s)` — from `file://` you'll get a clear fallback message instead.

```sql
SELECT dept, COUNT(*) AS n, AVG(salary) AS avg_salary
FROM employees WHERE active = true
GROUP BY dept ORDER BY n DESC LIMIT 10;

SELECT a.name, b.budget FROM employees a
JOIN departments b ON a.dept = b.dept;
```

Supports `SELECT [DISTINCT]`, `*`, `t.*`, aliases, `WHERE` (`AND/OR/NOT`, comparisons, `LIKE`, `IN`, `IS [NOT] NULL`), all four joins + cross, `GROUP BY` + `COUNT/SUM/AVG/MIN/MAX`, `ORDER BY`, `LIMIT/OFFSET`. Save any result as a table (chainable) or export it. Query result is chartable as `$query`.

## Visualize

Hand-rolled SVG (no chart lib): **bar** (group + sum/count/avg/min/max), **line**, **scatter**, **pie** — from any table or `$query`. Export **SVG** or **PNG**.

## Export

Per table: CSV · TSV · JSON · XLSX (frozen header + auto-filter + widths). Query results: CSV/XLSX. Everything: one `.zip` of CSVs.

## Dependencies & Renovate

The runtime deps are `@duckdb/duckdb-wasm` + `apache-arrow` (+ its `flatbuffers`/`tslib` transitive imports) and `xlsx` (SheetJS), managed as **exact pins in `package.json` + `package-lock.json`** — Mend Renovate's npm manager discovers them with zero extra config, and its PRs update the lockfile. Nothing else to do:

- `npm run vendor` copies the browser MVP build + Arrow ESM tree + SheetJS classic script from `node_modules` into `vendor/` (+ `manifest.json`), and **fails if any bare import isn't covered by the importmap** — a future Renovate bump that adds a dependency breaks the build, not users' browsers.
- `.github/workflows/pages.yml` re-runs that from the lockfile on every deploy, so the 38MB `.wasm` never lives in git and Renovate PRs flow straight to production after merge.
- Only the MVP (single-threaded) build is vendored: unlike the `coi`/`eh` builds it needs no COOP/COEP headers, so it runs on GitHub Pages.

```bash
npm ci && npm run vendor   # local dev: enables the DuckDB engine
# vendor/duckdb/manifest.json records the vendored versions
python3 -m http.server 8000  # → http://localhost:8000 (http needed for DuckDB; built-in engine also works from file://)
```

## Deliberately still hand-rolled

Crossing the npm Rubicon doesn't mean depending on everything. Kept because they're small, tested, and a dependency buys nothing:

- **Built-in mini-SQL** — default engine: instant, zero-load, works from `file://`. DuckDB is the upgrade path, not the replacement.
- **CSV/TSV parser, XML parser, ZIP reader/writer, XLSX writer, SVG charts** — each is a few hundred lines with committed tests (`npm test`) covering the real paths (incl. deflated zips and generated-XLSX round-trips). Swapping any of them would add supply-chain surface for no user-visible gain.

## Run / deploy

Push to `main` — the Pages workflow (`pages.yml`: install → test → vendor → deploy) publishes it.
