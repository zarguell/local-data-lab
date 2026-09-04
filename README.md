# Local Data Lab

Offline web data platform. Static site, fully self-contained, no CDN, no server. Open `index.html` directly or host anywhere.

## Pipeline

**Upload → Parse → Normalize → Join → SQL → Visualize → Export** — everything runs in your browser; files never leave your machine.

## Formats in

- **CSV / TSV / TXT** — auto delimiter detection (`,` `⇥` `;` `|`), quoted fields, embedded newlines, header toggle
- **JSON / NDJSON** — objects, arrays, primitives, JSON Lines; root-path picker; nested flattening + array explode
- **XML** — auto-detects the repeating record tag (or set it manually), attributes + nested elements flattened
- **XLSX / XLSM** — real import (stored + deflated zip entries, shared strings, inline strings, multi-sheet, Excel date cells → `YYYY-MM-DD`)
- **ZIP** — unpacks bundles of the above, one table per inner file
- Paste box (auto-detect) + one-click samples, incl. a generated-XLSX round-trip proof

## Normalize

Per-table schema with inferred types (`integer · number · boolean · date · string`), null/distinct stats, rename, cast, hide, trim, drop empty rows/cols, dedupe.

## Join

Inner / left / right / full-outer hash joins on any two tables + key columns, with automatic `*_right` disambiguation → new table.

## SQL

Hand-rolled in-browser engine (no WASM, no deps):

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

## Run / deploy

```bash
python3 -m http.server 8000  # → http://localhost:8000
```

Push to `main` — GitHub Pages serves it.
