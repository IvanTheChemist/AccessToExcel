# AI Coding Agent Instructions — AccessToExcel

Project focus: Windows-first tool to analyze and export MS Access (`.mdb/.accdb`) data and saved queries to CSV/Excel.

## Current Architecture
- `access_to_excel/__main__.py`: Module entry (`python -m access_to_excel`).
- `access_to_excel/cli.py`: Analyzer + query viewer.
  - Driver selection via `pyodbc.drivers()`; prefers `*.accdb`-capable driver.
  - Connects DSN-less: `DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=<path>`.
  - Lists tables via ODBC (`cursor.tables(tableType='TABLE')`).
  - Saved queries:
    - Primary: ODBC `VIEW` entries (often mapped from Access queries).
    - Secondary: `MSysObjects` (if accessible) — filter out `MSys*`.
    - Optional DAO fallback (requires `pywin32`) to read QueryDefs names/SQL/parameters and execute parameterized queries.

## CLI Usage
- Analyze schema/features:
  - `python -m access_to_excel --input <db.accdb> --analyze [--verbose]`
- Show saved query (definition + sample rows):
  - `python -m access_to_excel --input <db.accdb> --show-query "<QueryName>" --limit 15`
  - Parameters: `--param name=value` (repeatable). Name matching is flexible: full, inside `[]`, or segment after `.`.

## Requirements & Environment
- Windows + 64-bit Python (verified in this repo).
- 64-bit Access ODBC driver installed (Access Database Engine). Verify via `pyodbc.drivers()`.
- For DAO features (QueryDefs SQL/parameters, Access functions): install `pywin32` in the active venv.

## Conventions
- No hard-coded test data in code. Synthetic samples live in `docs/examples/` only.
- Data outputs should not be committed by default:
  - `.gitignore` ignores `*.xlsx`, `*.csv`, `out/`, `exports/`.
  - Synthetic samples are whitelisted: `docs/examples/*.csv|*.xlsx`.
- Logging: standard `logging`; `--verbose` enables DEBUG.

## Patterns to Follow
- Prefer ODBC for table discovery and non-parameterized query sampling.
- Use DAO for:
  - Reading `QueryDefs.SQL` and parameter metadata.
  - Executing queries with parameters or Access-specific functions (`Nz`, etc.).
- When adding export:
  - CSV first; Excel optional. Sanitize sheet names if Excel is added (31-char limit, no `[]:*?/\`).
  - Support `--tables`, `--views` (saved queries), `--include/--exclude` filters.

## Example Files
- `access_to_excel/cli.py`: reference for driver detection, ODBC views listing, DAO parameter handling.
- `docs/examples/sample_customers.csv`: synthetic sample; do not model real data from it.

## Notes & Recent Changes
- Analyzer CLI re-added under top-level package; supports `--analyze`, `--show-query`, `--param`, `--limit`.
- Verified driver presence and DSN-less connection; ODBC views enumerate saved queries.
- No test data strings found in code; synthetic examples retained in `docs/examples/` per policy.

If you need additional flags (e.g., `--export-csv <QueryName>`), follow the DAO-first execution path for parameterized queries and fall back to ODBC for simple views.