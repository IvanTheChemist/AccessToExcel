# AI Coding Agent Instructions — AccessToExcel

Project is currently empty. These instructions establish initial, concrete conventions to keep agents aligned as the repo is implemented.

## Scope and Goal
- Convert a Microsoft Access database (`.mdb`/`.accdb`) to an Excel workbook (`.xlsx`).
- Support exporting: specific tables, saved queries, or entire DB.
- Provide a Windows-first CLI; keep code Windows-compatible and avoid non-Windows-only APIs unless guarded.

## Planned Architecture (keep consistent as code lands)
- `src/access_to_excel/cli.py`: CLI entry point (`python -m access_to_excel`).
- `src/access_to_excel/access_reader.py`: Access I/O via `pyodbc` (ODBC Access Driver).
- `src/access_to_excel/exporter.py`: Dataframe-to-Excel writing using `pandas` + `openpyxl`/`xlsxwriter`.
- `src/access_to_excel/config.py`: Options: include/exclude lists, sheet naming, type conversions.
- `tests/`: Pytest tests; sample fixtures under `tests/fixtures/` with tiny `.accdb`.

## Data Flow
- Connect: `pyodbc.connect("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=<path>")`.
- Discover artifacts:
  - Tables: `SELECT Name FROM MSysObjects WHERE Type IN (1,4,6) AND Flags=0;` (guard for permissions).
  - Queries: list saved queries from `MSysObjects` where `Type=5`.
- Read: `pandas.read_sql(query_or_table_name, conn)` returns DataFrame.
- Write: `pandas.ExcelWriter(output, engine="openpyxl")` → one sheet per table/query.

## Windows/Access Requirements
- Ensure the Microsoft Access ODBC driver (Access Database Engine) is installed.
- Prefer 64-bit Python with matching 64-bit driver; mismatches cause connection failure.
- If ODBC unavailable, optionally fall back to `ADO` via `win32com` (behind a feature flag).

## Developer Workflows
- Env setup:
  - `pyproject.toml` with `pandas`, `pyodbc`, `openpyxl`. Optional: `xlsxwriter`, `pytest`.
  - Create venv and install deps.
- Run (examples):
  - `python -m access_to_excel --input g:\path\db.accdb --output g:\out.xlsx`.
  - `python -m access_to_excel --tables Customers,Orders`.
  - `python -m access_to_excel --queries Q_ActiveOrders`.
- Test: `pytest -q` with fixture DBs; mock I/O for unit tests.
- Logging: Use stdlib `logging`; default INFO; `--verbose` for DEBUG.

## Conventions and Patterns
- CLI options: `--input`, `--output`, `--tables`, `--queries`, `--sheet-prefix`, `--include/--exclude`.
- Error surfaces: return exit codes; friendly messages for driver/permission issues.
- Sheet naming: sanitize to Excel limits (<=31 chars, no `[]:*?/\\`). Deduplicate.
- Types: preserve Access field types; coerce datetimes to timezone-naive Excel-friendly values.
- Large tables: chunked reads/writes if needed; avoid loading entire DB into memory.

## Integration Points
- ODBC DSN-less connection preferred; allow `--dsn <name>` as alternative.
- External dependency: Microsoft Access Database Engine Redistributable (documented in README).

## File References (to create)
- `src/access_to_excel/`: main package.
- `tests/`: unit and integration tests.
- `README.md`: install prerequisites, usage examples, troubleshooting (driver bitness, MSysObjects visibility).

## Pull Requests
- Keep changes small and focused; add tests for new CLI options.
- Update `README.md` when adding or changing flags.

If any of the above is unclear or you need different defaults (e.g., .NET instead of Python, CSV export instead of Excel), please comment and I will adjust these instructions.