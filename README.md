# AccessToExcel

Convert Microsoft Access databases (`.mdb`/`.accdb`) to Excel (`.xlsx`) or CSV on Windows.

## Prerequisites
- Windows with the Microsoft Access Database Engine ODBC driver installed.
- Python 3.10+ with matching 32/64-bit to the driver.

## Install
```pwsh
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e .
```

## Usage
Analyze a database to determine features and schema:
```pwsh
python -m access_to_excel --input "g:\path\db.accdb" --analyze --verbose
```
Export everything to Excel and CSV:
```pwsh
python -m access_to_excel --input "g:\path\db.accdb" --output "g:\out.xlsx" --csv-dir "g:\out-csv" --verbose
```
Export specific tables/queries:
```pwsh
python -m access_to_excel --input "g:\path\db.accdb" --tables Customers,Orders --queries Q_ActiveOrders --output "g:\out.xlsx"
```

## Notes
- If MSysObjects is not accessible due to permissions, table discovery falls back to ODBC `cursor.tables`.
- Saved queries must return rowsets to be exported; complex parameterized queries may fail.
- Excel sheet names are sanitized and deduplicated (31-char limit).

## Tests
```pwsh
pytest -q
```
