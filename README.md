# AccessToExcel

A Windows-first command-line tool to analyze and export MS Access (`.mdb`/`.accdb`) database data and saved queries to CSV/Excel format.

## Features

- **Database Analysis**: Inspect Access database schema, tables, and saved queries
- **Query Viewer**: Display saved query definitions with sample data
- **Parameterized Query Support**: Execute queries with parameters via DAO (requires `pywin32`)
- **Flexible Driver Detection**: Automatically selects the appropriate ODBC driver for your database
- **DSN-less Connection**: Connect directly to `.mdb`/`.accdb` files without configuring ODBC data sources

## Requirements

### System Requirements
- **Windows** operating system
- **64-bit Python** 3.7 or later
- **64-bit MS Access ODBC Driver** (Access Database Engine)
  - Verify installation: Run `python -c "import pyodbc; print(pyodbc.drivers())"` and look for drivers containing "Access"
  - Download from Microsoft if needed: [Microsoft Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=54920)

### Python Dependencies
- `pyodbc` (required)
- `pywin32` (optional, for DAO features like QueryDefs SQL/parameters and Access-specific functions)

## Installation

1. **Clone the repository:**
   ```powershell
   git clone https://github.com/IvanTheChemist/AccessToExcel.git
   cd AccessToExcel
   ```

2. **Create and activate a virtual environment:**
   ```powershell
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   ```

3. **Install dependencies:**
   ```powershell
   pip install pyodbc
   # Optional: For full DAO support (parameters, Access functions)
   pip install pywin32
   ```

## Usage

### Analyze Database Schema

Inspect tables, views, and saved queries in your Access database:

```powershell
python -m access_to_excel --input "C:\path\to\database.accdb" --analyze
```

Add `--verbose` for detailed debug logging:

```powershell
python -m access_to_excel --input "C:\path\to\database.accdb" --analyze --verbose
```

### Show Saved Query

Display a saved query's definition and sample rows:

```powershell
python -m access_to_excel --input "C:\path\to\database.accdb" --show-query "QueryName"
```

Customize the number of sample rows:

```powershell
python -m access_to_excel --input "C:\path\to\database.accdb" --show-query "QueryName" --limit 25
```

### Execute Parameterized Queries

For queries with parameters, provide values using `--param`:

```powershell
python -m access_to_excel --input "C:\path\to\database.accdb" --show-query "SalesByRegion" --param "Region=West" --param "Year=2024"
```

Parameter name matching is flexible:
- Full name: `RegionParameter`
- Inside brackets: `[RegionParameter]`
- Segment after period: `Forms.RegionParameter`

### Using a Named DSN

Alternatively, connect via an ODBC DSN:

```powershell
python -m access_to_excel --dsn "MyAccessDSN" --analyze
```

## Architecture Overview

- **`access_to_excel/__main__.py`**: Module entry point
- **`access_to_excel/cli.py`**: Main analyzer and query viewer logic
  - Driver selection via `pyodbc.drivers()`
  - DSN-less ODBC connection
  - Table discovery via ODBC
  - Saved query detection (ODBC views + optional DAO QueryDefs)
- **`access_to_excel/__init__.py`**: Version information

### Connection Strategy

1. **ODBC**: Primary method for connecting and listing tables/views
   - Connection string: `DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=<path>`
   - Lists tables via `cursor.tables(tableType='TABLE')`
   - Lists saved queries as ODBC `VIEW` entries

2. **DAO (via pywin32)**: Optional fallback for advanced features
   - Reads QueryDefs names, SQL definitions, and parameters
   - Executes parameterized queries
   - Supports Access-specific functions (e.g., `Nz`, `IIf`)

## Project Structure

```
AccessToExcel/
├── access_to_excel/         # Main package
│   ├── __init__.py          # Version info
│   ├── __main__.py          # Entry point
│   └── cli.py               # CLI logic
├── docs/
│   └── examples/            # Synthetic sample data (whitelisted)
│       ├── README.md
│       └── sample_customers.csv
├── tests/                   # Test files
├── .github/
│   └── copilot-instructions.md
├── .gitignore               # Ignores *.xlsx, *.csv, out/, exports/
└── README.md
```

## Development Conventions

- **No Test Data in Code**: No hard-coded test data strings in source code
- **Synthetic Samples**: Example files live in `docs/examples/` only
- **Ignored Outputs**: Data exports (`*.xlsx`, `*.csv`, `out/`, `exports/`) are gitignored by default
- **Logging**: Standard Python `logging`; use `--verbose` for DEBUG level

## Roadmap

Potential future features:
- CSV/Excel export for tables and queries (`--export-csv`, `--export-xlsx`)
- Filtering support (`--include`, `--exclude` patterns)
- Batch export (all tables, all queries)
- Excel export with sanitized sheet names (31-char limit, no special chars)

## Contributing

Contributions are welcome! Please follow the existing code conventions and ensure:
- No test data strings in code
- Synthetic examples in `docs/examples/` only
- Export outputs are gitignored
- ODBC-first approach with DAO fallback for advanced features

## License

[Specify your license here]

## Version

Current version: **0.0.1**

## Troubleshooting

### "No Access driver found"
- Install the 64-bit Microsoft Access Database Engine
- Verify with: `python -c "import pyodbc; print(pyodbc.drivers())"`
- Ensure Python and Access driver architectures match (both 64-bit)

### "Table/Query not found"
- Use `--analyze` first to list available tables and queries
- Check spelling and capitalization (Access object names are case-sensitive in queries)

### "Module not found" or Import Errors
- Activate your virtual environment: `.venv\Scripts\Activate.ps1`
- Install dependencies: `pip install pyodbc pywin32`

### Parameterized Queries Not Working
- Ensure `pywin32` is installed: `pip install pywin32`
- Check parameter names with `--analyze` or by inspecting the query in Access
