import argparse
import logging
import os
from typing import List, Tuple

import pandas as pd

from . import __version__
from .config import ExportConfig
from .access_reader import AccessReader
from .exporter import write_excel, write_csvs


def parse_args() -> ExportConfig:
    p = argparse.ArgumentParser(description="Convert MS Access (.mdb/.accdb) to Excel (.xlsx) or CSV")
    p.add_argument("--input", dest="input_path", help="Path to Access database (.mdb/.accdb)")
    p.add_argument("--dsn", dest="dsn", help="ODBC DSN name (alternative to --input)")
    p.add_argument("--output", dest="output_path", help="Excel output .xlsx path")
    p.add_argument("--csv-dir", dest="csv_dir", help="Directory to write CSV files (one per table/query)")
    p.add_argument("--tables", help="Comma-separated table names to export")
    p.add_argument("--queries", help="Comma-separated saved query names to export")
    p.add_argument("--sheet-prefix", default="", help="Prefix to add to sheet names")
    p.add_argument("--include", help="Comma-separated names to include (filters discovered items)")
    p.add_argument("--exclude", help="Comma-separated names to exclude")
    p.add_argument("--verbose", action="store_true", help="Enable debug logging")
    p.add_argument("--analyze", action="store_true", help="Analyze the Access file and print schema & features")
    p.add_argument("--version", action="store_true", help="Print version and exit")
    args = p.parse_args()

    if args.version:
        print(__version__)
        raise SystemExit(0)

    cfg = ExportConfig(
        input_path=args.input_path,
        dsn=args.dsn,
        output_path=args.output_path,
        csv_dir=args.csv_dir,
        tables=[s.strip() for s in args.tables.split(",") if args.tables] if args.tables else [],
        queries=[s.strip() for s in args.queries.split(",") if args.queries] if args.queries else [],
        sheet_prefix=args.sheet_prefix or "",
        include=[s.strip() for s in args.include.split(",") if args.include] if args.include else [],
        exclude=[s.strip() for s in args.exclude.split(",") if args.exclude] if args.exclude else [],
        verbose=args.verbose,
        analyze_only=args.analyze,
    )
    return cfg


def setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s %(name)s: %(message)s")


def analyze(reader: AccessReader) -> None:
    reader.connect()
    tables = reader.list_tables()
    queries = reader.list_queries()
    schema = reader.get_schema_overview()

    print("=== Access Database Analysis ===")
    print(f"Tables ({len(tables)}):")
    for t in tables:
        print(f" - {t} ({len(schema.get(t, {}))} columns)")
    print(f"Saved Queries ({len(queries)}):")
    for q in queries:
        print(f" - {q}")
    # Feature detection: attachments, memo/long text, OLE/BINARY, datetime, numeric
    feature_counts = {
        "ATTACHMENT": 0,
        "LONGCHAR": 0,  # memo
        "OLEOBJECT": 0,
        "BINARY": 0,
        "DATETIME": 0,
        "DECIMAL": 0,
        "DOUBLE": 0,
        "INTEGER": 0,
    }
    for t, cols in schema.items():
        for _, type_name in cols.items():
            tn = type_name.upper()
            for k in list(feature_counts.keys()):
                if k in tn:
                    feature_counts[k] += 1
    print("Column Type Signals:")
    for k, v in feature_counts.items():
        if v:
            print(f" - {k}: {v}")
    reader.close()


def main() -> None:
    cfg = parse_args()
    setup_logging(cfg.verbose)

    reader = AccessReader(db_path=cfg.input_path, dsn=cfg.dsn)

    if cfg.analyze_only:
        analyze(reader)
        return

    reader.connect()
    items: List[Tuple[str, pd.DataFrame]] = []

    targets = []
    if cfg.tables:
        targets.extend([("table", t) for t in cfg.tables])
    if cfg.queries:
        targets.extend([("query", q) for q in cfg.queries])

    if not targets:
        # Discover all if none specified
        for t in reader.list_tables():
            targets.append(("table", t))
        for q in reader.list_queries():
            targets.append(("query", q))

    # Apply include/exclude filters
    if cfg.include:
        targets = [x for x in targets if x[1] in cfg.include]
    if cfg.exclude:
        targets = [x for x in targets if x[1] not in cfg.exclude]

    for kind, name in targets:
        try:
            df = reader.read_table(name) if kind == "table" else reader.read_query(name)
            items.append((cfg.sheet_prefix + name, df))
        except Exception as e:
            logging.warning("Skipping %s '%s': %s", kind, name, e)

    # Export
    if cfg.output_path:
        write_excel(cfg.output_path, items)
        logging.info("Wrote Excel: %s", cfg.output_path)
    if cfg.csv_dir:
        write_csvs(cfg.csv_dir, items)
        logging.info("Wrote CSVs: %s", cfg.csv_dir)

    if not cfg.output_path and not cfg.csv_dir:
        logging.error("No output specified. Use --output .xlsx and/or --csv-dir path.")
        raise SystemExit(2)

    reader.close()

if __name__ == "__main__":
    main()
