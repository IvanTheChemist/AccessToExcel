import argparse
import logging
from typing import Dict, List, Optional

import pyodbc

from . import __version__

logger = logging.getLogger(__name__)


def _pick_access_driver(db_path: Optional[str]) -> Optional[str]:
    drivers = [d for d in pyodbc.drivers()]
    ext = (str(db_path).lower().rsplit(".", 1)[-1] if db_path else None)
    # Prefer accdb-capable
    for d in drivers:
        if "Access" in d and "accdb" in d.lower():
            return f"{{{d}}}"
    # Fallback: any Access driver (mdb-only)
    for d in drivers:
        if "Access" in d:
            if ext == "accdb" and "accdb" not in d.lower():
                continue
            return f"{{{d}}}"
    return None


def parse_args():
    p = argparse.ArgumentParser(description="Analyze MS Access (.mdb/.accdb) database schema and features")
    p.add_argument("--input", dest="input_path", help="Path to Access database (.mdb/.accdb)")
    p.add_argument("--dsn", dest="dsn", help="ODBC DSN name (alternative to --input)")
    p.add_argument("--analyze", action="store_true", help="Run analyzer and print summary")
    p.add_argument("--verbose", action="store_true", help="Enable debug logging")
    p.add_argument("--version", action="store_true", help="Print version and exit")
    args = p.parse_args()
    return args


def setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s %(name)s: %(message)s")


def _connect(input_path: Optional[str], dsn: Optional[str]) -> pyodbc.Connection:
    if dsn:
        conn_str = f"DSN={dsn}"
    elif input_path:
        driver = _pick_access_driver(input_path)
        if not driver:
            raise RuntimeError(
                "No Microsoft Access ODBC driver found. Install the 64-bit 'Microsoft Access Database Engine' and retry."
            )
        conn_str = f"DRIVER={driver};DBQ={input_path}"
    else:
        raise ValueError("Either --input or --dsn must be provided")
    logger.debug("Connecting with: %s", conn_str)
    return pyodbc.connect(conn_str)


def analyze(conn: pyodbc.Connection, db_path: Optional[str]) -> None:
    cur = conn.cursor()

    # Tables via ODBC API
    tables = [r.table_name for r in cur.tables(tableType='TABLE')]
    print(f"Tables ({len(tables)}):")
    for t in tables:
        print(f" - {t}")

    # Saved queries via MSysObjects, then DAO fallback
    queries: List[str] = []
    try:
        for row in cur.execute("SELECT Name FROM MSysObjects WHERE Type=5 AND Flags=0;"):
            name = row[0]
            if name and not str(name).startswith("MSys"):
                queries.append(name)
    except pyodbc.Error:
        logger.debug("MSysObjects not accessible; trying DAO fallback for saved queries")
        if db_path:
            try:
                import win32com.client as win32  # type: ignore
                engine = None
                for progid in ("DAO.DBEngine.120", "DAO.DBEngine.160", "DAO.DBEngine"):
                    try:
                        engine = win32.Dispatch(progid)
                        break
                    except Exception:
                        continue
                if engine is not None:
                    db = engine.OpenDatabase(db_path)
                    try:
                        for q in db.QueryDefs:
                            name = str(q.Name)
                            if name and not name.startswith("MSys") and not name.startswith("~"):
                                queries.append(name)
                    finally:
                        try:
                            db.Close()
                        except Exception:
                            pass
            except Exception:
                logger.debug("DAO fallback not available (install pywin32) or failed")
    print(f"Saved Queries ({len(queries)}):")
    for q in queries:
        print(f" - {q}")

    # Also list ODBC views which often map to Access saved queries
    try:
        views = [r.table_name for r in cur.tables(tableType='VIEW')]
    except pyodbc.Error:
        views = []
    if views:
        print(f"ODBC Views ({len(views)}):")
        for v in views:
            print(f" - {v}")

    # Column type signals
    feature_counts: Dict[str, int] = {
        "ATTACHMENT": 0,
        "LONGCHAR": 0,
        "OLEOBJECT": 0,
        "BINARY": 0,
        "DATETIME": 0,
        "DECIMAL": 0,
        "DOUBLE": 0,
        "INTEGER": 0,
    }
    for t in tables:
        try:
            for c in cur.columns(table=t):
                tn = str(getattr(c, 'type_name', getattr(c, 'data_type', ''))).upper()
                for k in list(feature_counts.keys()):
                    if k in tn:
                        feature_counts[k] += 1
        except pyodbc.Error:
            continue
    print("Column Type Signals:")
    for k, v in feature_counts.items():
        if v:
            print(f" - {k}: {v}")


def main() -> None:
    args = parse_args()
    if args.version:
        print(__version__)
        return
    setup_logging(args.verbose)
    if not args.analyze:
        print("Use --analyze to run the analyzer.")
        return
    conn = _connect(args.input_path, args.dsn)
    try:
        analyze(conn, args.input_path)
    finally:
        conn.close()

if __name__ == "__main__":
    main()
