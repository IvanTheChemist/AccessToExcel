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
    p.add_argument("--show-query", dest="show_query", help="Show saved query definition and sample rows")
    p.add_argument("--limit", type=int, default=15, help="Sample row limit for --show-query (default: 15)")
    p.add_argument("--param", action="append", help="Parameter override for saved queries in name=value form; repeatable")
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


def show_query(conn: pyodbc.Connection, db_path: Optional[str], name: str, limit: int, params: Optional[List[str]]) -> None:
    print(f"=== Saved Query: {name} ===")
    # Try to display SQL definition via DAO
    sql_text: Optional[str] = None
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
                        if str(q.Name).lower() == name.lower():
                            sql_text = str(q.SQL)
                            # Try to print parameters metadata
                            try:
                                params_info = []
                                for p in q.Parameters:
                                    pname = str(p.Name)
                                    # Type codes are DAO constants; we print raw value
                                    ptype = getattr(p, 'Type', None)
                                    params_info.append((pname, ptype))
                                if params_info:
                                    print("Parameters:")
                                    for pname, ptype in params_info:
                                        print(f" - {pname} (Type={ptype})")
                            except Exception:
                                pass
                            break
                finally:
                    try:
                        db.Close()
                    except Exception:
                        pass
        except Exception:
            pass
    if sql_text:
        print("Definition:")
        print(sql_text)
    else:
        print("Definition: (not available) Install 'pywin32' to enable DAO QueryDefs reading.")

    # Try DAO execution (supports Access functions and parameterized queries)
    executed = False
    if db_path:
        try:
            import re
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
                    qdef = None
                    for q in db.QueryDefs:
                        if str(q.Name).lower() == name.lower():
                            qdef = q
                            break
                    if qdef is not None:
                        # Apply parameters if provided
                        provided: Dict[str, str] = {}
                        if params:
                            for p in params:
                                if "=" in p:
                                    k, v = p.split("=", 1)
                                    provided[k.strip().lower()] = v
                        # Normalize and set values
                        for p in qdef.Parameters:
                            pname = str(p.Name)
                            keys_to_try = [pname.lower()]
                            # Try bracket content
                            m = re.search(r"\[(.*?)\]", pname)
                            if m:
                                keys_to_try.append(m.group(1).lower())
                            # Try segment after dot
                            if "." in pname:
                                keys_to_try.append(pname.split(".")[-1].lower())
                            val = None
                            for k in keys_to_try:
                                if k in provided:
                                    val = provided[k]
                                    break
                            if val is not None:
                                try:
                                    p.Value = val
                                except Exception:
                                    pass
                        rs = qdef.OpenRecordset()
                        try:
                            # Print header
                            cols = [str(f.Name) for f in rs.Fields]
                            print(f"Sample Rows (limit={limit}) via DAO:")
                            print(",".join(cols))
                            count = 0
                            while count < limit and (not rs.EOF):
                                values = []
                                for f in rs.Fields:
                                    try:
                                        v = f.Value
                                    except Exception:
                                        v = None
                                    values.append("" if v is None else str(v))
                                print(",".join(values))
                                rs.MoveNext()
                                count += 1
                            executed = True
                        finally:
                            try:
                                rs.Close()
                            except Exception:
                                pass
                finally:
                    try:
                        db.Close()
                    except Exception:
                        pass
        except Exception:
            pass

    if not executed:
        # Fallback to ODBC execution for non-parameterized queries
        cur = conn.cursor()
        try:
            rows = cur.execute(f"SELECT * FROM [{name}]").fetchmany(limit)
            if rows:
                print(f"Sample Rows (limit={limit}):")
                cols = [c[0] for c in cur.description]
                print(",".join(cols))
                for r in rows:
                    print(",".join(str(x) if x is not None else "" for x in r))
            else:
                print("No rows returned.")
        except Exception as e:
            print(f"Execution failed: {e}")


def main() -> None:
    args = parse_args()
    if args.version:
        print(__version__)
        return
    setup_logging(args.verbose)
    if not args.analyze:
        if args.show_query:
            conn = _connect(args.input_path, args.dsn)
            try:
                show_query(conn, args.input_path, args.show_query, args.limit, args.param)
            finally:
                conn.close()
        else:
            print("Use --analyze or --show-query <name>.")
        return
    conn = _connect(args.input_path, args.dsn)
    try:
        analyze(conn, args.input_path)
    finally:
        conn.close()

if __name__ == "__main__":
    main()
