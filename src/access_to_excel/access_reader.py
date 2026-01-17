import logging
from typing import List, Tuple, Dict, Optional

import pyodbc
import pandas as pd

logger = logging.getLogger(__name__)

ACCESS_DRIVER = "{Microsoft Access Driver (*.mdb, *.accdb)}"

class AccessReader:
    def __init__(self, db_path: Optional[str] = None, dsn: Optional[str] = None) -> None:
        self.db_path = db_path
        self.dsn = dsn
        self.conn: Optional[pyodbc.Connection] = None

    def connect(self) -> None:
        try:
            if self.dsn:
                conn_str = f"DSN={self.dsn}"
            elif self.db_path:
                conn_str = f"DRIVER={ACCESS_DRIVER};DBQ={self.db_path}"
            else:
                raise ValueError("Either db_path or dsn must be provided")
            logger.debug("Connecting with: %s", conn_str)
            self.conn = pyodbc.connect(conn_str)
        except pyodbc.Error as e:
            msg = (
                "Failed to connect to Access database. Ensure the Microsoft Access "
                "Database Engine ODBC driver is installed and bitness (32/64-bit) "
                "matches your Python interpreter. Original error: %s" % e
            )
            logger.error(msg)
            raise

    def close(self) -> None:
        if self.conn:
            self.conn.close()
            self.conn = None

    def list_tables(self) -> List[str]:
        assert self.conn is not None, "Not connected"
        # MSysObjects: Type IN (1,4,6) and Flags=0 returns user tables.
        sql = "SELECT Name FROM MSysObjects WHERE Type IN (1,4,6) AND Flags=0;"
        try:
            names = [row[0] for row in self.conn.cursor().execute(sql).fetchall()]
            logger.debug("Found tables: %s", names)
            return names
        except pyodbc.Error:
            # Fallback to ODBC table enumeration
            logger.debug("MSysObjects not accessible; falling back to cursor.tables")
            cur = self.conn.cursor()
            names = [row.table_name for row in cur.tables(tableType='TABLE')]
            return names

    def list_queries(self) -> List[str]:
        assert self.conn is not None, "Not connected"
        sql = "SELECT Name FROM MSysObjects WHERE Type=5 AND Flags=0;"
        try:
            names = [row[0] for row in self.conn.cursor().execute(sql).fetchall()]
            logger.debug("Found queries: %s", names)
            return names
        except pyodbc.Error:
            logger.debug("MSysObjects queries not accessible")
            return []

    def get_schema_overview(self) -> Dict[str, Dict[str, str]]:
        """Return {table_name: {column_name: type_name}} summary."""
        assert self.conn is not None, "Not connected"
        cur = self.conn.cursor()
        overview: Dict[str, Dict[str, str]] = {}
        for t in self.list_tables():
            cols: Dict[str, str] = {}
            try:
                for c in cur.columns(table=t):
                    # c.type_name may include 'LONGCHAR', 'VARCHAR', 'DATETIME', 'DECIMAL', 'BINARY', 'IMAGE', etc.
                    cols[str(c.column_name)] = str(getattr(c, 'type_name', getattr(c, 'data_type', 'UNKNOWN')))
            except pyodbc.Error:
                cols = {}
            overview[t] = cols
        return overview

    def read_table(self, name: str) -> pd.DataFrame:
        assert self.conn is not None, "Not connected"
        return pd.read_sql(f"SELECT * FROM [{name}]", self.conn)

    def read_query(self, name: str) -> pd.DataFrame:
        assert self.conn is not None, "Not connected"
        # Access allows selecting from saved queries by name.
        try:
            return pd.read_sql(f"SELECT * FROM [{name}]", self.conn)
        except Exception:
            # Fallback: execute query definition via MSysObjects? Non-trivial; skip for now.
            raise RuntimeError(f"Unable to read saved query '{name}'. Ensure it returns a rowset.")
