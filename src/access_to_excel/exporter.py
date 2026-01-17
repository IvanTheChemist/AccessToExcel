import logging
import os
from typing import Dict, Iterable, List, Tuple

import pandas as pd

logger = logging.getLogger(__name__)

ILLEGAL_CHARS = set('[]:*?/\\')
MAX_SHEET_LEN = 31


def sanitize_sheet_name(name: str) -> str:
    n = ''.join('_' if ch in ILLEGAL_CHARS else ch for ch in name)
    n = n.strip()
    if len(n) == 0:
        n = "Sheet"
    if len(n) > MAX_SHEET_LEN:
        n = n[:MAX_SHEET_LEN]
    return n


def dedupe_names(names: Iterable[str]) -> List[str]:
    seen: Dict[str, int] = {}
    result: List[str] = []
    for name in names:
        base = sanitize_sheet_name(name)
        if base not in seen:
            seen[base] = 1
            result.append(base)
        else:
            i = seen[base]
            while True:
                candidate = base
                suffix = f"_{i}"
                if len(candidate) + len(suffix) > MAX_SHEET_LEN:
                    candidate = candidate[: MAX_SHEET_LEN - len(suffix)]
                candidate = candidate + suffix
                if candidate not in seen:
                    seen[base] = i + 1
                    seen[candidate] = 1
                    result.append(candidate)
                    break
                i += 1
    return result


def write_excel(output_path: str, sheets: List[Tuple[str, pd.DataFrame]], engine: str = "openpyxl") -> None:
    sheet_names = dedupe_names([name for name, _ in sheets])
    with pd.ExcelWriter(output_path, engine=engine) as writer:
        for (orig_name, df), sheet_name in zip(sheets, sheet_names):
            logger.debug("Writing sheet '%s' from '%s' rows=%d", sheet_name, orig_name, len(df))
            # Ensure datetimes are timezone-naive for Excel
            for col in df.select_dtypes(include=['datetimes']):
                df[col] = df[col].dt.tz_localize(None)
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def write_csvs(dir_path: str, files: List[Tuple[str, pd.DataFrame]]) -> None:
    os.makedirs(dir_path, exist_ok=True)
    names = dedupe_names([name for name, _ in files])
    for (orig_name, df), base in zip(files, names):
        filename = f"{base}.csv"
        out = os.path.join(dir_path, filename)
        logger.debug("Writing CSV '%s' from '%s' rows=%d", filename, orig_name, len(df))
        df.to_csv(out, index=False)
