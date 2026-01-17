from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class ExportConfig:
    input_path: Optional[str] = None
    dsn: Optional[str] = None
    output_path: Optional[str] = None  # .xlsx
    csv_dir: Optional[str] = None      # directory for CSVs
    tables: List[str] = field(default_factory=list)
    queries: List[str] = field(default_factory=list)
    sheet_prefix: str = ""
    include: List[str] = field(default_factory=list)
    exclude: List[str] = field(default_factory=list)
    verbose: bool = False
    analyze_only: bool = False
    chunk_size: Optional[int] = None  # future: for large tables
