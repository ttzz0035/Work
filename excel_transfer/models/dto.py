# excel_transfer/models/dto.py
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, List, Optional, Dict, Any


LogFn = Callable[[str], None]


# -------------------------------------------------
# Transfer
# -------------------------------------------------
@dataclass
class TransferRequest:
    csv_paths: List[str] = field(default_factory=list)
    out_of_range_mode: str = "error"


# -------------------------------------------------
# Grep
# -------------------------------------------------
@dataclass
class GrepRequest:
    # --- existing (DO NOT CHANGE SEMANTICS) ---
    root_dir: str = ""
    keyword: str = ""
    ignore_case: bool = True
    use_regex: bool = False

    # --- file selector (by filename regex) ---
    file_name_regex: Optional[str] = None  # applied to basename

    # --- sheet selector (search scope) ---
    sheet_name_regex: Optional[str] = None
    sheet_indices: Optional[List[int]] = None  # 1-based indices

    # --- search offset (hit -> target) ---
    offset_row: int = 0
    offset_col: int = 0

    # --- replace enable flag ---
    # True: replace is enabled ("" means empty-string replace)
    # False: search-only (do not replace)
    replace_enabled: bool = False

    # --- replace ---
    replace_pattern: str = ""

    # --- execution mode ---
    # "preview" | "auto"
    replace_mode: str = "preview"

# -------------------------------------------------
# Count
# -------------------------------------------------
@dataclass
class CountRequest:
    files: List[str] = field(default_factory=list)
    sheet: str = ""
    start_cell: str = "B2"
    direction: str = "row"
    tolerate_blanks: int = 0
    mode: str = "jump"


# -------------------------------------------------
# Diff
# -------------------------------------------------
@dataclass
class DiffRequest:
    file_a: str = ""
    file_b: str = ""
    range_a: str = ""
    range_b: str = ""

    # 差分ベース
    base_file: str = "B"   # "A" or "B"

    # 比較オプション
    compare_formula: bool = False
    include_context: bool = True
    compare_shapes: bool = False

    # シート比較モード
    # "index": インデックス一致（既定）
    # "name" : シート名一致
    sheet_mode: str = "index"

    # 将来拡張用（明示シート指定）
    sheet_a: str = ""
    sheet_b: str = ""

# -------------------------------------------------
# Diff (Result)  ※単一DTOのみ追加
# -------------------------------------------------
@dataclass
class DiffResult:
    diff_path: str
    json_path: str

    meta: Dict[str, Any] = field(default_factory=dict)
    summary: Dict[str, Any] = field(default_factory=dict)

    diff_cells: List[Dict[str, Any]] = field(default_factory=list)
    diff_shapes: List[Dict[str, Any]] = field(default_factory=list)