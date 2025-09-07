# excel_transfer/models/dto.py
from dataclasses import dataclass
from typing import List, Callable, Optional, Literal, Tuple

LogFn = Callable[[str], None]

@dataclass
class TransferRequest:
    csv_paths: List[str]
    # 追加: 範囲外セルの扱い
    out_of_range_mode: Literal["clamp", "skip", "error"] = "clamp"

@dataclass
class GrepRequest:
    root_dir: str
    keyword: str
    ignore_case: bool = True
    use_regex: bool = False

@dataclass
class DiffRequest:
    file_a: str
    file_b: str
    key_cols: List[str]
    compare_formula: bool = False
    include_context: bool = True
    context_radius: int = 2
    max_context_items: int = 30
    compare_shapes: bool = False

@dataclass
class CountRequest:
    files: List[str]
    sheet: str
    start_cell: str
    direction: Literal["row","col"]
    tolerate_blanks: int = 0
    mode: Literal["jump","scan"] = "jump"
