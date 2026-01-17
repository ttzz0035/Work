# excel_transfer/models/dto.py
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, List


LogFn = Callable[[str], None]


@dataclass
class TransferRequest:
    csv_paths: List[str] = field(default_factory=list)
    out_of_range_mode: str = "error"


@dataclass
class GrepRequest:
    root_dir: str = ""
    keyword: str = ""
    ignore_case: bool = True
    use_regex: bool = False


@dataclass
class CountRequest:
    files: List[str] = field(default_factory=list)
    sheet: str = ""
    start_cell: str = "B2"
    direction: str = "row"
    tolerate_blanks: int = 0
    mode: str = "jump"


@dataclass
class DiffRequest:
    file_a: str = ""
    file_b: str = ""
    range_a: str = ""
    range_b: str = ""

    # ★ 追加（UI / Service 両対応）
    base_file: str = "B"   # "A" or "B"

    # 既存互換用
    key_cols: List[str] = field(default_factory=list)

    compare_formula: bool = False
    include_context: bool = True
    compare_shapes: bool = False

    sheet_a: str = ""
    sheet_b: str = ""
