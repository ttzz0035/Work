# excel_transfer/models/dto.py
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, List


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
    root_dir: str = ""
    keyword: str = ""
    ignore_case: bool = True
    use_regex: bool = False


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

    # 既存互換用（未使用でも残す）
    key_cols: List[str] = field(default_factory=list)

    # 比較オプション
    compare_formula: bool = False
    include_context: bool = True
    compare_shapes: bool = False

    # ★ 追加：シート比較モード
    # "index" : インデックス一致（既定・後方互換）
    # "name"  : シート名一致
    sheet_mode: str = "index"

    # 将来拡張用（明示シート指定）
    sheet_a: str = ""
    sheet_b: str = ""
