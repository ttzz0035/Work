from typing import Optional, Callable
import re
from openpyxl.utils import column_index_from_string
import xlwings as xw

def compile_matcher(keyword: str, use_regex: bool = False, ignore_case: bool = False) -> Callable[[str], bool]:
    """文字列を受け取り一致判定する関数を返す"""
    if use_regex:
        flags = re.IGNORECASE if ignore_case else 0
        reg = re.compile(keyword, flags)
        return lambda s: (False if s is None else reg.search(str(s)) is not None)
    else:
        tgt = keyword.lower() if ignore_case else keyword
        return lambda s: False if s is None else \
            (tgt in str(s).lower() if ignore_case else tgt in str(s))

def find_in_column(sht: xw.Sheet, col_letter: str, matcher: Callable[[str], bool]) -> Optional[int]:
    """列を上から走査して一致する行番号を返す"""
    col = column_index_from_string(col_letter.upper())
    used = sht.used_range
    for r in range(used.row, used.last_cell.row + 1):
        if matcher(sht.range((r, col)).value):
            return r
    return None

def find_in_row(sht: xw.Sheet, row_num: int, matcher: Callable[[str], bool]) -> Optional[int]:
    """行を左から走査して一致する列番号を返す"""
    used = sht.used_range
    for c in range(used.column, used.last_cell.column + 1):
        if matcher(sht.range((row_num, c)).value):
            return c
    return None
