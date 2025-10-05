# SPDX-License-Identifier: MIT
"""
DB操作の窓口。実運用では本ファイルの中身を差し替えてください。
UI側は get_items() のみを利用します。
"""
from typing import List, Tuple

def get_items() -> List[Tuple[int, str]]:
    """
    Dropdownに供給する項目リストを返す。
    返却形式: [(id:int, name:str), ...]
    """
    # TODO: 実DBから取得する実装に差し替え
    return [(1, "Job A"), (2, "Job B"), (3, "Job C")]
