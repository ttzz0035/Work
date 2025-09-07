# tests/test_grep.py
import csv
from pathlib import Path
import openpyxl
from services.grep import run_grep
from models.dto import GrepRequest
from tests.conftest import wb_write

def test_grep_case_insensitive(ctx, tmp_path):
    # 検索対象のExcelを作成（2ファイル）
    f1 = tmp_path / "a.xlsx"
    f2 = tmp_path / "b.xlsx"
    wb_write(f1, {"S": [["Header"], ["Foobar"], ["bar"]]})
    wb_write(f2, {"S": [["Header"], ["foo"], ["zzz"]]})

    # 実行（ignore_case=True, keyword='FOO'）
    out, cnt = run_grep(
        GrepRequest(root_dir=str(tmp_path), keyword="FOO", ignore_case=True),
        ctx, logger=None, append_log=lambda *_: None
    )
    # 検証
    assert Path(out).name == "grep_results.csv"
    assert cnt == 2  # "Foobar" と "foo" の2ヒット

    # CSV内容ざっくり確認
    rows = list(csv.DictReader(open(out, encoding="utf-8-sig")))
    assert any("Foobar" in r["value"] for r in rows)
    assert any(r["value"] == "foo" for r in rows)
