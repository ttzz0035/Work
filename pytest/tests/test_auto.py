"""target/配下のPythonファイルを自動テストする。"""

import ast
import itertools
from pathlib import Path

import pytest
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st

from framework import PythonDriver


TARGET = Path("target")
FILES = sorted(TARGET.rglob("*.py"))


def analyze_inputs(source):
    """タグ名と比較条件の境界値を抽出する。"""
    tree = ast.parse(source)
    tags = set()
    thresholds = set()

    for node in ast.walk(tree):
        # get_tag("temperature")形式の入力を認識する。
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            if node.func.id in ("get_tag", "$getTag") and node.args:
                arg = node.args[0]
                if isinstance(arg, ast.Constant) and isinstance(arg.value, str):
                    tags.add(arg.value)

        # 比較式中の数値を集めて境界値を生成する。
        if isinstance(node, ast.Compare):
            for child in ast.walk(node):
                if isinstance(child, ast.Constant):
                    value = child.value
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        thresholds.add(value)

    values = {-1, 0, 1, 50, 100}
    for value in thresholds:
        values.update([value - 1, value, value + 1])

    return sorted(tags), sorted(values)


def cases(names, values, limit=512):
    """境界値を組み合わせて入力ケースを作る。"""
    if not names:
        yield {}
        return

    for combination in itertools.islice(
        itertools.product(values, repeat=len(names)), limit
    ):
        yield dict(zip(names, combination))


def run_case(driver, tags, http):
    """1ケース実行し、例外も診断情報として保持する。"""
    try:
        driver.run(tags, http)
    except Exception as error:
        # エラーを無視せず、ケース単位で記録する。
        return repr(error)
    return None


@pytest.mark.parametrize("file", FILES, ids=lambda p: str(p))
def test_automatic_coverage(file):
    """境界値・HTTP応答・Hypothesisで網羅率を検証する。"""
    driver = PythonDriver(file)
    names, values = analyze_inputs(driver.source)

    # HTTP成功、異常応答、通信失敗を模擬する。
    http_cases = [
        {"status": 200, "data": {"running": True}},
        {"status": 200, "data": {"running": False}},
        {"status": 201, "data": {}},
        {"status": 400, "data": {}},
        {"status": 500, "data": {}},
        {"error": "network error"},
    ]

    errors = set()

    # HTTP依存を持たないコードではHTTPシナリオを増やさない。
    uses_http = "requests" in driver.source
    scenarios = http_cases if uses_http else [None]

    for tags in cases(names, values):
        for scenario in scenarios:
            error = run_case(driver, tags, scenario)
            if error:
                errors.add(error)

    # Hypothesisによる追加の入力生成。
    if names:
        strategy = st.fixed_dictionaries({
            name: st.one_of(
                st.sampled_from(values),
                st.integers(-1000, 1000),
            )
            for name in names
        })

        @settings(
            max_examples=100,
            deadline=None,
            suppress_health_check=[HealthCheck.too_slow],
        )
        @given(strategy)
        def explore(tags):
            for scenario in scenarios:
                error = run_case(driver, tags, scenario)
                if error:
                    errors.add(error)

        explore()

    # 網羅率不足でもHTMLを生成してからFAILにする。
    result = driver.report(Path("coverage") / file.stem, threshold=100)
    print(f"\n{file}: {result}, executions={driver.runs}")

    # 未処理例外があった場合も明示して失敗させる。
    assert not errors, f"Unhandled exceptions: {sorted(errors)}"


def test_target_exists():
    """テスト対象が存在しない場合の誤PASSを防ぐ。"""
    assert FILES, "target/にPythonファイルがありません"