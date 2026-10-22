"""C2計測の基本動作を検証する。"""

from framework import PythonDriver


def test_c2_and_short_circuit(tmp_path):
    """andの左右を個別計測し、短絡評価を保持する。"""
    file = tmp_path / "sample.py"
    file.write_text(
        'a = get_tag("a")\n'
        'b = get_tag("b")\n'
        'if a > 10 and b < 20:\n'
        '    set_tag("result", True)\n'
        'else:\n'
        '    set_tag("result", False)\n',
        encoding="utf-8",
    )

    driver = PythonDriver(file)
    driver.run({"a": 0, "b": 0})
    driver.run({"a": 11, "b": 30})
    result = driver.run({"a": 11, "b": 10})

    assert result["tags"]["result"] is True
    assert driver.summary()["c2"] == 100
    assert len(driver.summary()["conditions"]) == 2


def test_c2_or_short_circuit(tmp_path):
    """or条件でも真偽を個別に計測する。"""
    file = tmp_path / "sample.py"
    file.write_text(
        'a = get_tag("a")\n'
        'b = get_tag("b")\n'
        'if a > 10 or b > 20:\n'
        '    set_tag("result", True)\n'
        'else:\n'
        '    set_tag("result", False)\n',
        encoding="utf-8",
    )

    driver = PythonDriver(file)
    driver.run({"a": 11, "b": 0})
    driver.run({"a": 0, "b": 21})
    driver.run({"a": 0, "b": 0})

    assert driver.summary()["c2"] == 100