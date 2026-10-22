"""Ruff・pytest・Docsを一括実行し、全工程の結果を集計する。"""

import subprocess
import sys
from pathlib import Path


def main():
    """一部が失敗しても後続を実行し、最後に総合判定する。"""
    root = Path(__file__).resolve().parent.parent

    tasks = [
        (
            "Ruff 静的解析",
            [sys.executable, "-m", "ruff", "check", "target"],
        ),
        (
            "pytest C0/C1/C2",
            [
                sys.executable,
                "-m",
                "pytest",
                "-v",
                "--html=docs/test-results.html",
                "--self-contained-html",
            ],
        ),
        (
            "Docs 設計書・処理フロー・関数依存図",
            [sys.executable, "-m", "tools.docs"],
        ),
    ]

    (root / "docs").mkdir(parents=True, exist_ok=True)

    results = []

    for name, command in tasks:
        print(f"\n{'=' * 65}", flush=True)
        print(f"[RUN] {name}", flush=True)
        print(f"{'=' * 65}", flush=True)

        try:
            completed = subprocess.run(command, cwd=root, check=False)
            passed = completed.returncode == 0
            results.append((name, passed))
        except OSError as error:
            print(f"[FAIL] {name}: {error}", flush=True)
            results.append((name, False))

    print(f"\n{'=' * 65}", flush=True)
    print("QUALITY SUMMARY", flush=True)
    print(f"{'=' * 65}", flush=True)

    for name, passed in results:
        print(f'[{"PASS" if passed else "FAIL"}] {name}', flush=True)

    print("\nGenerated reports:", flush=True)

    reports = [
        root / "docs" / "test-results.html",
        root / "docs" / "index.html",
    ]

    for report in reports:
        state = "EXISTS" if report.is_file() else "MISSING"
        print(f"[{state}] {report.relative_to(root)}", flush=True)

    failures = sum(not passed for _, passed in results)
    print(f"\nResult: {len(results) - failures}/{len(results)} PASS", flush=True)

    if failures:
        raise SystemExit(1)


if __name__ == "__main__":
    main()