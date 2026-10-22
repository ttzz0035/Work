"""Pythonコード実行、HTTPスタブ、C0/C1/C2計測。"""

import ast
import html
import json
import sys
import types
from collections import defaultdict
from pathlib import Path
from unittest.mock import patch

from coverage import Coverage


class C2Transformer(ast.NodeTransformer):
    """条件式の真偽を記録するAST変換器。"""

    def __init__(self):
        self.conditions = {}
        self.sequence = 0

    def measure(self, node):
        # and/orは各基本条件を個別計測し、短絡評価を保持する。
        if isinstance(node, ast.BoolOp):
            node.values = [self.measure(x) for x in node.values]
            return node

        key = str(self.sequence)
        self.sequence += 1
        self.conditions[key] = {
            "line": node.lineno,
            "expression": ast.unparse(node),
            "true": 0,
            "false": 0,
        }
        call = ast.Call(
            func=ast.Name(id="__c2", ctx=ast.Load()),
            args=[ast.Constant(key), node],
            keywords=[],
        )
        return ast.copy_location(call, node)

    def visit_If(self, node):
        node.test = self.measure(self.visit(node.test))
        node.body = [self.visit(x) for x in node.body]
        node.orelse = [self.visit(x) for x in node.orelse]
        return node

    def visit_While(self, node):
        node.test = self.measure(self.visit(node.test))
        node.body = [self.visit(x) for x in node.body]
        node.orelse = [self.visit(x) for x in node.orelse]
        return node

    def visit_IfExp(self, node):
        node.test = self.measure(self.visit(node.test))
        node.body = self.visit(node.body)
        node.orelse = self.visit(node.orelse)
        return node


class HttpStub:
    """requests.getを模擬するHTTPクライアント。"""

    def __init__(self, status=200, data=None, error=None):
        self.status = status
        self.data = data if data is not None else {}
        self.error = error
        self.calls = []

    def get(self, url, **kwargs):
        self.calls.append({"url": url, "kwargs": kwargs})
        if self.error:
            raise ConnectionError(self.error)

        response = types.SimpleNamespace(
            status_code=self.status,
            status=self.status,
            text=json.dumps(self.data),
            json=lambda: self.data,
        )

        def raise_for_status():
            if self.status >= 400:
                error = RuntimeError(f"HTTP {self.status}")
                error.response = response
                raise error

        response.raise_for_status = raise_for_status
        return response


class PythonDriver:
    """Pythonスクリプトを条件計測付きで実行する。"""

    def __init__(self, file):
        self.file = Path(file).resolve()
        self.source = self.file.read_text(encoding="utf-8")
        tree = ast.parse(self.source, filename=str(self.file))
        self.transformer = C2Transformer()
        tree = ast.fix_missing_locations(self.transformer.visit(tree))
        self.code = compile(tree, str(self.file), "exec")
        self.coverage = Coverage(
            branch=True, include=[str(self.file)],
            data_file=None,
        )
        self.runs = 0
        self.exceptions = []

    def run(self, tags=None, http=None):
        """入力タグとHTTP応答を指定して1回実行する。"""
        tags = dict(tags or {})
        writes = []
        client = HttpStub(**(http or {}))

        def get_tag(name):
            if name not in tags:
                raise KeyError(f"Unknown tag: {name}")
            return tags[name]

        def set_tag(name, value):
            tags[name] = value
            writes.append((name, value))

        def c2(key, value):
            condition = self.transformer.conditions[key]
            condition["true" if bool(value) else "false"] += 1
            return value

        # requestsを仮想モジュールに置換し、実通信を防止する。
        requests = types.ModuleType("requests")
        requests.get = client.get
        requests.exceptions = types.SimpleNamespace(
            RequestException=Exception, ConnectionError=ConnectionError
        )
        environment = {
            "__name__": "__tested_script__",
            "__file__": str(self.file),
            "__c2": c2,
            "get_tag": get_tag,
            "set_tag": set_tag,
            "$getTag": get_tag,
            "$setTag": set_tag,
            "requests": requests,
        }

        self.runs += 1
        self.coverage.start()
        try:
            with patch.dict(sys.modules, {"requests": requests}):
                exec(self.code, environment)
        except Exception as error:
            self.exceptions.append(repr(error))
            raise
        finally:
            self.coverage.stop()

        return {"tags": tags, "writes": writes, "http_calls": client.calls}

    def summary(self):
        """C2の個別条件と網羅率を返す。"""
        conditions = list(self.transformer.conditions.values())
        covered = sum(c["true"] > 0 and c["false"] > 0 for c in conditions)
        return {
            "runs": self.runs,
            "c2": 100 * covered / len(conditions) if conditions else 100.0,
            "conditions": conditions,
        }

    def report(self, directory, threshold=100):
        """標準レポートとC1/C2比較HTMLを生成する。"""
        output = Path(directory)
        output.mkdir(parents=True, exist_ok=True)
        self.coverage.html_report(directory=str(output / "lines"))
        self.coverage.json_report(outfile=str(output / "coverage.json"))

        data = json.loads((output / "coverage.json").read_text("utf-8"))
        detail = next(iter(data["files"].values()))
        total = detail["summary"]
        c0 = (
            100 * total["covered_lines"] / total["num_statements"]
            if total["num_statements"] else 100.0
        )
        c1 = 100 * total["covered_branches"] / total["num_branches"] if total["num_branches"] else 100
        c2 = self.summary()["c2"]
        branches = defaultdict(lambda: {"covered": 0, "missing": 0})

        for line, _ in detail.get("executed_branches", []):
            branches[line]["covered"] += 1
        for line, _ in detail.get("missing_branches", []):
            branches[line]["missing"] += 1

        rows = []
        for c in self.transformer.conditions.values():
            b = branches.get(c["line"])
            c1_result = f'{b["covered"]}実行 / {b["missing"]}未実行' if b else "該当なし"
            passed = c["true"] > 0 and c["false"] > 0
            rows.append(
                f'<tr><td>{c["line"]}</td><td>{html.escape(c["expression"])}</td>'
                f'<td>{c1_result}</td><td>{c["true"]}</td><td>{c["false"]}</td>'
                f'<td>{"PASS" if passed else "FAIL"}</td></tr>'
            )

        page = f"""<!doctype html><html lang="ja"><meta charset="utf-8">
<title>Python Coverage</title>
<style>
body{{font:15px system-ui;max-width:1100px;margin:35px auto;padding:15px;background:#f5f7fb}}
main{{background:white;padding:25px;border-radius:10px}}
table{{width:100%;border-collapse:collapse}}
td,th{{padding:10px;border:1px solid #dce3ed;text-align:left}}
th{{background:#edf2f8}}
</style><main>
<h1>{html.escape(self.file.name)} カバレッジ</h1>
<h2>C0: {c0:.2f}%　C1: {c1:.2f}%　C2: {c2:.2f}%</h2>
<p><a href="lines/index.html">標準カバレッジ詳細・ソースコード</a></p>
<h3>C1とC2の条件別比較</h3>
<table><tr><th>行</th><th>基本条件</th><th>C1分岐</th>
<th>C2 True</th><th>C2 False</th><th>C2判定</th></tr>
{''.join(rows)}</table>
<p>C1は分岐経路、C2は基本条件それぞれの真偽を計測します。</p>
</main></html>"""
        (output / "index.html").write_text(page, encoding="utf-8")
        if min(c0, c1, c2) + 1e-8 < threshold:
            raise AssertionError(f"Coverage insufficient: C0={c0:.2f}, C1={c1:.2f}, C2={c2:.2f}")
        return {"c0": c0, "c1": c1, "c2": c2}