"""Pythonの構造フロー・関数依存図・設計書をオフライン生成する。"""

import argparse
import ast
import html
import re
from html.parser import HTMLParser
from pathlib import Path

from pyvis.network import Network


def esc(value):
    return html.escape(str(value), quote=True)


def table(headers, rows):
    head = "".join(f"<th>{esc(v)}</th>" for v in headers)
    body = "".join(
        "<tr>" + "".join(f"<td>{esc(v)}</td>" for v in row) + "</tr>"
        for row in rows
    )
    return f"<table><tr>{head}</tr>{body}</table>"


def analyze(source):
    tree = ast.parse(source)
    classes, functions, conditions, tags, tests = [], [], [], [], []

    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            classes.append((node.lineno, node.name))
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            kind = "async" if isinstance(node, ast.AsyncFunctionDef) else "def"
            functions.append((node.lineno, node.name, kind))
        if isinstance(node, (ast.If, ast.While, ast.IfExp)):
            conditions.append((node.lineno, type(node).__name__, ast.unparse(node.test)))
        if isinstance(node, ast.Call):
            name = ast.unparse(node.func)
            if name in ("get_tag", "set_tag") and node.args:
                tags.append((node.lineno, name, ast.unparse(node.args[0])))
            if name.startswith(("test.", "pytest.")):
                tests.append((node.lineno, name))

    return tree, classes, functions, conditions, tags, tests


def network(hierarchical=False):
    """ローカル描画用の共通ネットワークを生成する。"""
    graph = Network(
        height="900px", width="100%", directed=True,
        bgcolor="#ffffff", font_color="#263348",
        cdn_resources="in_line",
    )
    graph.set_options(f"""{{
      "layout": {{
        "hierarchical": {{
          "enabled": {str(hierarchical).lower()},
          "direction": "UD",
          "sortMethod": "directed",
          "nodeSpacing": 190,
          "levelSeparation": 160
        }}
      }},
      "physics": {{"enabled": false}},
      "nodes": {{
        "font": {{"size": 17, "face": "Consolas"}},
        "margin": 12,
        "widthConstraint": {{"maximum": 350}}
      }},
      "edges": {{
        "arrows": {{"to": {{"enabled": true, "scaleFactor": 0.55}}}},
        "smooth": {{
          "enabled": true,
          "type": "cubicBezier",
          "forceDirection": "vertical"
        }}
      }},
      "interaction": {{
        "hover": true,
        "zoomView": true,
        "dragView": true
      }}
    }}""")
    return graph


def render(graph, destination):
    """オフラインHTMLを生成し、画像コピーボタンを追加する。"""
    if not graph.nodes:
        graph.add_node("_empty", label="対象なし")

    page = graph.generate_html(notebook=False)

    # PyVis標準テンプレート内の不要なBootstrap CDNを除去する。
    def remove_bootstrap(match):
        element = match.group(0)
        opening = element.split(">", 1)[0]
        return "" if re.search(r"https?://[^\s\"']*bootstrap", opening, re.I) else element

    page = re.sub(r"<link\b[^>]*>", remove_bootstrap, page, flags=re.I | re.S)
    page = re.sub(
        r"<script\b[^>]*>.*?</script\s*>",
        remove_bootstrap, page, flags=re.I | re.S,
    )

    class ResourceChecker(HTMLParser):
        def __init__(self):
            super().__init__()
            self.external = []

        def handle_starttag(self, tag, attrs):
            fields = {
                "script": ("src",),
                "link": ("href",),
                "img": ("src", "srcset"),
                "iframe": ("src",),
            }
            for key, value in attrs:
                if key in fields.get(tag, ()) and value and re.match(r"^(?:https?:)?//", value):
                    self.external.append(f"{tag}.{key}={value}")

    checker = ResourceChecker()
    checker.feed(page)

    if checker.external:
        raise RuntimeError("外部リソース参照: " + "; ".join(checker.external))

    # PyVisのCanvasをPNG化し、ブラウザのクリップボードへコピーする。
    # 現在画面に表示されている範囲を対象とする。
    controls = """
<style>
#copy-image {
  position:fixed;top:14px;right:18px;z-index:9999;
  padding:10px 18px;border:1px solid #cdd6e3;border-radius:7px;
  background:#ffffff;color:#243047;font:14px system-ui;
  cursor:pointer;box-shadow:0 2px 8px #0002;
}
#copy-status {
  position:fixed;top:60px;right:18px;z-index:9999;
  font:12px system-ui;color:#243047;background:#fff;
}
</style>
<button id="copy-image" type="button">画像をコピー</button>
<div id="copy-status" role="status"></div>
<script>
document.getElementById("copy-image").addEventListener("click", async () => {
  const status = document.getElementById("copy-status");
  try {
    if (!navigator.clipboard || !window.ClipboardItem) {
      throw new Error("このブラウザでは画像コピーが利用できません");
    }
    if (typeof network === "undefined" || !network.canvas) {
      throw new Error("グラフの描画が完了していません");
    }

    const canvas = network.canvas.frame.canvas;
    const output = document.createElement("canvas");
    output.width = canvas.width;
    output.height = canvas.height;

    const context = output.getContext("2d");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, output.width, output.height);
    context.drawImage(canvas, 0, 0);

    const blob = await new Promise((resolve, reject) => {
      output.toBlob(
        (result) => result ? resolve(result) : reject(new Error("PNG変換失敗")),
        "image/png"
      );
    });

    await navigator.clipboard.write([
      new ClipboardItem({ "image/png": blob })
    ]);
    status.textContent = "画像をコピーしました";
  } catch (error) {
    status.textContent = "コピー失敗: " + error.message;
  }
});
</script>
"""

    page = page.replace("</body>", controls + "</body>")
    destination.write_text(page, encoding="utf-8")


class StructuralFlow:
    """JS版js2flowchartに合わせた構文階層のフロー図を生成する。"""

    COLORS = {
        "start": "#fff0a8",
        "function": "#a8d7ae",
        "block": "#ffe08a",
        "process": "#9bd0f5",
        "call": "#b9df9f",
        "condition": "#d9a6de",
        "exception": "#f4a2a2",
    }

    STEP_Y = 105
    CHILD_X = 165
    TRUE_X = 95
    FALSE_X = 660

    def __init__(self):
        self.graph = network()
        self.serial = 0

    def node(self, label, x, y, kind="process", line=None, full=None):
        self.serial += 1
        name = f"n{self.serial}"

        shape = "diamond" if kind == "condition" else (
            "ellipse" if kind == "start" else "box"
        )

        self.graph.add_node(
            name,
            label=label if len(label) <= 90 else label[:87] + "...",
            title=esc(f"L{line}: {full or label}" if line else full or label),
            shape=shape,
            color=self.COLORS[kind],
            x=x,
            y=y,
            fixed={"x": True, "y": True},
        )
        return name

    def edge(self, source, target, label=""):
        if source is not None:
            self.graph.add_edge(
                source, target, label=label, color="#8b99ac"
            )

    def block(self, statements, x, y, parent, first_label=""):
        """同じ階層の文を並べ、子ブロックの占有範囲を避ける。"""
        cursor = y
        previous = parent
        first = True

        for stmt in statements:
            if (
                isinstance(stmt, ast.Expr)
                and isinstance(stmt.value, ast.Constant)
                and isinstance(stmt.value.value, str)
            ):
                continue

            entry, bottom = self.statement(stmt, x, cursor)
            self.edge(previous, entry, first_label if first else "")

            # JS版同様、制御構文の後続は構文上の親から接続する。
            if isinstance(stmt, (
                ast.If, ast.Try, ast.TryStar, ast.Match,
                ast.For, ast.AsyncFor, ast.While,
            )):
                previous = parent
            else:
                previous = entry

            cursor = max(cursor, bottom) + self.STEP_Y
            first = False

        return cursor

    def statement(self, stmt, x, y):
        """単一文を処理ノードまたは制御構造へ展開する。"""
        line = getattr(stmt, "lineno", None)
        code = ast.unparse(stmt)
        first_line = code.splitlines()[0]

        if isinstance(stmt, ast.If):
            entry = self.node(
                f"if {ast.unparse(stmt.test)}",
                x, y, "condition", line,
            )

            true_bottom = self.block(
                stmt.body, x + self.TRUE_X, y + 145, entry, "True"
            )

            if stmt.orelse:
                false_bottom = self.block(
                    stmt.orelse, x + self.FALSE_X,
                    y + 145, entry, "False"
                )
            else:
                false_bottom = y + 145

            return entry, max(true_bottom, false_bottom)

        if isinstance(stmt, (ast.Try, ast.TryStar)):
            entry = self.node("try", x, y, "block", line)

            # tryの正常系を先に全て描画する。
            bottom = self.block(
                stmt.body, x + self.CHILD_X,
                y + 120, entry,
            )

            # exceptはtryの正常系より下に表示する。
            for handler in stmt.handlers:
                exception = ast.unparse(handler.type) if handler.type else ""
                label = f"except {exception}"
                if handler.name:
                    label += f" as {handler.name}"

                header_y = bottom + 30
                handler_node = self.node(
                    label, x + self.CHILD_X,
                    header_y, "exception", handler.lineno,
                )
                self.edge(entry, handler_node, "except")

                bottom = self.block(
                    handler.body, x + self.CHILD_X * 2,
                    header_y + 125, handler_node,
                )

            if stmt.orelse:
                header_y = bottom + 30
                other = self.node(
                    "else", x + self.CHILD_X,
                    header_y, "block",
                )
                self.edge(entry, other)
                bottom = self.block(
                    stmt.orelse, x + self.CHILD_X * 2,
                    header_y + 125, other,
                )

            if stmt.finalbody:
                header_y = bottom + 30
                final = self.node(
                    "finally", x + self.CHILD_X,
                    header_y, "block",
                )
                self.edge(entry, final)
                bottom = self.block(
                    stmt.finalbody, x + self.CHILD_X * 2,
                    header_y + 125, final,
                )

            return entry, bottom

        if isinstance(stmt, (
            ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef,
        )):
            category = "Class" if isinstance(stmt, ast.ClassDef) else "Function"
            entry = self.node(
                f"{category} {stmt.name}()",
                x, y, "function", line,
            )
            bottom = self.block(
                stmt.body, x + self.CHILD_X,
                y + 120, entry,
            )
            return entry, bottom

        if isinstance(stmt, (ast.For, ast.AsyncFor, ast.While)):
            condition = (
                ast.unparse(stmt.test)
                if isinstance(stmt, ast.While)
                else f"{ast.unparse(stmt.target)} in {ast.unparse(stmt.iter)}"
            )
            entry = self.node(
                f"{type(stmt).__name__} {condition}",
                x, y, "condition", line,
            )
            body_bottom = self.block(
                stmt.body, x + self.TRUE_X,
                y + 145, entry, "反復",
            )
            else_bottom = (
                self.block(
                    stmt.orelse, x + self.FALSE_X,
                    y + 145, entry, "else",
                )
                if stmt.orelse else y + 145
            )
            return entry, max(body_bottom, else_bottom)

        if isinstance(stmt, (ast.With, ast.AsyncWith)):
            entry = self.node(first_line, x, y, "block", line)
            bottom = self.block(
                stmt.body, x + self.CHILD_X,
                y + 120, entry,
            )
            return entry, bottom

        if isinstance(stmt, ast.Match):
            entry = self.node(
                f"match {ast.unparse(stmt.subject)}",
                x, y, "condition", line,
            )
            bottoms = []

            for index, case in enumerate(stmt.cases):
                case_x = x + self.CHILD_X + index * self.FALSE_X
                label = f"case {ast.unparse(case.pattern)}"
                if case.guard:
                    label += f" if {ast.unparse(case.guard)}"

                case_node = self.node(
                    label, case_x, y + 145, "block"
                )
                self.edge(entry, case_node, f"case {index + 1}")
                bottoms.append(
                    self.block(
                        case.body, case_x + self.CHILD_X,
                        y + 270, case_node,
                    )
                )

            return entry, max(bottoms, default=y + 145)

        kind = "process"
        if isinstance(stmt, ast.Expr) and isinstance(stmt.value, ast.Call):
            kind = "call"

        entry = self.node(first_line, x, y, kind, line, code)
        return entry, y

    def build(self, tree):
        """ファイル直下の全構文を1枚の図に配置する。"""
        start = self.node(
            "Program: source module", 0, 0, "start"
        )
        self.block(tree.body, 0, 130, start)
        return self.graph


def generate_calls(tree):
    """ASTから関数・メソッドの呼び出し関係を生成する。"""
    graph = network(hierarchical=True)
    definitions = {}

    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            definitions[node.name] = node
        elif isinstance(node, ast.ClassDef):
            for child in node.body:
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    definitions[f"{node.name}.{child.name}"] = child

    class Visitor(ast.NodeVisitor):
        def __init__(self):
            self.calls = []

        def visit_Call(self, node):
            self.calls.append(node.func)
            self.generic_visit(node)

        def visit_FunctionDef(self, node):
            return

        visit_AsyncFunctionDef = visit_FunctionDef
        visit_ClassDef = visit_FunctionDef

    module = ast.Module(
        body=[
            node for node in tree.body
            if not isinstance(
                node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)
            )
        ],
        type_ignores=[],
    )

    added, edges = set(), set()

    def add(name):
        if name not in added:
            graph.add_node(
                name, label=name, shape="box",
                color="#dceafe" if name in definitions else "#edf2f8",
            )
            added.add(name)

    for owner, definition in {**definitions, "<module>": module}.items():
        add(owner)
        visitor = Visitor()

        for statement in definition.body:
            visitor.visit(statement)

        for call in visitor.calls:
            target = ast.unparse(call)
            if "." in owner and target.startswith(("self.", "cls.")):
                target = owner.split(".")[0] + "." + target.split(".", 1)[1]

            add(target)
            if (owner, target) not in edges:
                graph.add_edge(owner, target)
                edges.add((owner, target))

    return graph


def generate(files, output):
    """設計書と2種類の図を作り、最後にエラーを集計する。"""
    flows = output / "flows"
    flows.mkdir(parents=True, exist_ok=True)

    sections = []
    failures = []

    for file in files:
        prefix = file.as_posix().replace("/", "_").replace(".", "_")

        try:
            source = file.read_text(encoding="utf-8")
            tree, classes, functions, conditions, tags, tests = analyze(source)
        except Exception as error:
            failures.append(f"{file}: AST: {error}")
            print(f"[FAIL] AST {file}: {error}")
            continue

        links = []

        for label, filename, factory in (
            ("処理フロー", f"{prefix}.flow.html", lambda: StructuralFlow().build(tree)),
            ("関数依存図", f"{prefix}.calls.html", lambda: generate_calls(tree)),
        ):
            try:
                render(factory(), flows / filename)
                links.append(
                    f'<a href="flows/{esc(filename)}">{esc(label)}</a>'
                )
                print(f"[PASS] {label}: {file}")
            except Exception as error:
                failures.append(f"{file}: {label}: {error}")
                links.append(f"{esc(label)}: FAIL ({esc(error)})")
                print(f"[FAIL] {label}: {file}: {error}")

        if output.name != "test":
            links.append(
                f'<a href="../coverage/{esc(file.stem)}/index.html">'
                "カバレッジ</a>"
            )

        sections.append(f"""
<section><h2>{esc(file)}</h2>
<h3>クラス一覧</h3>{table(["行", "クラス"], classes)}
<h3>関数・メソッド一覧</h3>{table(["行", "名前", "種類"], functions)}
<h3>制御条件</h3>{table(["行", "構文", "条件式"], conditions)}
<h3>入出力操作</h3>{table(["行", "操作", "引数"], tags)}
<h3>テスト呼び出し</h3>{table(["行", "呼び出し"], tests)}
<h3>解析成果物</h3><p>{" | ".join(links)}</p>
</section>""")

    diagnostics = "".join(f"<li>{esc(error)}</li>" for error in failures)
    api_link = (
        '<p><a href="api/index.html">API仕様書</a></p>'
        if output.name == "test" else ""
    )

    page = f"""<!doctype html><html lang="ja">
<head><meta charset="utf-8"><title>Python Design Report</title>
<style>
body{{font:15px system-ui,sans-serif;max-width:1100px;margin:35px auto;
padding:20px;background:#f5f7fb}}
section{{background:white;padding:25px;margin:20px 0;border-radius:10px}}
table{{width:100%;border-collapse:collapse;margin:12px 0}}
td,th{{padding:10px;border:1px solid #dce3ed;text-align:left}}
th{{background:#edf2f8}}a{{color:#2466c0}}
</style></head><body>
<h1>Python 設計書</h1>
{api_link}
{''.join(sections)}
{"<h2>生成エラー</h2><ul>" + diagnostics + "</ul>" if failures else ""}
</body></html>"""

    (output / "index.html").write_text(page, encoding="utf-8")

    print(f"Files: {len(files)}, Failed: {len(failures)}")
    print(f"HTML: {(output / 'index.html').resolve()}")

    if not files or failures:
        raise RuntimeError(
            f"Documentation generation failed: {len(failures)} errors"
        )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true")
    args = parser.parse_args()

    if args.test:
        files = [Path("framework.py")]
        files += sorted(Path("tests").rglob("*.py"))
        files += sorted(Path("tools").rglob("*.py"))
        output = Path("docs/test")
    else:
        files = sorted(Path("target").rglob("*.py"))
        output = Path("docs")

    generate(files, output)


if __name__ == "__main__":
    main()