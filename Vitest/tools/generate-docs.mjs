import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { parse } from "@babel/parser";
import traverseModule from "@babel/traverse";
import generateModule from "@babel/generator";
import * as t from "@babel/types";
import { Resvg } from "@resvg/resvg-js";

const require = createRequire(import.meta.url);
const flow = require("js2flowchart");
const traverse = traverseModule.default ?? traverseModule;
const generate = generateModule.default ?? generateModule;

const ROOT = process.cwd();
const SOURCE = path.join(ROOT, "targets");
const DOCS = path.join(ROOT, "docs");
const FLOWS = path.join(DOCS, "flows");
const JSDOC = path.join(DOCS, "jsdoc-src");

// フロー図で式を省略しないための表示文字数上限。
const MAX_FLOW_TEXT_LENGTH = 500;

function esc(value) {
  return String(value).replace(/[&<>"']/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}

function discover(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(entry => {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) return discover(file);
    return entry.isFile() && file.endsWith(".js") ? [file] : [];
  }).sort();
}

function wrap(source) {
  return `async function fuxaScript() {\n${source}\n}`;
}

function analyze(source) {
  const tree = parse(wrap(source), { sourceType: "script" });
  const inputs = new Set();
  const outputs = new Set();
  const conditions = [];

  traverse(tree, {
    CallExpression(p) {
      if (p.node.callee.type !== "Identifier") return;
      const name = p.node.callee.name;
      if (name !== "$getTag" && name !== "$setTag") return;

      const arg = p.node.arguments[0];
      const tag = arg?.type === "StringLiteral"
        ? arg.value : `[動的] ${arg ? generate(arg).code : "引数なし"}`;

      if (name === "$getTag") inputs.add(tag);
      else outputs.add(tag);
    },
    "IfStatement|ConditionalExpression|WhileStatement|DoWhileStatement|ForStatement"(p) {
      if (!p.node.test) return;
      conditions.push({
        line: p.node.loc.start.line - 1,
        kind: p.node.type,
        expression: generate(p.node.test).code,
      });
    },
    SwitchStatement(p) {
      conditions.push({
        line: p.node.loc.start.line - 1,
        kind: "SwitchStatement",
        expression: generate(p.node.discriminant).code,
      });
    },
  });

  conditions.sort((a, b) => a.line - b.line);
  return { inputs, outputs, conditions };
}

function description(source, name) {
  const match = source.match(/@description\s+([^\r\n*]+)/);
  if (match) return match[1].trim();
  const comment = source.match(/^\s*\/\*\*?\s*(.+)/);
  return comment ? comment[1].trim() : name;
}

// テンプレートリテラルをフロー図専用の単一表示式へ変換する。
// ${}内の参照変数を独立した処理ノードとして展開させない。
// 変換は図の生成時だけ行い、実際のソースやテストには影響させない。
function flowDisplayCode(code) {
  const tree = parse(code, { sourceType: "script" });

  traverse(tree, {
    TemplateLiteral(p) {
      // テンプレート文字列全体を、元の表記を保持した文字列にする。
      const display = generate(p.node, { comments: false }).code;
      p.replaceWith(t.stringLiteral(display));
      p.skip();
    },
  });

  return generate(tree, { comments: true }).code;
}

// SVGの表示上限を拡大し、長い代入式の省略を抑制する。
function renderFlow(code) {
  const tree = flow.convertCodeToFlowTree(flowDisplayCode(code));
  const renderer = flow.createSVGRender();

  renderer.applyTheme({
    common: {
      maxNameLength: MAX_FLOW_TEXT_LENGTH,
    },
  });

  return renderer.buildShapesTree(tree).print();
}

// SVGは表示用、PNGはクリップボード用として分離する。
function diagramViewer(svgName, title, png) {
  const base64 = png.toString("base64");

  return `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>${esc(title)}</title>
<style>
body{margin:0;background:#fff;font:14px system-ui;color:#243047}
#copy-image{
  position:fixed;top:14px;right:18px;z-index:9999;
  padding:10px 18px;border:1px solid #cdd6e3;border-radius:7px;
  background:#fff;color:#243047;font:14px system-ui;
  cursor:pointer;box-shadow:0 2px 8px #0002
}
#copy-image:disabled{opacity:.5}
#copy-status{
  position:fixed;top:60px;right:18px;z-index:9999;
  font:12px system-ui;color:#243047;background:#fff
}
.diagram{overflow:auto;padding:20px}
#diagram{display:block;max-width:none;height:auto}
</style>
</head>
<body>
<button id="copy-image" type="button">画像をコピー</button>
<div id="copy-status" role="status"></div>
<div class="diagram">
<img id="diagram" src="${esc(svgName)}" alt="${esc(title)}">
</div>
<script>
const button = document.getElementById("copy-image");
const status = document.getElementById("copy-status");
const pngBase64 = "${base64}";

button.addEventListener("click", async () => {
  button.disabled = true;
  try {
    if (!navigator.clipboard?.write || !window.ClipboardItem) {
      throw new Error("画像クリップボードが利用できません");
    }

    // 事前生成PNGを復元する。Canvasは使用しない。
    const bytes = Uint8Array.from(atob(pngBase64), c => c.charCodeAt(0));
    const blob = new Blob([bytes], {type: "image/png"});
    await navigator.clipboard.write([
      new ClipboardItem({"image/png": Promise.resolve(blob)})
    ]);
    status.textContent = "画像をコピーしました";
  } catch (error) {
    status.textContent = "コピー失敗: " + error.message;
  } finally {
    button.disabled = false;
  }
});
</script>
</body>
</html>`;
}

function diagrams(code, name) {
  // 処理フローのみ表示用ASTを使用する。
  const flowSvg = renderFlow(code);

  // 関数依存図は元のコードをそのまま解析する。
  const builder = flow.createFlowTreeBuilder();
  builder.setAbstractionLevel(flow.ABSTRACTION_LEVELS.FUNCTION_DEPENDENCIES);
  const callsSvg = flow.convertFlowTreeToSvg(builder.build(code));

  for (const [type, svg, title] of [
    ["flow", flowSvg, "処理フロー"],
    ["calls", callsSvg, "関数依存図"],
  ]) {
    // 変換は生成時に行い、ブラウザ側では画像を再描画しない。
    const png = new Resvg(svg, { background: "white" }).render().asPng();
    if (!png.length) throw new Error(`${name}.${type}: PNG生成失敗`);

    const svgName = `${name}.${type}.svg`;
    fs.writeFileSync(path.join(FLOWS, svgName), svg, "utf8");
    fs.writeFileSync(path.join(FLOWS, `${name}.${type}.png`), png);
    fs.writeFileSync(
      path.join(FLOWS, `${name}.${type}.html`),
      diagramViewer(svgName, title, png),
      "utf8"
    );
  }
}

function tagTable(inputs, outputs) {
  const rows = [
    ...[...inputs].map(v => ["入力", v]),
    ...[...outputs].map(v => ["出力", v]),
  ];
  return `<table><tr><th>方向</th><th>タグ名</th></tr>${rows.map(
    ([direction, name]) =>
      `<tr><td>${esc(direction)}</td><td><code>${esc(name)}</code></td></tr>`
  ).join("")}</table>`;
}

function conditionTable(conditions) {
  return `<table><tr><th>行</th><th>構文</th><th>条件式</th></tr>${conditions.map(
    c => `<tr><td>${c.line}</td><td>${esc(c.kind)}</td><td><code>${esc(c.expression)}</code></td></tr>`
  ).join("")}</table>`;
}

function jsdocSource(source, name, summary) {
  // ファイル先頭のコメントを取り出す。
  const header = source.match(/^\s*(\/\*\*?[\s\S]*?\*\/)/);
  const body = header ? source.slice(header[0].length) : source;

  let moduleDoc;

  if (header) {
    // 通常のブロックコメントもJSDocとして扱う。
    const content = header[1]
      .replace(/^\/\*\*?/, "")
      .replace(/\*\/$/, "")
      .split(/\r?\n/)
      .map(line => line.replace(/^\s*\*?\s?/, ""))
      .map(line => line.replace(/^@file\s+/, "ファイル: "));

    moduleDoc = [
      "/**",
      ` * @module ${name}`,
      ...content.map(line => line ? ` * ${line}` : " *"),
      " */",
    ].join("\n");
  } else {
    moduleDoc = [
      "/**",
      ` * @module ${name}`,
      ` * @description ${summary.replace(/\*\//g, "* /")}`,
      " */",
    ].join("\n");
  }

  // awaitを含むFUXAスクリプトを解析する。
  const tree = parse(wrap(body), {
    sourceType: "script",
    attachComment: true,
  });

  const statements = tree.program.body[0].body.body;

  // 実際の関数・クラス定義だけをJSDoc用ソースへ出力する。
  const definitions = statements
    .filter(node => [
      "FunctionDeclaration",
      "ClassDeclaration",
    ].includes(node.type))
    .map(node => generate(node, { comments: true }).code);

  return [moduleDoc, ...definitions].join("\n\n") + "\n";
}

function build(file) {
  const relative = path.relative(SOURCE, file);
  const name = relative.replace(/[\\/]/g, "_").replace(/\.js$/, "");
  const source = fs.readFileSync(file, "utf8");
  const summary = description(source, name);
  const failures = [];

  // 各成果物を独立して生成し、一部が失敗しても継続する。
  let tagHtml = "";
  let conditionHtml = "";

  try {
    const info = analyze(source);
    tagHtml = tagTable(info.inputs, info.outputs);
    conditionHtml = conditionTable(info.conditions);
    console.log(`[PASS] AST: ${relative}`);
  } catch (error) {
    failures.push(`AST: ${error.message}`);
  }

  try {
    fs.writeFileSync(
      path.join(JSDOC, `${name}.js`),
      jsdocSource(source, name, summary),
      "utf8"
    );
    console.log(`[PASS] JSDoc: ${relative}`);
  } catch (error) {
    failures.push(`JSDoc: ${error.message}`);
  }

  try {
    diagrams(wrap(source), name);
    console.log(`[PASS] Diagrams: ${relative}`);
  } catch (error) {
    failures.push(`Diagrams: ${error.message}`);
  }

  for (const error of failures) {
    errors.push(`${relative}: ${error}`);
    console.error(`[FAIL] ${relative}: ${error}`);
  }

  return `<section><h2>${esc(relative)}</h2>
<p>${esc(summary)}</p>
<h3>入出力タグ</h3>${tagHtml}
<h3>制御条件</h3>${conditionHtml}
<h3>解析成果物</h3>
<p><a href="flows/${name}.flow.html">処理フロー</a> |
<a href="flows/${name}.calls.html">関数依存図</a> |
<a href="../coverage/${name}/index.html">カバレッジ</a></p>
</section>`;
}

if (!fs.existsSync(SOURCE)) throw new Error("targets directory not found");
const files = discover(SOURCE);
if (!files.length) throw new Error("No FUXA targets found");

fs.mkdirSync(FLOWS, { recursive: true });
fs.mkdirSync(JSDOC, { recursive: true });

const sections = [];
const errors = [];

for (const file of files) {
  try {
    sections.push(build(file));
  } catch (error) {
    errors.push(`${file}: ${error.message}`);
    console.error(`[FAIL] ${file}: ${error.message}`);
  }
}

const page = `<!doctype html><html lang="ja"><head>
<meta charset="utf-8"><title>FUXA Script Design</title>
<style>
body{font:15px system-ui,sans-serif;max-width:1100px;margin:35px auto;
padding:20px;background:#f5f7fb;color:#202b3d}
section{background:white;padding:25px;margin:20px 0;
border:1px solid #dce3ed;border-radius:10px}
h2{border-bottom:2px solid #3975ce;padding-bottom:8px}
table{border-collapse:collapse;width:100%;margin:12px 0}
td,th{border:1px solid #dce3ed;padding:10px;text-align:left}
th{background:#edf2f8}code{overflow-wrap:anywhere}a{color:#2466c0}
</style></head><body>
<h1>FUXA Server Script 設計書</h1>
<p><a href="api/index.html">JSDoc 仕様書</a></p>
${sections.join("\n")}
${errors.length ? `<h2>生成エラー</h2><ul>${errors.map(
  e => `<li>${esc(e)}</li>`
).join("")}</ul>` : ""}
</body></html>`;

fs.writeFileSync(path.join(DOCS, "index.html"), page, "utf8");
console.log(`Processed: ${files.length}, Failed: ${errors.length}`);
process.exitCode = errors.length ? 1 : 0;