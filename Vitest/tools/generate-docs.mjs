import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { parse } from "@babel/parser";
import traverseModule from "@babel/traverse";
import generateModule from "@babel/generator";

const require = createRequire(import.meta.url);
const flow = require("js2flowchart");
const traverse = traverseModule.default ?? traverseModule;
const generate = generateModule.default ?? generateModule;
const ROOT = process.cwd();
const SOURCE = path.join(ROOT, "targets");
const DOCS = path.join(ROOT, "docs");
const FLOWS = path.join(DOCS, "flows");
const JSDOC = path.join(DOCS, "jsdoc-src");

// HTMLに埋め込むソース由来の文字列をエスケープする。
function esc(value) {
  return String(value).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}

// FUXAのJSファイルを再帰的に収集する。
function discover(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) return discover(file);
    return entry.isFile() && entry.name.endsWith(".js") ? [file] : [];
  }).sort();
}

// FUXAの関数本体を、解析用のasync関数にする。
function wrap(source) {
  return `async function fuxaScript() {\n${source}\n}`;
}

// BabelでFUXA固有のタグと条件を抽出する。
function analyze(source) {
  const ast = parse(wrap(source), { sourceType: "script" });
  const inputs = new Set();
  const outputs = new Set();
  const conditions = [];

  traverse(ast, {
    CallExpression(p) {
      if (p.node.callee.type !== "Identifier") return;

      const name = p.node.callee.name;
      if (name !== "$getTag" && name !== "$setTag") return;

      // 動的タグ名も式のまま記録し、一覧から欠落させない。
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

// コメントから概要を読み取る。説明がなければファイル名を使う。
function description(source, name) {
  const match = source.match(/@description\s+([^\r\n*]+)/);
  if (match) return match[1].trim();
  const comment = source.match(/^\s*\/\/\s*(.+)/);
  return comment ? comment[1] : name;
}

// 既存ツールに処理フローと関数依存構造を生成させる。
function diagrams(code, name) {
  fs.writeFileSync(path.join(FLOWS, `${name}.flow.svg`), flow.convertCodeToSvg(code));

  const builder = flow.createFlowTreeBuilder();
  builder.setAbstractionLevel(flow.ABSTRACTION_LEVELS.FUNCTION_DEPENDENCIES);
  fs.writeFileSync(path.join(FLOWS, `${name}.calls.svg`), flow.convertFlowTreeToSvg(builder.build(code)));
}

// 入出力タグのHTML表を生成する。
function tagTable(inputs, outputs) {
  const rows = [
    ...[...inputs].map((v) => ["入力", v]),
    ...[...outputs].map((v) => ["出力", v]),
  ];
  return `<table><tr><th>方向</th><th>タグ名</th></tr>${rows.map(
    ([direction, name]) => `<tr><td>${esc(direction)}</td><td><code>${esc(name)}</code></td></tr>`
  ).join("")}</table>`;
}

// 制御条件と元ソースの行番号を一覧化する。
function conditionTable(conditions) {
  return `<table><tr><th>行</th><th>構文</th><th>条件式</th></tr>${conditions.map(
    (c) => `<tr><td>${c.line}</td><td>${esc(c.kind)}</td><td><code>${esc(c.expression)}</code></td></tr>`
  ).join("")}</table>`;
}

// JSDocにはモジュール定義を追加した解析専用ソースを渡す。
function jsdocSource(source, name, summary) {
  const clean = summary.replace(/\*\//g, "* /");
  return `/**\n * @module ${name}\n * @description ${clean}\n */\n${wrap(source)}\n`;
}

// HTML設計書とツール出力をファイル単位で生成する。
function build(file) {
  const relative = path.relative(SOURCE, file);
  const name = relative.replace(/[\\/]/g, "_").replace(/\.js$/, "");
  const source = fs.readFileSync(file, "utf8");
  const summary = description(source, name);
  const info = analyze(source);
  const code = wrap(source);

  fs.writeFileSync(path.join(JSDOC, `${name}.js`), jsdocSource(source, name, summary));
  diagrams(code, name);

  return `<section><h2>${esc(relative)}</h2>
<p>${esc(summary)}</p>
<h3>入出力タグ</h3>${tagTable(info.inputs, info.outputs)}
<h3>制御条件</h3>${conditionTable(info.conditions)}
<h3>解析成果物</h3>
<p><a href="flows/${name}.flow.svg">処理フロー</a> |
<a href="flows/${name}.calls.svg">関数依存図</a> |
<a href="../coverage/${name}/index.html">カバレッジ</a></p>
</section>`;
}

// 個々の生成エラーを集計し、他ファイルの処理は継続する。
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
    console.log(`[PASS] ${path.relative(SOURCE, file)}`);
  } catch (error) {
    errors.push(file);
    console.error(`[FAIL] ${file}: ${error.message}`);
  }
}

// 統合HTMLはツールの成果物にリンクするだけにする。
const html = `<!doctype html><html lang="ja"><head>
<meta charset="utf-8"><title>FUXA Script Design</title>
<style>
body{font:15px system-ui,sans-serif;max-width:1100px;margin:35px auto;padding:20px;background:#f5f7fb;color:#202b3d}
section{background:white;padding:25px;margin:20px 0;border:1px solid #dce3ed;border-radius:10px}
h2{border-bottom:2px solid #3975ce;padding-bottom:8px}
table{border-collapse:collapse;width:100%;margin:12px 0}
td,th{border:1px solid #dce3ed;padding:10px;text-align:left}
th{background:#edf2f8}code{overflow-wrap:anywhere}a{color:#2466c0}
</style></head><body><h1>FUXA Server Script 設計書</h1>
<p><a href="api/index.html">JSDoc 仕様書</a></p>
${sections.join("\n")}
</body></html>`;

fs.writeFileSync(path.join(DOCS, "index.html"), html);
console.log(`Processed: ${files.length}, Failed: ${errors.length}`);
process.exitCode = errors.length ? 1 : 0;