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
const DOCS = path.join(ROOT, "docs", "test");
const FLOWS = path.join(DOCS, "flows");

// HTMLに表示するソース由来の文字列をエスケープする。
function esc(value) {
  return String(value).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}

// テスト基盤に含まれるJavaScriptファイルを再帰的に収集する。
function discover(dir) {
  if (!fs.existsSync(dir)) throw new Error(`Directory not found: ${dir}`);
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) return discover(file);
    return entry.isFile() && /\.(js|mjs|cjs)$/.test(entry.name) ? [file] : [];
  }).sort();
}

// JSDocまたは先頭コメントから概要を取得する。
function description(source, name) {
  const match = source.match(/@description\s+([^\r\n*]+)/);
  if (match) return match[1].trim();
  const comment = source.match(/^\s*\/\/\s*(.+)/);
  return comment ? comment[1] : name;
}

// Babelで構造情報を収集する。ソースコードは変更しない。
function analyze(source) {
  const ast = parse(source, { sourceType: "unambiguous" });
  const classes = [];
  const functions = [];
  const conditions = [];
  const tests = [];

  // 条件式の種類とソース行番号を記録する。
  const addCondition = (p) => {
    if (!p.node.test) return;
    conditions.push({
      line: p.node.loc.start.line,
      kind: p.node.type,
      expression: generate(p.node.test).code,
    });
  };

  traverse(ast, {
    ClassDeclaration(p) {
      classes.push({
        line: p.node.loc.start.line,
        name: p.node.id?.name ?? "(anonymous)",
      });
    },
    FunctionDeclaration(p) {
      functions.push({
        line: p.node.loc.start.line,
        name: p.node.id?.name ?? "(anonymous)",
      });
    },
    ClassMethod(p) {
      const owner = p.findParent((x) => x.isClassDeclaration());
      const className = owner?.node.id?.name ?? "Class";
      functions.push({
        line: p.node.loc.start.line,
        name: `${className}.${generate(p.node.key).code}`,
      });
    },
    ObjectMethod(p) {
      functions.push({
        line: p.node.loc.start.line,
        name: generate(p.node.key).code,
      });
    },
    VariableDeclarator(p) {
      const node = p.node.init;
      if (!node || !["ArrowFunctionExpression", "FunctionExpression"].includes(node.type)) return;
      functions.push({
        line: p.node.loc.start.line,
        name: generate(p.node.id).code,
      });
    },
    CallExpression(p) {
      const callee = p.node.callee;
      const name = callee.type === "Identifier"
        ? callee.name
        : callee.type === "MemberExpression" ? generate(callee).code : "";

      // Vitestのテスト定義を静的に収集する。
      if (!/^(test|it|describe)(\.(skip|only|todo|each))*$/.test(name)) return;
      const arg = p.node.arguments[0];

      tests.push({
        line: p.node.loc.start.line,
        type: name,
        name: arg?.type === "StringLiteral"
          ? arg.value : arg ? generate(arg).code : "(unnamed)",
      });
    },
    IfStatement: addCondition,
    ConditionalExpression: addCondition,
    WhileStatement: addCondition,
    DoWhileStatement: addCondition,
    ForStatement: addCondition,
    SwitchStatement(p) {
      conditions.push({
        line: p.node.loc.start.line,
        kind: "SwitchStatement",
        expression: generate(p.node.discriminant).code,
      });
    },
  });

  return { classes, functions, conditions, tests };
}

// 一覧表を共通の表示形式で生成する。
function table(headers, rows) {
  const head = headers.map((h) => `<th>${esc(h)}</th>`).join("");
  const body = rows.map((row) =>
    `<tr>${row.map((cell) => `<td>${esc(cell)}</td>`).join("")}</tr>`
  ).join("");

  return `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

// 元ファイル全体をjs2flowchartに渡す。関数の再構成は行わない。
function diagrams(source, name) {
  const links = [];
  const errors = [];

  try {
    const svg = flow.convertCodeToSvg(source);
    fs.writeFileSync(path.join(FLOWS, `${name}.flow.svg`), svg);
    links.push(`<a href="flows/${name}.flow.svg">処理フロー</a>`);
  } catch (error) {
    errors.push(`処理フロー: ${error.message}`);
  }

  try {
    const builder = flow.createFlowTreeBuilder();
    builder.setAbstractionLevel(flow.ABSTRACTION_LEVELS.FUNCTION_DEPENDENCIES);
    const svg = flow.convertFlowTreeToSvg(builder.build(source));
    fs.writeFileSync(path.join(FLOWS, `${name}.calls.svg`), svg);
    links.push(`<a href="flows/${name}.calls.svg">関数依存図</a>`);
  } catch (error) {
    errors.push(`関数依存図: ${error.message}`);
  }

  return { links, errors };
}

// ファイル単位で解析し、FUXA版と同じ形式のHTMLを生成する。
function build(file) {
  const relative = path.relative(ROOT, file).replace(/\\/g, "/");
  const name = relative.replace(/\.(js|mjs|cjs)$/, "").replace(/\//g, "_");
  const source = fs.readFileSync(file, "utf8");
  const info = analyze(source);
  const chart = diagrams(source, name);
  const status = chart.errors.length
    ? `<p>図生成非対応: ${esc(chart.errors.join(" / "))}</p>` : "";

  const html = `<section>
<h2>${esc(relative)}</h2>
<p>${esc(description(source, relative))}</p>
<h3>クラス一覧</h3>
${table(["行", "クラス"], info.classes.map((c) => [c.line, c.name]))}
<h3>関数・メソッド一覧</h3>
${table(["行", "関数"], info.functions.map((f) => [f.line, f.name]))}
<h3>Vitest テストケース定義</h3>
${table(["行", "種類", "テスト名"], info.tests.map((t) => [t.line, t.type, t.name]))}
<h3>制御条件</h3>
${table(["行", "構文", "条件式"], info.conditions.map((c) =>
  [c.line, c.kind, c.expression]))}
<h3>解析成果物</h3>
<p>${chart.links.join(" | ") || "対応する図はありません"}</p>
${status}
</section>`;

  console.log(`[${chart.errors.length ? "PARTIAL" : "PASS"}] ${relative}`);
  chart.errors.forEach((e) => console.warn(`  ${e}`));
  return html;
}

// テスト基盤のドキュメントを独立したディレクトリへ生成する。
fs.mkdirSync(FLOWS, { recursive: true });

const files = [
  path.join(ROOT, "framework.cjs"),
  ...discover(path.join(ROOT, "tests")),
  ...discover(path.join(ROOT, "tools")),
];

const sections = [];
let failed = 0;

for (const file of files) {
  try {
    sections.push(build(file));
  } catch (error) {
    failed++;
    console.error(`[FAIL] ${file}: ${error.message}`);
  }
}

// JSDocとBabel解析結果を別々に保持して参照する。
const html = `<!doctype html><html lang="ja"><head>
<meta charset="utf-8"><title>FUXA Test Framework Design</title>
<style>
body{font:15px system-ui,sans-serif;max-width:1100px;margin:35px auto;padding:20px;background:#f5f7fb;color:#202b3d}
section{background:white;padding:25px;margin:20px 0;border:1px solid #dce3ed;border-radius:10px}
h2{border-bottom:2px solid #3975ce;padding-bottom:8px}
table{border-collapse:collapse;width:100%;margin:12px 0}
td,th{border:1px solid #dce3ed;padding:10px;text-align:left;overflow-wrap:anywhere}
th{background:#edf2f8}a{color:#2466c0}
</style></head><body>
<h1>FUXA Test Framework 設計書</h1>
<p><a href="api/index.html">JSDoc 仕様書</a></p>
${sections.join("\n")}
</body></html>`;

fs.writeFileSync(path.join(DOCS, "index.html"), html);
console.log(`Files: ${files.length}, Analysis failures: ${failed}`);
process.exitCode = failed ? 1 : 0;