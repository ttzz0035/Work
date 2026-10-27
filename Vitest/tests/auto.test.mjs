
import fs from "node:fs";
import path from "node:path";
import fc from "fast-check";
import { describe, it } from "vitest";
import { parse } from "@babel/parser";
import traverseModule from "@babel/traverse";
import generateModule from "@babel/generator";
import framework from "../framework.cjs";

const traverse = traverseModule.default ?? traverseModule;
const generate = generateModule.default ?? generateModule;
const { FuxaDriver } = framework;
const SCRIPT_DIR = path.resolve("targets");
const MAX_COMBINATIONS = 1024;
const SEED = 12345;

// JSファイルを再帰的に検出する。
function discover(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(entry => {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) return discover(file);
    return entry.isFile() && file.endsWith(".js") ? [file] : [];
  }).sort();
}

// ASTから負数を含む定数を取得する。
function constantOf(node) {
  if (node?.type === "NumericLiteral") return node.value;
  if (node?.type === "StringLiteral") return node.value;
  if (node?.type === "BooleanLiteral") return node.value;
  if (node?.type === "NullLiteral") return null;
  if (node?.type === "UnaryExpression" && node.operator === "-") {
    const value = constantOf(node.argument);
    return typeof value === "number" ? -value : undefined;
  }
  return undefined;
}

// 比較条件を安全に評価する。
function compare(a, op, b) {
  switch (op) {
    case "<": return a < b;
    case "<=": return a <= b;
    case ">": return a > b;
    case ">=": return a >= b;
    case "==": return a == b;
    case "===": return a === b;
    case "!=": return a != b;
    case "!==": return a !== b;
    default: throw new Error(`Unsupported operator: ${op}`);
  }
}

// await式を取り除き、元の式のPathを返す。
function unwrapAwait(p) {
  if (!p?.node) return null;
  return p.isAwaitExpression() ? p.get("argument") : p;
}

// BabelのPathから$getTagの参照名を静的に解析する。
// 動的にしか決定できない参照は、呼び出し元で未解決として扱う。
function resolveTagCall(p, tags, unresolved, filename) {
  const expression = unwrapAwait(p);
  if (!expression?.isCallExpression()) return null;
  if (!expression.get("callee").isIdentifier({ name: "$getTag" })) return null;

  const argument = expression.get("arguments")[0];

  if (!argument?.node) {
    unresolved.set(`missing:${expression.node.start}`, {
      line: Math.max(1, (expression.node.loc?.start.line ?? 2) - 1),
      expression: "$getTag()",
      reason: "Missing $getTag argument",
    });
    return null;
  }

  // Babelのスコープ情報を使って、変数・文字列結合を評価する。
  const evaluated = argument.evaluate();

  if (evaluated.confident && typeof evaluated.value === "string") {
    const name = evaluated.value;
    if (!tags.has(name)) tags.set(name, new Set());
    return name;
  }

  // 解決できない場合も、解析全体を例外で停止させない。
  const expressionCode = generate(argument.node).code;

  unresolved.set(`dynamic:${expression.node.start}`, {
    line: Math.max(1, (expression.node.loc?.start.line ?? 2) - 1),
    expression: expressionCode,
    reason: `INPUT_SCHEMA_REQUIRED: ${filename}`,
  });

  return null;
}

// 比較式の左辺・右辺が参照するFUXAタグを解決する。
function resolveTagReference(p, tags, variables, unresolved, filename) {
  const expression = unwrapAwait(p);
  if (!expression?.node) return null;

  if (expression.isIdentifier()) {
    return variables.get(expression.node.name) ?? null;
  }

  return resolveTagCall(expression, tags, unresolved, filename);
}

// 境界値を対象タグへ登録する。
function registerBoundaryValues(values, value) {
  if (typeof value === "number") {
    [value - 1, value, value + 1].forEach(v => values.add(v));
  } else if (typeof value === "string") {
    values.add(value);
    values.add(`${value}__other__`);
  } else if (typeof value === "boolean") {
    values.add(true);
    values.add(false);
  } else {
    values.add(null);
    values.add(0);
    values.add(1);
  }
}

// FUXAタグ、条件式、Axiosの使用を解析する。
function analyze(filename) {
  const source = fs.readFileSync(filename, "utf8");
  const ast = parse(`async function __test__(){\n${source}\n}`, {
    sourceType: "script",
  });
  const tags = new Map();
  const variables = new Map();
  const predicates = [];
  const fields = new Set();
  const unresolved = new Map();
  let usesAxios = false;

  // タグとHTTPレスポンスの参照を収集する。
  traverse(ast, {
    CallExpression(p) {
      resolveTagCall(p, tags, unresolved, filename);
      const c = p.node.callee;
      if (c.type === "MemberExpression" &&
          c.object.type === "Identifier" &&  c.object.name === "axios") {
        usesAxios = true;
      }

      // require("axios")の場合もHTTPシナリオを生成する。
      if (c.type === "Identifier" &&
          c.name === "require" &&
          p.node.arguments[0]?.type === "StringLiteral" &&
          p.node.arguments[0].value === "axios") {
        usesAxios = true;
      }
    },
    VariableDeclarator(p) {
      if (!p.get("id").isIdentifier()) return;
      const name = resolveTagCall(p.get("init"), tags, unresolved, filename);
      if (name !== null) variables.set(p.node.id.name, name);
    },
    MemberExpression(p) {
      const n = p.node;
      if (n.computed || n.property.type !== "Identifier") return;
      const parent = n.object;
      if (parent?.type === "MemberExpression" &&
          parent.property.type === "Identifier" &&
          parent.property.name === "data") {fields.add(n.property.name);}
    },
  });

  // 比較条件と、真偽を切り替えられる境界値を抽出する。
  traverse(ast, {
    BinaryExpression(p) {
      const n = p.node;
      if (!["<", "<=", ">", ">=", "==", "===", "!=", "!=="].includes(n.operator)) return;
      const left = resolveTagReference(p.get("left"), tags, variables, unresolved, filename);
      const right = resolveTagReference(p.get("right"), tags, variables, unresolved, filename);
      if (left === null && right === null) return;
      if (left !== null && right !== null) {
        throw new Error(`Tag-to-tag comparison unsupported: ${filename}`);
      }

      const tag = left ?? right;
      const value = constantOf(left !== null ? n.right : n.left);
      if (value === undefined) {throw new Error(`Unsupported comparison: ${filename}`);}

      // 条件を評価する関数を保存する。左右の順序も維持する。
      predicates.push(input => left !== null
        ? compare(input[tag], n.operator, value)
        : compare(value, n.operator, input[tag]));

      registerBoundaryValues(tags.get(tag), value);
    },
  });

  // 境界値の根拠がないタグも未対応として記録する。
  // カバレッジ取得自体は継続できるよう、基本入力を用意する。
  for (const [tag, values] of tags) {
    if (values.size) continue;

    unresolved.set(`constraints:${tag}`, {
      line: null,
      expression: tag,
      reason: "INPUT_SCHEMA_REQUIRED: No input constraints",
    });

    [false, true, 0, 1, -1, "", "OK", null].forEach(v => values.add(v));
  }
  return {
    tags,
    predicates,
    fields,
    usesAxios,
    unresolved: [...unresolved.values()],
  };
}

// 全候補を生成し、条件の真偽パターンで優先順位を決める。
function generateCases(tags, predicates) {
  const entries = [...tags].map(([name, values]) => [name, [...values]]);
  const total = entries.reduce((n, [, values]) => n * values.length, 1);
  if (total > MAX_COMBINATIONS) {
    throw new Error(`Input combinations ${total} exceed ${MAX_COMBINATIONS}`);
  }
  if (!entries.length) return [{}];

  // fast-checkによる再現可能な探索候補を作る。
  const arbitrary = fc.record(Object.fromEntries(
    entries.map(([name, values]) => [name, fc.constantFrom(...values)])
  ));
  const sampled = fc.sample(arbitrary, { seed: SEED, numRuns: 32 });

  // 全境界値の組合せも保持し、探索漏れを防ぐ。
  let exhaustive = [{}];
  for (const [name, values] of entries) {
    exhaustive = exhaustive.flatMap(row =>
      values.map(value => ({ ...row, [name]: value }))
    );
  }

  const unique = new Map();
  for (const input of [...sampled, ...exhaustive]) {
    unique.set(JSON.stringify(input), input);
  }

  // 条件のTrue/Falseを特徴量に変換する。
  const remaining = [...unique.values()].map(input => ({
    input,
    bits: predicates.map(predicate => Number(predicate(input))),
  }));
  const ordered = [];
  const covered = new Set();

  // 新しい条件状態と条件ペアを多く含む入力から選択する。
  while (remaining.length) {
    let best = 0;
    let bestScore = -1;

    for (let i = 0; i < remaining.length; i++) {
      const bits = remaining[i].bits;
      let score = 0;

      for (let a = 0; a < bits.length; a++) {
        if (!covered.has(`${a}:${bits[a]}`)) score += 100;
        for (let b = a + 1; b < bits.length; b++) {
          if (!covered.has(`${a}:${b}:${bits[a]}:${bits[b]}`)) score++;
        }
      }
      if (score > bestScore) {
        bestScore = score;
        best = i;
      }
    }

    const selected = remaining.splice(best, 1)[0];
    ordered.push(selected.input);
    selected.bits.forEach((bit, a) => covered.add(`${a}:${bit}`));
    for (let a = 0; a < selected.bits.length; a++) {
      for (let b = a + 1; b < selected.bits.length; b++) {
        covered.add(`${a}:${b}:${selected.bits[a]}:${selected.bits[b]}`);
      }
    }
  }
  return ordered;
}

// HTTPは通信せず、Axiosの戻り値・例外のみ模擬する。
function generateHttpCases(fields, enabled) {
  if (!enabled) return [null];
  const data = v => Object.fromEntries([...fields].map(name => [name, v]));
  return [
    ...[true, false, 0, 1, "OK"].map(v => ({ status: 200, data: data(v) })),
    { status: 201, data: data(true) },
    { status: 400, data: {} },
    { status: 500, data: {} },
    mock => mock.onAny().networkError(),
    mock => mock.onAny().timeout(),
  ];
}

// 全JSを独立したカバレッジテストへ登録する。
function registerTest(filename) {
  const name = path.relative(SCRIPT_DIR, filename);
  const reportName = name.replace(/[\\/]/g, "_").replace(/\.js$/, "");

  describe(`FUXA WhiteBox: ${name}`, () => {
    it("C0/C1/C2 coverage", async () => {
      const errors = [];
      let inputs = [];
      let httpCases = [];
      let driver = null;
      let executed = 0;

      // 解析と計測準備を分離し、一方が失敗しても実行可能な工程を試す。
      try {
        const { tags, predicates, fields, usesAxios, unresolved } = analyze(filename);
        inputs = generateCases(tags, predicates);
        httpCases = generateHttpCases(fields, usesAxios);

        // 未解決タグは報告するが、実行可能なケースを捨てない。
        for (const item of unresolved) {
          const location = item.line === null ? "" : `Line ${item.line}: `;
          errors.push(`${location}${item.reason}: ${item.expression}`);
        }
      } catch (error) {
        errors.push(`Analyze: ${error.message}`);
      }

      try {
        driver = new FuxaDriver(filename);
      } catch (error) {
        errors.push(`Instrument: ${error.message}`);
      }

      if (driver && inputs.length && httpCases.length) {
        // 全候補を試行し、網羅率達成または全ケース終了まで継続する。
        let complete = false;

        for (const input of inputs) {
          for (const httpCase of httpCases) {
            executed++;

            try {
              await driver.run(input, [], httpCase);
            } catch (error) {
              errors.push(`Case ${executed}: ${error.message}`);
            }

            try {
              const result = driver.summary();
              complete = result.c0 === 100 && result.c1 === 100 && result.c2 === 100;
            } catch (error) {
              errors.push(`Summary ${executed}: ${error.message}`);
            }

            // 過去のケースで失敗があれば、残りのケースも続行する。
            if (complete && errors.length === 0) break;
          }
          if (complete && errors.length === 0) break;
        }
      }

      console.log(`\n=== ${name} ===`);
      console.log(`Generated tag cases: ${inputs.length}`);
      console.log(`HTTP scenarios: ${httpCases.length}`);
      console.log(`Executed cases: ${executed}`);

      // 実行途中でエラーが発生しても、取得済みの計測結果を出力する。
      if (driver) {
        try {
          driver.report(`coverage/${reportName}`, 100);
        } catch (error) {
          errors.push(`Report: ${error.message}`);
        }
      }

      // すべてのレポート出力後に品質ゲートを判定する。
      if (errors.length) {
        for (const message of errors) console.error(`[FAIL] ${name}: ${message}`);
        throw new AggregateError(
          errors.map(message => new Error(message)),
          `${name}: ${errors.length} error(s)\n${errors.join("\n")}`
        );
      }
    }, 60000);
  });
}

if (!fs.existsSync(SCRIPT_DIR)) {throw new Error(`targets directory not found: ${SCRIPT_DIR}`);}
const files = discover(SCRIPT_DIR);
if (!files.length) {throw new Error("No FUXA JavaScript files found");}

for (const file of files) registerTest(file);