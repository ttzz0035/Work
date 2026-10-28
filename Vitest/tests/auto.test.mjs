
import fs from "node:fs";
import path from "node:path";
import fc from "fast-check";
import { describe, it } from "vitest";
import { parse } from "@babel/parser";
import traverseModule from "@babel/traverse";
import framework from "../framework.cjs";

const traverse = traverseModule.default ?? traverseModule;
const { FuxaDriver } = framework;
const SCRIPT_DIR = path.resolve("targets");
const MAX_COMBINATIONS = 2048;
const SEED = 12345;
const MAX_RUNTIME_ERRORS = 10;

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

// 入力候補の重複を除去する。
function addCandidate(values, value) {
  if (typeof value === "number" && !Number.isFinite(value)) return;
  const key = JSON.stringify(value);
  if (!values.some(existing => JSON.stringify(existing) === key)) values.push(value);
}

// 入力元を解析する。条件式自体はここでは評価しない。
function analyze(filename) {
  const source = fs.readFileSync(filename, "utf8");
  const ast = parse(`async function __test__(){\n${source}\n}`, { sourceType: "script" });
  const tags = new Map();
  const literals = [];
  const fields = new Set();
  const unresolved = new Set();
  let usesAxios = false;

  // FUXAタグの静的な参照名を取得する。
  traverse(ast, {
    CallExpression(p) {
      const callee = p.get("callee");
      if (callee.isIdentifier({ name: "$getTag" })) {
        const argument = p.get("arguments")[0];
        if (!argument?.node) {
          unresolved.add(`Missing $getTag argument: ${filename}`);
        } else {
          const evaluated = argument.evaluate();
          if (evaluated.confident && typeof evaluated.value === "string") {
            if (!tags.has(evaluated.value)) tags.set(evaluated.value, []);
          } else {
            unresolved.add(`Dynamic $getTag: ${filename}:${Math.max(1, (p.node.loc?.start.line ?? 2) - 1)}`);
          }
        }
      }

      // HTTP呼び出しを判別する。
      const node = p.node.callee;
      if (node.type === "MemberExpression" && node.object.type === "Identifier" && node.object.name === "axios") {
        usesAxios = true;
      }
      if (callee.isIdentifier({ name: "require" }) && p.node.arguments[0]?.value === "axios") {
        usesAxios = true;
      }
    },

    // 条件式に現れる比較値を収集する。
    // 関数名・演算子ごとの入力生成規則は定義しない。
    NumericLiteral(p) {
      addCandidate(literals, p.node.value);
    },
    StringLiteral(p) {
      addCandidate(literals, p.node.value);
    },
    BooleanLiteral(p) {
      addCandidate(literals, p.node.value);
    },
    UnaryExpression(p) {
      const value = constantOf(p.node);
      if (value !== undefined) addCandidate(literals, value);
    },
    MemberExpression(p) {
      const node = p.node;
      if (node.computed || node.property.type !== "Identifier") return;
      const parent = node.object;
      if (parent?.type === "MemberExpression" && parent.property.type === "Identifier" && parent.property.name === "data") {
        fields.add(node.property.name);
      }
    },
  });

  return { tags, literals, fields, usesAxios, unresolved: [...unresolved] };
}

// 比較値の近傍と汎用値から入力候補を生成する。
function generateValues(literals) {
  const values = [];
  const defaults = [false, true, null, -1, 0, 1, -0.5, 0.5, 1.5, "", "OK", "0", "1", [], [0], [1], {}, { value: 0 }];

  for (const value of defaults) addCandidate(values, value);

  // 条件式に現れた定数の境界付近を探索する。
  for (const value of literals) {
    addCandidate(values, value);
    if (typeof value === "number") {
      addCandidate(values, value - 1);
      addCandidate(values, value + 1);
      addCandidate(values, value - 0.5);
      addCandidate(values, value + 0.5);
    }
    if (typeof value === "string") addCandidate(values, `${value}__other__`);
  }

  return values;
}

// 候補の全組合せ、またはfast-checkによるサンプリングを生成する。
function generateCases(tags, literals) {
  const values = generateValues(literals);
  const names = [...tags.keys()];
  if (!names.length) return [{}];

  const unique = new Map();
  const add = input => {
    if (unique.size < MAX_COMBINATIONS) unique.set(JSON.stringify(input), input);
  };

  // 変数同士の比較を通すため、各入力に異なる値を割り当てる。
  for (const value of values) {
    add(Object.fromEntries(names.map(name => [name, value])));
  }

  for (const name of names) {
    for (const value of values) {
      const input = Object.fromEntries(names.map(key => [key, 0]));
      input[name] = value;
      add(input);
    }
  }

  // 入力ペアの組合せを優先して生成する。
  for (let i = 0; i < names.length; i++) {
    for (let j = i + 1; j < names.length; j++) {
      for (const left of values) {
        for (const right of values) {
          const input = Object.fromEntries(names.map(name => [name, 0]));
          input[names[i]] = left;
          input[names[j]] = right;
          add(input);
        }
      }
    }
  }

  // 全境界値の組合せも保持し、探索漏れを防ぐ。
  const total = values.length ** names.length;
  if (total <= MAX_COMBINATIONS) {
    let exhaustive = [{}];
    for (const name of names) {
      exhaustive = exhaustive.flatMap(row => values.map(value => ({ ...row, [name]: value })));
    }
    for (const input of exhaustive) add(input);
  }

  // fast-checkによる再現可能な探索候補を作る。
  const arbitrary = fc.record(Object.fromEntries(names.map(name => [name, fc.constantFrom(...values)])));
  const sampled = fc.sample(arbitrary, { seed: SEED, numRuns: MAX_COMBINATIONS });

  for (const input of sampled) add(input);

  return [...unique.values()];
}

// 実行済みの入力を変異させ、未探索の候補を生成する。
function mutate(input, names, values, random) {
  const result = structuredClone(input);
  if (!names.length) return result;

  const count = Math.max(1, Math.ceil(names.length * random()));
  for (let i = 0; i < count; i++) {
    const name = names[Math.floor(random() * names.length)];
    const current = result[name];

    // 数値の場合は、実行した値の近傍も探索する。
    if (typeof current === "number" && random() < 0.5) {
      const offsets = [-1, -0.5, 0.5, 1];
      result[name] = current + offsets[Math.floor(random() * offsets.length)];
    } else {
      result[name] = structuredClone(values[Math.floor(random() * values.length)]);
    }
  }

  return result;
}

// 現在のカバレッジから、実行済み箇所を識別する。
function coverageSignature(driver) {
  const file = driver.coverage.files().includes(driver.file)
    ? driver.coverage.fileCoverageFor(driver.file)
    : null;

  const statements = file ? Object.entries(file.s).filter(([, count]) => count > 0).map(([id]) => `S:${id}`) : [];
  const branches = file ? Object.entries(file.b).flatMap(([id, counts]) => counts.flatMap((count, index) => count > 0 ? [`B:${id}:${index}`] : [])) : [];
  const functions = file ? Object.entries(file.f).filter(([, count]) => count > 0).map(([id]) => `F:${id}`) : [];
  const conditions = [...driver.conditions.entries()].flatMap(([id, condition]) => [
    ...(condition.trueCount > 0 ? [`C:${id}:true`] : []),
    ...(condition.falseCount > 0 ? [`C:${id}:false`] : []),
  ]);

  return new Set([...statements, ...branches, ...functions, ...conditions]);
}

// 実行前後のカバレッジを比較し、新規に網羅された箇所を取得する。
function coverageGain(previous, current) {
  return [...current].filter(key => !previous.has(key)).length;
}

// HTTPは通信せず、Axiosの戻り値・例外のみ模擬する。
function generateHttpCases(fields, enabled) {
  if (!enabled) return [null];
  const data = value => Object.fromEntries([...fields].map(name => [name, value]));

  return [
    ...[true, false, 0, 1, "OK"].map(value => ({ status: 200, data: data(value) })),
    { status: 201, data: data(true) },
    { status: 400, data: {} },
    { status: 500, data: {} },
    mock => mock.onAny().networkError(),
    mock => mock.onAny().timeout(),
  ];
}

// 最終結果にテスト対象のソース位置を表示する。
function sourceLocation(filename, line) {
  const lines = fs.readFileSync(filename, "utf8").split(/\r?\n/);
  return `${path.relative(process.cwd(), filename)}:${line}\n  ${line} | ${lines[line - 1] ?? ""}`;
}

// 未網羅の文・分岐・基本条件をテスト対象コードから報告する。
function coverageErrors(driver, filename) {
  const errors = [];
  if (!driver.coverage.files().includes(driver.file)) return ["Coverage unavailable: no executed instrumented script"];

  const file = driver.coverage.fileCoverageFor(driver.file);

  // Istanbulは対象ソースをasync関数で包んだ位置情報を使用する。
  const lineOf = line => Math.max(1, line - 1);

  for (const [id, count] of Object.entries(file.s)) {
    if (count > 0) continue;
    const location = file.statementMap[id];
    errors.push(`[C0] ${sourceLocation(filename, lineOf(location.start.line))}`);
  }

  for (const [id, counts] of Object.entries(file.b)) {
    if (counts.every(count => count > 0)) continue;
    const branch = file.branchMap[id];
    errors.push(`[C1] ${sourceLocation(filename, lineOf(branch.line))}\n  Branch: ${counts.map((count, index) => `${index + 1}=${count}`).join(", ")}`);
  }

  for (const condition of driver.summary().conditions) {
    if (condition.covered) continue;
    errors.push(`[C2] ${sourceLocation(filename, condition.line)}\n  ${condition.expression}: True=${condition.trueCount}, False=${condition.falseCount}`);
  }

  return [...new Set(errors)];
}

// 全JSを独立したカバレッジテストへ登録する。
function registerTest(filename) {
  const name = path.relative(SCRIPT_DIR, filename);
  const reportName = name.replace(/[\\/]/g, "_").replace(/\.js$/, "");

  describe(`FUXA WhiteBox: ${name}`, () => {
    it("C0/C1/C2 coverage", async () => {
      const errors = [];
      const runtimeErrors = new Map();
      let inputs = [];
      let httpCases = [];
      let driver = null;
      let executed = 0;
      let info = null;

      // 解析と計測準備を分離し、一方が失敗しても実行可能な工程を試す。
      try {
        info = analyze(filename);
        inputs = generateCases(info.tags, info.literals);
        httpCases = generateHttpCases(info.fields, info.usesAxios);

        // 入力生成に必要な情報が不足している条件を記録する。
        for (const issue of info.unresolved) errors.push(`INPUT_SCHEMA_REQUIRED: ${issue}`);
      } catch (error) {
        errors.push(`Analyze: ${error.message}`);
      }

      try {
        driver = new FuxaDriver(filename);
      } catch (error) {
        errors.push(`Instrument: ${error.message}`);
      }

      if (driver && info && inputs.length && httpCases.length) {
        const names = [...info.tags.keys()];
        const values = generateValues(info.literals);
        const selected = [];
        const seen = new Set();
        const queue = [...inputs];
        let complete = false;

        // fast-checkと同じシードから再現可能な変異系列を作る。
        const randomValues = fc.sample(fc.double({ min: 0, max: 1, noNaN: true }), { seed: SEED, numRuns: MAX_COMBINATIONS * 16 });
        let randomIndex = 0;
        const random = () => randomValues[randomIndex++ % randomValues.length];

        // 実行済みケースの中からカバレッジを増やした入力を優先する。
        for (let index = 0; index < queue.length && executed < MAX_COMBINATIONS * 2; index++) {
          const input = queue[index];
          const key = JSON.stringify(input);
          if (seen.has(key)) continue;
          seen.add(key);

          for (const httpCase of httpCases) {
            if (executed >= MAX_COMBINATIONS * 2) break;
            executed++;
            const previous = coverageSignature(driver);

            try {
              await driver.run(input, [], httpCase);
            } catch (error) {
              const message = error.message;
              runtimeErrors.set(message, (runtimeErrors.get(message) ?? 0) + 1);
            }

            const current = coverageSignature(driver);
            const gain = coverageGain(previous, current);

            if (gain > 0) {
              selected.push({ input: structuredClone(input), gain });

              // 新しく網羅できた入力を起点に、追加候補を生成する。
              for (let i = 0; i < 8 && queue.length < MAX_COMBINATIONS * 4; i++) {
                queue.push(mutate(input, names, values, random));
              }
            }

            const result = driver.summary();
            complete = result.c0 === 100 && result.c1 === 100 && result.c2 === 100;
            if (complete) break;
          }
          if (complete) break;
        }

        // 未網羅が残っている場合、効果の高かった入力を優先して追加探索する。
        if (!complete && executed < MAX_COMBINATIONS * 2 && selected.length) {
          selected.sort((a, b) => b.gain - a.gain);

          for (const item of selected.slice(0, 16)) {
            for (let i = 0; i < 16 && executed < MAX_COMBINATIONS * 2; i++) {
              const input = mutate(item.input, names, values, random);
              const key = JSON.stringify(input);
              if (seen.has(key)) continue;
              seen.add(key);
              executed++;

              try {
                await driver.run(input, [], null);
              } catch (error) {
                const message = error.message;
                runtimeErrors.set(message, (runtimeErrors.get(message) ?? 0) + 1);
              }

              const result = driver.summary();
              if (result.c0 === 100 && result.c1 === 100 && result.c2 === 100) {
                complete = true;
                break;
              }
            }
            if (complete) break;
          }
        }
      }

      // 実行途中でエラーが発生しても、取得済みの計測結果を出力する。
      if (driver) {
        try {
          // HTML生成と品質判定を分離する。閾値判定は下で一度だけ行う。
          driver.report(`coverage/${reportName}`, 0);
        } catch (error) {
          errors.push(`Report: ${error.message}`);
        }

        // 未網羅条件に対応する対象ソースを表示する。
        try {
          errors.push(...coverageErrors(driver, filename));
        } catch (error) {
          errors.push(`Coverage detail: ${error.message}`);
        }
      }

      // 実行時エラーを種類別に集約し、同一エラーの大量表示を避ける。
      for (const [message, count] of [...runtimeErrors].slice(0, MAX_RUNTIME_ERRORS)) {
        errors.push(`[RUNTIME] ${message} (${count} cases)`);
      }

      if (runtimeErrors.size > MAX_RUNTIME_ERRORS) {
        errors.push(`[RUNTIME] Additional error types: ${runtimeErrors.size - MAX_RUNTIME_ERRORS}`);
      }

      // すべてのレポート出力後に品質ゲートを判定する。
      if (errors.length) {
        const result = driver?.summary();
        console.error(`\n========== [FAIL] ${name} ==========`);

        if (result) {
          console.error(`C0=${result.c0.toFixed(2)}% C1=${result.c1.toFixed(2)}% C2=${result.c2.toFixed(2)}%`);
        }

        for (const message of [...new Set(errors)]) console.error(message);

        // Vitestへは失敗1件だけを渡し、各行を個別エラーにしない。
        throw new Error(`Coverage failed: ${name}`);
      }
    }, 60000);
  });
}

if (!fs.existsSync(SCRIPT_DIR)) throw new Error(`targets directory not found: ${SCRIPT_DIR}`);
const files = discover(SCRIPT_DIR);
if (!files.length) throw new Error("No FUXA JavaScript files found");

for (const file of files) registerTest(file);
