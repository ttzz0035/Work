
import fs from "node:fs";
import { spawnSync } from "node:child_process";

// npm run qualityで起動したnpm本体を使用する。
const npmCli = process.env.npm_execpath;

if (!npmCli || !fs.existsSync(npmCli)) {
  throw new Error("npm run qualityから実行してください。");
}

// テスト関連Docsは別コマンドで生成する。
// qualityでは品質検査とFUXA Docsのみを実行する。
const checks = [
  { name: "ESLint 静的解析", script: "lint" },
  { name: "Vitest C0/C1/C2", script: "test" },
  { name: "FUXA 設計書・フロー", script: "docs:flow" },
  { name: "FUXA JSDoc", script: "docs:api" },
];

const results = [];

// 各工程が失敗しても後続の実行を継続する。
for (const check of checks) {
  const result = spawnSync(process.execPath, [npmCli, "run", check.script], {
    encoding: "utf8",
    shell: false,
    windowsHide: true,
    maxBuffer: 50 * 1024 * 1024,
  });

  const passed = !result.error && result.status === 0;
  const output = [result.stdout, result.stderr, result.error?.message].filter(Boolean).join("\n");

  results.push({ name: check.name, passed, output });
}

// 最後に全工程の結果を集計する。
console.log("\n========== QUALITY SUMMARY ==========\n");

for (const result of results) {
  console.log(`[${result.passed ? "PASS" : "FAIL"}] ${result.name}`);
}

const passed = results.filter((result) => result.passed).length;
const failed = results.length - passed;

console.log(`\nPassed: ${passed}`);
console.log(`Failed: ${failed}`);
console.log(`Total : ${results.length}`);

// 失敗した工程のログだけを最後に表示する。
if (failed > 0) {
  console.log("\n========== FAILED DETAILS ==========");

  for (const result of results.filter((result) => !result.passed)) {
    console.log(`\n========== ${result.name} ==========\n`);
    console.log(result.output || "実行結果を取得できませんでした。");
  }
}

process.exitCode = failed > 0 ? 1 : 0;
