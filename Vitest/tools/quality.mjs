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
  console.log(`\n========== ${check.name} ==========\n`);

  const result = spawnSync(process.execPath, [npmCli, "run", check.script], {
    stdio: "inherit",
    shell: false,
    windowsHide: true,
  });

  const passed = !result.error && result.status === 0;
  results.push({ name: check.name, passed });

  if (result.error) console.error(`[ERROR] ${result.error.message}`);
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

process.exitCode = failed > 0 ? 1 : 0;