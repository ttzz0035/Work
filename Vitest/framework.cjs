const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const axios = require("axios");
const AxiosMockAdapter = require("axios-mock-adapter");
const parser = require("@babel/parser");
const traverse = require("@babel/traverse").default;
const generate = require("@babel/generator").default;
const t = require("@babel/types");
const { createInstrumenter } = require("istanbul-lib-instrument");
const { createCoverageMap } = require("istanbul-lib-coverage");
const libReport = require("istanbul-lib-report");
const reports = require("istanbul-reports");

// ソース由来の文字列をHTMLで安全に表示する。
function esc(value) {
  return String(value).replace(/[&<>"']/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}

// FUXAのタグ読取・書込を実機接続なしで再現する。
class FuxaStub {
  constructor(tags = {}) {
    this.tags = new Map(Object.entries(tags));
    this.reads = [];
    this.writes = [];
  }

  getTag(id) {
    this.reads.push(id);
    if (!this.tags.has(id)) throw new Error(`Unknown tag: ${id}`);
    return this.tags.get(id);
  }

  setTag(id, value) {
    this.tags.set(id, value);
    this.writes.push({ id, value });
  }

  getTagId(name) { return name; }
  value(id) { return this.tags.get(id); }
  writesTo(id) { return this.writes.filter(e => e.id === id); }
}

// FUXAスクリプトの実行、計測、レポート生成を担当する。
class FuxaDriver {
  constructor(file, parameters = []) {
    this.file = path.resolve(file);
    this.parameters = parameters;
    this.coverage = createCoverageMap({});
    this.conditions = new Map();
    this.sequence = 0;
    this.prepare();
  }

  prepare() {
    const source = fs.readFileSync(this.file, "utf8");
    if (!this.parameters.every(n => /^[A-Za-z_$][\w$]*$/.test(n))) {
      throw new Error("Invalid parameter name");
    }

    // FUXAの関数本体を解析用async関数で包む。
    const wrapped = `async function __fuxaScript(${this.parameters.join(",")}) {\n${source}\n}`;
    const ast = parser.parse(wrapped, { sourceType: "script" });

    // C2計測できない構文は解析時に明示的に拒否する。
    traverse(ast, {
      SwitchStatement() { throw new Error("Unsupported C2 syntax: switch"); },
      LogicalExpression(p) {
        if (p.node.operator === "??") throw new Error("Unsupported C2 syntax: nullish coalescing");
      },
      OptionalMemberExpression() { throw new Error("Unsupported C2 syntax: optional chaining"); },
      OptionalCallExpression() { throw new Error("Unsupported C2 syntax: optional call"); },
    });

    // 複合条件を基本条件に分解する。短絡評価は維持する。
    const instrumentCondition = node => {
      if (t.isLogicalExpression(node)) {
        node.left = instrumentCondition(node.left);
        node.right = instrumentCondition(node.right);
        return node;
      }

      const id = String(this.sequence++);
      this.conditions.set(id, {
        expression: generate(node).code,
        line: Math.max(1, (node.loc?.start.line ?? 2) - 1),
        trueCount: 0,
        falseCount: 0,
      });
      return t.callExpression(t.identifier("__c2"), [t.stringLiteral(id), node]);
    };

    // C2の評価コードを条件式へ挿入する。
    traverse(ast, {
      IfStatement(p) { p.node.test = instrumentCondition(p.node.test); },
      ConditionalExpression(p) { p.node.test = instrumentCondition(p.node.test); },
      WhileStatement(p) { p.node.test = instrumentCondition(p.node.test); },
      DoWhileStatement(p) { p.node.test = instrumentCondition(p.node.test); },
      ForStatement(p) {
        if (p.node.test) p.node.test = instrumentCondition(p.node.test);
      },
    });

    // C0・C1はIstanbul標準の計測処理を利用する。
    this.code = createInstrumenter({
      coverageVariable: "__coverage__",
      compact: false,
      preserveComments: true,
    }).instrumentSync(generate(ast, { comments: true }).code, this.file);
  }

  async run(tags = {}, args = [], httpScenario = null) {
    const stub = new FuxaStub(tags);
    const http = axios.create();
    const mock = new AxiosMockAdapter(http, { onNoMatch: "throwException" });

    // HTTP応答をテストケース単位でスタブ化する。
    try {
      if (typeof httpScenario === "function") {
        httpScenario(mock);
      } else if (httpScenario !== null) {
        const { status = 200, data = {} } = httpScenario;
        mock.onAny().reply(status, data);
      }

      const context = vm.createContext({
        console,
        axios: http,
        __args: args,
        __c2: (id, value) => {
          const c = this.conditions.get(id);
          if (!c) throw new Error(`Unknown C2 condition: ${id}`);
          if (Boolean(value)) c.trueCount++;
          else c.falseCount++;
          return value;
        },
        $getTag: id => stub.getTag(id),
        $setTag: (id, value) => stub.setTag(id, value),
        $getTagId: name => stub.getTagId(name),
      });

      const script = new vm.Script(
        `${this.code}\nglobalThis.__result = __fuxaScript(...__args);`,
        { filename: this.file }
      );

      try {
        script.runInContext(context, { timeout: 1000 });
        const result = await context.__result;
        return { result, stub, httpHistory: mock.history };
      } finally {
        // 例外が発生した経路もカバレッジに反映する。
        if (context.__coverage__) this.coverage.merge(context.__coverage__);
      }
    } finally {
      mock.restore();
    }
  }

  summary() {
    const coverage = this.coverage.getCoverageSummary().data;
    const conditions = [...this.conditions.entries()].map(([id, c]) => ({
      id, ...c, covered: c.trueCount > 0 && c.falseCount > 0,
    }));

    return {
      c0: coverage.statements.pct,
      c1: coverage.branches.pct,
      functions: coverage.functions.pct,
      lines: coverage.lines.pct,
      c2: conditions.length
        ? conditions.filter(c => c.covered).length / conditions.length * 100
        : 100,
      conditions,
    };
  }

  // Istanbulの分岐情報と独自C2計測結果を同じHTMLへ統合する。
  writeC2Report(directory, result) {
    const file = this.coverage.fileCoverageFor(this.file);
    const branches = Object.entries(file.branchMap).map(([id, b]) => {
      const counts = file.b[id];
      return {
        line: b.line,
        type: b.type,
        counts,
        covered: counts.every(n => n > 0),
      };
    });

    const c1Rows = branches.map(b => `<tr>
      <td>${b.line}</td><td>${esc(b.type)}</td>
      <td>${b.counts.map((n, i) => `経路${i + 1}: ${n}回`).join(" / ")}</td>
      <td class="${b.covered ? "pass" : "fail"}">${b.covered ? "PASS" : "FAIL"}</td>
    </tr>`).join("");

    const c2Rows = result.conditions.map(c => `<tr>
      <td>${c.line}</td><td><code>${esc(c.expression)}</code></td>
      <td>${c.trueCount}</td><td>${c.falseCount}</td>
      <td class="${c.covered ? "pass" : "fail"}">${c.covered ? "PASS" : "FAIL"}</td>
    </tr>`).join("");

    // 指標とC1/C2の詳細比較表をソースの上に配置する。
    const panel = `<section id="fuxa-detail">
<style>
#fuxa-detail{background:#fff;color:#263244;padding:20px;margin:20px 0;border:1px solid #d6dee8;border-radius:8px}
#fuxa-detail .metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:15px 0}
#fuxa-detail .metric{background:#eef2f7;padding:16px;border-radius:6px}
#fuxa-detail .metric strong{display:block;font-size:26px}
#fuxa-detail table{width:100%;border-collapse:collapse;margin:12px 0 25px}
#fuxa-detail th,#fuxa-detail td{border:1px solid #d7dfe8;padding:10px;text-align:left}
#fuxa-detail th{background:#edf2f7}
#fuxa-detail .pass{background:#e7f5e9}
#fuxa-detail .fail{background:#ffe5e5}
#fuxa-detail code{overflow-wrap:anywhere}
</style>
<h2>FUXA カバレッジ詳細</h2>
<div class="metrics">
<div class="metric">C0 文網羅<strong>${result.c0.toFixed(2)}%</strong></div>
<div class="metric">C1 分岐網羅<strong>${result.c1.toFixed(2)}%</strong></div>
<div class="metric">C2 基本条件網羅<strong>${result.c2.toFixed(2)}%</strong></div>
</div>
<h2>C1：分岐経路</h2>
<table><thead><tr><th>行</th><th>分岐種別</th><th>経路別実行回数</th><th>判定</th></tr></thead>
<tbody>${c1Rows}</tbody></table>
<h2>C2：基本条件</h2>
<table><thead><tr><th>行</th><th>条件式</th><th>True</th><th>False</th><th>判定</th></tr></thead>
<tbody>${c2Rows}</tbody></table>
<p>C1は分岐経路、C2は各基本条件の真偽を評価します。</p>
</section>`;

    // Istanbulが生成したソース詳細ページを探す。
    const findPages = dir => fs.readdirSync(dir, { withFileTypes: true }).flatMap(e => {
      const name = path.join(dir, e.name);
      if (e.isDirectory()) return findPages(name);
      return e.isFile() && e.name.endsWith(".html") && e.name !== "index.html" ? [name] : [];
    });

    const target = findPages(directory).find(p =>
      path.basename(p) === `${path.basename(this.file)}.html`
    );
    if (!target) throw new Error(`Istanbul source page not found: ${this.file}`);

    // ページの先頭側に比較表を追加する。
    const original = fs.readFileSync(target, "utf8");
    const marker = "</h1>";
    const position = original.indexOf(marker);
    if (position < 0) throw new Error(`Istanbul heading not found: ${target}`);

    const updated = original.slice(0, position + marker.length) +
      panel + original.slice(position + marker.length);
    fs.writeFileSync(target, updated, "utf8");
  }

  report(directory = "coverage", threshold = 100) {
    const output = path.resolve(directory);
    const errors = [];

    fs.mkdirSync(output, { recursive: true });

    let result = null;
    try {
      result = this.summary();
    } catch (error) {
      errors.push(`Coverage summary: ${error.message}`);
    }

    // Istanbulの標準HTMLを生成する。
    // HTML・テキスト・C1/C2詳細の失敗を独立して集計する。
    if (this.coverage.files().length) {
      try {
        const context = libReport.createContext({ dir: output, coverageMap: this.coverage });
        reports.create("html").execute(context);
        console.log(`[PASS] Coverage HTML: ${output}`);
      } catch (error) {
        errors.push(`Coverage HTML: ${error.message}`);
      }

      try {
        const context = libReport.createContext({ dir: output, coverageMap: this.coverage });
        reports.create("text").execute(context);
      } catch (error) {
        errors.push(`Coverage text: ${error.message}`);
      }
    } else {
      errors.push("Coverage unavailable: no executed instrumented script");
    }

    // C1/C2の詳細をIstanbulのソース画面へ追加する。
    if (result && this.coverage.files().includes(this.file)) {
      try {
        this.writeC2Report(output, result);
        console.log("[PASS] C1/C2 HTML");
      } catch (error) {
        errors.push(`C1/C2 HTML: ${error.message}`);
      }
    }

    console.log("\n=== FUXA COVERAGE ===");

    if (result) {
      console.log(`C0: ${result.c0.toFixed(2)}%`);
      console.log(`C1: ${result.c1.toFixed(2)}%`);
      console.log(`C2: ${result.c2.toFixed(2)}%`);

      for (const c of result.conditions) {
        console.log(`[${c.covered ? "PASS" : "FAIL"}] ${c.expression} TRUE=${c.trueCount} FALSE=${c.falseCount}`);
      }

      // すべてのレポート出力後に品質ゲートを判定する。
      if ([result.c0, result.c1, result.c2].some(v => v < threshold)) {
        errors.push(
          `Coverage below ${threshold}%: C0=${result.c0.toFixed(2)}, ` +
          `C1=${result.c1.toFixed(2)}, C2=${result.c2.toFixed(2)}`
        );
      }
    }

    if (errors.length) {
      for (const message of errors) console.error(`[FAIL] ${message}`);
      throw new AggregateError(errors.map(message => new Error(message)), errors.join("\n"));
    }

    return result;
  }
}

module.exports = { FuxaDriver, FuxaStub };