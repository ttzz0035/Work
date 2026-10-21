import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, it, expect } from "vitest";
import framework from "../framework.cjs";

const { FuxaDriver } = framework;

// 一時的なFUXAスクリプトを作成してテストする。
async function withScript(source, callback) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "fuxa-c2-"));
  const filename = path.join(directory, "test.js");

  try {
    fs.writeFileSync(filename, source, "utf8");
    return await callback(filename);
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
}

describe("FUXA C2 Instrumenter", () => {
  it("AND条件のTrue/Falseを計測する", async () => {
    const source = `
      const a = $getTag("a");
      const b = $getTag("b");
      if (a && b) $setTag("result", true);
      else $setTag("result", false);
    `;

    await withScript(source, async (file) => {
      const driver = new FuxaDriver(file);

      await driver.run({ a: true, b: true });
      await driver.run({ a: true, b: false });
      await driver.run({ a: false, b: true });

      const result = driver.summary();
      expect(result.c2).toBe(100);
      expect(result.conditions).toHaveLength(2);
      expect(result.conditions.every((c) => c.covered)).toBe(true);
    });
  });

  it("OR条件の短絡評価を保持する", async () => {
    const source = `
      const a = $getTag("a");
      const b = $getTag("b");
      if (a || b) $setTag("result", true);
      else $setTag("result", false);
    `;

    await withScript(source, async (file) => {
      const driver = new FuxaDriver(file);

      await driver.run({ a: true, b: false });
      expect(driver.summary().conditions[1].trueCount).toBe(0);
      expect(driver.summary().conditions[1].falseCount).toBe(0);

      await driver.run({ a: false, b: true });
      await driver.run({ a: false, b: false });

      expect(driver.summary().c2).toBe(100);
    });
  });

  it("NOT条件を正しく計測する", async () => {
    const source = `
      const enabled = $getTag("enabled");
      if (!enabled) $setTag("result", "OFF");
      else $setTag("result", "ON");
    `;

    await withScript(source, async (file) => {
      const driver = new FuxaDriver(file);

      await driver.run({ enabled: true });
      await driver.run({ enabled: false });

      const result = driver.summary();
      expect(result.c2).toBe(100);
      expect(result.conditions[0].expression).toBe("!enabled");
    });
  });

  it("C2未達を100%と誤判定しない", async () => {
    const source = `
      const value = $getTag("value");
      if (value > 10) $setTag("result", true);
      else $setTag("result", false);
    `;

    await withScript(source, async (file) => {
      const driver = new FuxaDriver(file);

      await driver.run({ value: 20 });

      expect(driver.summary().c2).toBe(0);
      expect(driver.summary().conditions[0].covered).toBe(false);
    });
  });

  it("三項演算子の条件を計測する", async () => {
    const source = `
      const value = $getTag("value");
      $setTag("result", value > 10 ? "HIGH" : "LOW");
    `;

    await withScript(source, async (file) => {
      const driver = new FuxaDriver(file);

      await driver.run({ value: 20 });
      await driver.run({ value: 10 });

      expect(driver.summary().c2).toBe(100);
    });
  });

  it("switchは未対応として失敗させる", async () => {
    const source = `
      const value = $getTag("value");
      switch (value) {
        case 1: $setTag("result", true); break;
        default: $setTag("result", false);
      }
    `;

    await withScript(source, async (file) => {
      expect(() => new FuxaDriver(file)).toThrow("Unsupported C2 syntax");
    });
  });

  it("Nullish演算子は未対応として失敗させる", async () => {
    const source = `
      const value = $getTag("value") ?? 0;
      $setTag("result", value);
    `;

    await withScript(source, async (file) => {
      expect(() => new FuxaDriver(file)).toThrow("Unsupported C2 syntax");
    });
  });

  it("Optional chainingは未対応として失敗させる", async () => {
    const source = `
      const value = $getTag("value");
      if (value?.enabled) $setTag("result", true);
    `;

    await withScript(source, async (file) => {
      expect(() => new FuxaDriver(file)).toThrow("Unsupported C2 syntax");
    });
  });
});