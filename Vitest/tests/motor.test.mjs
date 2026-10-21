import { describe, it, expect, afterAll } from "vitest";
import framework from "../framework.cjs";

const { FuxaDriver } = framework;
const driver = new FuxaDriver("targets/motor.js");

describe("FUXA Motor Script", () => {
  it("温度と圧力が正常条件の場合、モーターON", async () => {
    const { stub } = await driver.run({
      temperature: 60,
      pressure: 80,
    });

    expect(stub.value("motor")).toBe(true);
    expect(stub.writesTo("motor")).toHaveLength(1);
  });

  it("圧力条件が不成立の場合、モーターOFF", async () => {
    const { stub } = await driver.run({
      temperature: 60,
      pressure: 120,
    });

    expect(stub.value("motor")).toBe(false);
  });

  it("温度条件が不成立の場合、モーターOFF", async () => {
    const { stub } = await driver.run({
      temperature: 40,
      pressure: 80,
    });

    expect(stub.value("motor")).toBe(false);
    expect(stub.reads).toEqual(["temperature", "pressure"]);
  });
});

afterAll(() => {
  driver.report("coverage", 100);
});