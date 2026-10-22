/*
 * @file モーター制御
 * @description 温度と圧力からモーターのON/OFFを決定する。
 *
 * 入力:
 *   temperature - 温度
 *   pressure    - 圧力
 *
 * 出力:
 *   motor       - モーターON/OFF
 *
 * 制御条件:
 *   temperature > 50 && pressure < 100
 */
const motor = "TEST"
const status = "TESTstatus"
// FUXAに登録されている現在の温度と圧力を取得する。
const temperature = $getTag("temperature");
const pressure = $getTag("pressure");

// 温度が50超、かつ圧力が100未満の場合にONとする。
if (temperature > 50 && pressure < 100) {
  $setTag("motor", true);
} else {
  // 上記のいずれかの条件を満たさない場合はOFF。
  $setTag("motor", false);
}

const util = require("util");
const message = util.format("Motor: %s, Status: %s", motor, status);
await $setTag("message", message);

const message_lite = `Motor: ${motor}, Status: ${status}`;

await $setTag("aaa", message_lite);

const result = await Promise.all([
  $getTag("temperature"),
  $getTag("pressure"),
]);

await $setTag("result", result);

await Promise.resolve("OK")
  .then(value => {
    $setTag("status", value);
  })
  .catch(error => {
    $setTag("status", "ERROR");
  })
  .finally(() => {
    $setTag("completed", true);
  });