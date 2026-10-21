/*
 * @file 複合条件によるモーター制御
 * @description 温度・圧力を評価して運転状態を分類する。
 *
 * 入力:
 *   temperature - 温度
 *   pressure    - 圧力
 *
 * 出力:
 *   status      - 判定結果
 *   motor       - モーターON/OFF
 *
 * 判定結果:
 *   INVALID       - 入力異常
 *   CRITICAL      - 危険状態
 *   HIGH_TEMP     - 高温警告
 *   HIGH_PRESSURE - 高圧警告
 *   WARNING       - その他警告
 *   NORMAL        - 正常
 */

// 現在の温度・圧力をFUXAタグから取得する。
const temperature = $getTag("temperature");
const pressure = $getTag("pressure");

// 異常値を判定する。
// どちらかが負数であれば処理を停止する状態とする。
if (temperature < 0 || pressure < 0) {
  $setTag("status", "INVALID");
  $setTag("motor", false);
}

// 温度・圧力ともに危険域の場合。
else if (temperature > 80 && pressure > 120) {
  $setTag("status", "CRITICAL");
  $setTag("motor", false);
}

// 温度または圧力が警告域の場合。
else if (temperature > 60 || pressure > 100) {

  // 温度と圧力の組合せによる追加判定。
  if (temperature >= 70 && pressure <= 110) {
    $setTag("status", "HIGH_TEMP");
    $setTag("motor", false);
  }

  // 圧力が高く、温度も一定以上の場合。
  else if (pressure >= 110 && temperature > 50) {
    $setTag("status", "HIGH_PRESSURE");
    $setTag("motor", false);
  }

  // その他の警告状態。
  else {
    $setTag("status", "WARNING");
    $setTag("motor", true);
  }
}

// すべての条件が正常な場合。
else {
  $setTag("status", "NORMAL");
  $setTag("motor", true);
}