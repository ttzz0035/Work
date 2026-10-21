/*
 * @file カバレッジNG判定の検証
 * @description 到達不能な条件を意図的に含むテスト用コード。
 *
 * 目的:
 *   C0・C1・C2の未達検出が正常に機能することを確認する。
 *
 * 期待結果:
 *   C0 < 100%
 *   C1 < 100%
 *   C2 < 100%
 *
 * 通常の制御スクリプトとしては使用しない。
 */

// FUXAホワイトボックステストのNG判定確認用。
// 意図的に到達不能な条件を含めている。

const temperature = $getTag("temperature");

// 正常に実行できる分岐。
if (temperature > 50) {
  $setTag("motor", true);
} else {
  $setTag("motor", false);
}

// 同じ値が同時に100超と0未満になることはない。
// このAND条件は論理的に常にFalseとなる。
if (temperature > 100 && temperature < 0) {
  // 実行不可能な処理なので、C0未達になる。
  $setTag("status", "IMPOSSIBLE");
} else {
  $setTag("status", "NORMAL");
}