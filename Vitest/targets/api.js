/*
 * @file Axiosによる運転状態取得
 * @description HTTP応答からモーターの運転状態を決定する。
 *
 * 外部依存:
 *   axios - HTTPクライアント
 *
 * API:
 *   GET https://example.com/api/status
 *
 * 出力:
 *   motor   - モーターON/OFF
 *   status  - 処理結果
 *   checked - 処理完了
 *
 * テスト時:
 *   axios-mock-adapterを使い、実HTTP通信を行わない。
 */

// Axiosの戻り値に応じてFUXAタグを更新する。
try {
  // 外部APIから現在の運転状態を取得する。
  const response = await axios.get("https://example.com/api/status");

  // HTTP 200かつrunning=trueの場合は運転状態とする。
  if (response.status === 200 && response.data.running === true) {
    $setTag("motor", true);
    $setTag("status", "RUNNING");
  } else {
    // 正常なHTTP応答だが運転条件が成立しない。
    $setTag("motor", false);
    $setTag("status", "STOPPED");
  }
} catch (error) {
  // API取得に失敗した場合はモーターOFFとする。
  $setTag("motor", false);

  // HTTP 500以上はサーバー側のエラーとして分類する。
  if (error.response && error.response.status >= 500) {
    $setTag("status", "SERVER_ERROR");
  } else {
    // HTTP 400系や通信エラーなどを分類する。
    $setTag("status", "REQUEST_ERROR");
  }
}

// 正常系・異常系の合流後に実行する共通処理。
$setTag("checked", true);