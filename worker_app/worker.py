# SPDX-License-Identifier: MIT
import logging
import time
from datetime import datetime

def run_worker(runtime: dict, append_logs: callable, update_status: callable):
    """
    本番テンプレート: 登録/照合いずれにも対応。
    - runtime: 実行状態(dict)
    - append_logs(): UIログ反映
    - update_status(): UIステータス反映
    """
    rt = runtime
    rt["started_at"] = datetime.now()
    rt["ticks"] = 0
    rt["running"] = True

    logging.info(
        f"[WORKER] 開始: モード={rt.get('mode','?')} / "
        f"ジョブID={rt.get('item_id','-')} / "
        f"期間={rt.get('start')}～{rt.get('end')}"
    )

    while rt["running"]:
        rt["ticks"] += 1
        rt["last_tick_at"] = datetime.now()

        # 実行状況を都度ログ
        logging.info(
            f"[WORKER] 実行中: "
            f"モード={rt.get('mode','?')} / "
            f"ジョブID={rt.get('item_id','-')} / "
            f"期間={rt.get('start')}～{rt.get('end')} / "
            f"実行回数={rt['ticks']}回"
        )

        append_logs()
        update_status()

        # ▼▼▼ 本番処理（登録/照合）をここに実装 ▼▼▼
        try:
            if rt["mode"] == "register":
                # TODO: 登録ロジック
                # 例) insert_to_db(...), call_api(...)
                pass
            elif rt["mode"] == "verify":
                # TODO: 照合ロジック
                # 例) compare_records(...), fetch_and_match(...)
                pass
            else:
                logging.warning(f"[WORKER] 不明なモード={rt.get('mode')}")
                rt["running"] = False
        except Exception as e:
            logging.error(f"[WORKER] エラー: {e}")
            rt["running"] = False
            break
        # ▲▲▲ ここまで差し替えポイント ▲▲▲

        time.sleep(1.0)

        # デバッグ上限（本番では外部条件で停止）
        if rt["ticks"] >= 10:
            logging.info(f"[WORKER] 規定回数に到達 → 終了 (実行回数={rt['ticks']})")
            rt["running"] = False

    logging.info(f"[WORKER] 終了: 実行回数={rt['ticks']}回")
    append_logs()
    update_status()
