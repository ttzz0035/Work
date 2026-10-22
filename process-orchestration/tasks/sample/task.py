
import math
import sys
import time
from pathlib import Path

# Task単体実行時にもプロジェクト共通モジュールを参照できるようにする。
BASE = Path(__file__).resolve().parents[2]

if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

# プロジェクト共通モジュール
from ui_log import UILamp, UILayout, UIProgress, UIValue, emit_ui

# Task同階層モジュール
# from tasks.sample.collector import Collector
# from tasks.sample.analyzer import Analyzer


class Task:
    """共通UI Schemaに従って表示データを定期更新するサンプルTask。"""

    def run(self, config, node_name):
        # Workerから渡されたNode名を取得する。
        print(f"[NODE] {node_name}", flush=True)

        interval = config.get("interval", 1000) / 1000
        if interval <= 0:
            raise ValueError("interval must be positive")

        count = 0
        states = (
            ("#22C55E", "Running"),
            ("#F59E0B", "Warning"),
            ("#EF4444", "Error"),
        )

        while True:
            # サンプル測定値を生成する。
            temperature = round(25 + math.sin(count / 5) * 3, 1)
            pressure = round(101.3 + math.cos(count / 7) * 2, 1)
            progress = (count * 5) % 101
            execution_time = round(10 + math.sin(count / 5) * 4, 2)
            color, text = states[(count // 5) % len(states)]

            # Dashboardのグリッドを4列に設定する。
            emit_ui(UILayout(columns=4))

            # 測定値と進捗のカード定義・最新値をまとめて送信する。
            emit_ui(UIValue(id="temperature", label="Temperature", value=temperature, unit="°C", row=0, column=0))
            emit_ui(UIValue(id="pressure", label="Pressure", value=pressure, unit="kPa", row=0, column=1))
            emit_ui(UIValue(id="connection", label="Connection", value="Connected", color="#22C55E", row=0, column=2))

            # 通常のProgressbarを定義・更新する（0〜100%）。
            emit_ui(UIProgress(id="progress", label="Progress", value=progress, row=1, column=0, columnspan=2))

            # 実行時間のカードを基準値・表示範囲・実測値とともに送信する。
            emit_ui(UIProgress(id="execution_time", label="Execution Time", reference=10.0, range=(-10.0, 5.0), unit="s", value=execution_time, row=1, column=2, columnspan=2))

            # ランプを1つだけ定義し、状態に応じて緑・黄・赤に切り替える。
            emit_ui(UILamp(id="state", label="State", color=color, text=text, row=3, column=0, columnspan=3))

            count += 1
            time.sleep(interval)


if __name__ == "__main__":
    import yaml

    # VS CodeからTaskを直接デバッグする。
    DEBUG_NODE = "XXX"

    # nodes.yamlから指定Nodeの設定を取得する。
    with open(BASE / "nodes.yaml", "r", encoding="utf-8") as f:
        node = yaml.safe_load(f)["nodes"][DEBUG_NODE]

    # Workerから呼び出す場合と同じ引数で実行する。
    Task().run(node["parameters"], DEBUG_NODE)
