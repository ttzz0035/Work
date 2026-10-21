import math
import sys
import time
from pathlib import Path

# Task単体実行時にもプロジェクト共通モジュールを参照できるようにする。
BASE = Path(__file__).resolve().parents[2]

if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from ui_log import UILamp, UILayout, UIProgress, UIValue, emit_ui


class Task:
    """共通UI Schemaに従って表示データを定期更新するサンプルTask。"""

    def run(self, config):
        interval = config.get("interval", 1000) / 1000
        if interval <= 0:
            raise ValueError("interval must be positive")

        # Dashboardのグリッドを3列に設定する。
        emit_ui(UILayout(columns=3))

        # 測定値と進捗のカードを定義する。
        emit_ui(UIValue(id="temperature", label="Temperature", unit="°C", row=0, column=0))
        emit_ui(UIValue(id="pressure", label="Pressure", unit="kPa", row=0, column=1))
        emit_ui(UIValue(id="connection", label="Connection", row=0, column=2))

        # 通常のProgressbarを定義する（0〜100%）。
        emit_ui(UIProgress(id="progress", label="Progress", row=1, column=0, columnspan=3))

        # 基準10秒を中心として、5〜15秒の範囲を表示する。
        emit_ui(UIProgress(
            id="execution_time", label="Execution Time",
            reference=10.0, range=5.0, unit="s",
            row=2, column=0, columnspan=3
        ))

        # ランプを1つだけ定義する。
        emit_ui(UILamp(id="state", label="State", color="green", text="Running", row=3, column=0, columnspan=3))

        count = 0
        states = (
            ("green", "Running"),
            ("yellow", "Warning"),
            ("red", "Error"),
        )

        while True:
            # サンプル測定値を生成する。
            temperature = round(25 + math.sin(count / 5) * 3, 1)
            pressure = round(101.3 + math.cos(count / 7) * 2, 1)
            progress = (count * 5) % 101
            execution_time = round(10 + math.sin(count / 5) * 4, 2)

            # 既存カードの値のみ更新する。
            emit_ui(UIValue(id="temperature", value=temperature))
            emit_ui(UIValue(id="pressure", value=pressure))
            emit_ui(UIValue(id="connection", value="Connected", color="green"))
            emit_ui(UIProgress(id="progress", value=progress))
            emit_ui(UIProgress(id="execution_time", value=execution_time))

            # 1つのランプを状態に応じて緑・黄・赤に切り替える。
            color, text = states[(count // 5) % len(states)]
            emit_ui(UILamp(id="state", value=True, color=color, text=text))

            count += 1
            time.sleep(interval)


if __name__ == "__main__":
    Task().run({"interval": 1000})
