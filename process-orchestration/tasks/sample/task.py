
import math
import time

from ui_log import UILamp, UILayout, UIProgress, UIStatus, UIValue, emit_ui


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
        emit_ui(UIStatus(id="connection", label="Connection", row=0, column=2))
        emit_ui(UIProgress(id="progress", label="Progress", row=1, column=0, columnspan=3))

        # ランプを1つだけ定義する。
        emit_ui(UILamp(id="state", label="State", color="green", text="Running", row=2, column=0, columnspan=3))

        count = 0

        while True:
            # サンプル測定値を生成する。
            temperature = round(25 + math.sin(count / 5) * 3, 1)
            pressure = round(101.3 + math.cos(count / 7) * 2, 1)
            progress = (count * 5) % 101

            # 既存カードの値のみ更新する。
            emit_ui(UIValue(id="temperature", value=temperature))
            emit_ui(UIValue(id="pressure", value=pressure))
            emit_ui(UIStatus(id="connection", value="Connected"))
            emit_ui(UIProgress(id="progress", value=progress))

            # 1つのランプを状態に応じて緑・黄・赤に切り替える。
            states = (
                ("green", "Running"),
                ("yellow", "Warning"),
                ("red", "Error"),
            )
            color, text = states[(count // 5) % len(states)]
            emit_ui(UILamp(id="state", value=True, color=color, text=text))

            count += 1
            time.sleep(interval)


if __name__ == "__main__":
    Task().run({"interval": 1000})
