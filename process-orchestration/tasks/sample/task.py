
import math
import time

from ui_log import UILayout, UIValue, UIProgress, UIStatus, emit_ui


class Task:
    """共通UI Schemaに従って表示データを定期更新するサンプルTask。"""

    def run(self, config):
        interval = config.get("interval", 1000) / 1000
        if interval <= 0:
            raise ValueError("interval must be positive")

        # Dashboardのグリッドを2列に設定する。
        emit_ui(UILayout(columns=2))

        # 初回のみUIカードを定義・生成する。
        emit_ui(UIValue(
            id="temperature", label="Temperature",
            unit="°C", row=0, column=0
        ))
        emit_ui(UIValue(
            id="pressure", label="Pressure",
            unit="kPa", row=0, column=1
        ))
        emit_ui(UIProgress(
            id="progress", label="Progress",
            row=1, column=0, columnspan=2
        ))
        emit_ui(UIStatus(
            id="connection", label="Connection",
            row=2, column=0, columnspan=2
        ))

        count = 0

        while True:
            # サンプルの測定値を生成する。
            temperature = round(25 + math.sin(count / 5) * 3, 1)
            pressure = round(101.3 + math.cos(count / 7) * 2, 1)
            progress = (count * 5) % 101

            # IDを指定して既存カードの値のみ更新する。
            emit_ui(UIValue(id="temperature", value=temperature))
            emit_ui(UIValue(id="pressure", value=pressure))
            emit_ui(UIProgress(id="progress", value=progress))
            emit_ui(UIStatus(id="connection", value="Connected"))

            count += 1
            time.sleep(interval)


if __name__ == "__main__":
    Task().run({"interval": 1000})
