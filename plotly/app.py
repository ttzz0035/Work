from __future__ import annotations

import json
from datetime import datetime, timedelta

import requests
from dash import Dash, Input, Output, dcc, html

from plotly.gantt import create_figure


# ============================================================
# FUXA接続設定
# ============================================================

# FUXAのHTTP API
FUXA_URL = "http://127.0.0.1:1881"

# ガントチャートとして表示するDAQタグ
FUXA_TAG_ID = "t_2b7faf71-ae894970"


# ============================================================
# 表示範囲モード
# ============================================================

# 現在時刻を終点として、直近N時間を表示する
DISPLAY_MODE_ROLLING = 0

# 今日の0:00を始点として、N時間分を表示する
DISPLAY_MODE_DAILY = 1

# 今週月曜日の0:00を始点として、N週間分を表示する
DISPLAY_MODE_WEEKLY = 2


# ============================================================
# 現在使用する表示設定
# ============================================================

# DAILY + 24の場合:
# 今日の0:00から翌日の0:00までを表示する
DISPLAY_MODE = DISPLAY_MODE_DAILY
DISPLAY_RANGE = 24


# ============================================================
# 将来FUXAタグから表示設定を取得する場合のタグ
# ============================================================

# FUXA_DISPLAY_MODE_TAG_ID = "表示モードタグID"
# FUXA_DISPLAY_RANGE_TAG_ID = "表示範囲タグID"


# ============================================================
# Dashアプリケーション
# ============================================================

app = Dash(__name__)


# ============================================================
# IFRAME全面表示用HTML
#
# html、body、Dashのルート要素を100%にして、
# FUXAのIFRAMEサイズへDash画面を追従させる
# ============================================================

app.index_string = """
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            html,
            body,
            #react-entry-point,
            #_dash-app-content {
                width: 100%;
                height: 100%;
                margin: 0;
                padding: 0;
                overflow: hidden;
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
"""


# ============================================================
# 固定データを返す旧実装
#
# FUXAを使用せず、表示確認だけを行う場合に使用する
# ============================================================

# def get_stack_items() -> list[dict]:
#     return [
#         {"category": "A", "duration_days": 3},
#         {"category": "B", "duration_days": 2},
#         {"category": "A", "duration_days": 4},
#         {"category": "C", "duration_days": 3},
#         {"category": "B", "duration_days": 2},
#     ]


# ============================================================
# 表示設定取得
#
# 現在は固定値を返す。
# 将来はこの関数内だけを変更し、FUXAタグから取得する。
# ============================================================

def get_display_settings() -> tuple[int, int]:
    return DISPLAY_MODE, DISPLAY_RANGE


# ============================================================
# 表示開始時刻・終了時刻生成
#
# modeとrangeを基に、DAQ取得範囲とX軸表示範囲を生成する。
# ============================================================

def get_display_range() -> tuple[datetime, datetime]:
    mode, display_range = get_display_settings()
    now = datetime.now().astimezone()

    # 現在時刻から直近N時間
    if mode == DISPLAY_MODE_ROLLING:
        end = now
        start = end - timedelta(hours=display_range)

        return start, end

    # 今日の0:00からN時間
    if mode == DISPLAY_MODE_DAILY:
        start = now.replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0,
        )
        end = start + timedelta(hours=display_range)

        return start, end

    # 今週月曜日の0:00からN週間
    if mode == DISPLAY_MODE_WEEKLY:
        start = (
            now - timedelta(days=now.weekday())
        ).replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0,
        )
        end = start + timedelta(weeks=display_range)

        return start, end

    raise ValueError(
        f"未対応の表示モードです: {mode}"
    )


# ============================================================
# FUXA DAQ履歴取得・区間変換
#
# 指定時間範囲のDAQ履歴を取得し、
# 同じ値が連続している期間を1つの区間にまとめる。
#
# 例:
# 0, 0, 0, 1, 1, 0
#
# 変換結果:
# 0の区間 → 1の区間 → 0の区間
# ============================================================

def get_stack_items(
    start: datetime,
    end: datetime,
) -> list[dict]:
    query = {
        "sids": [
            FUXA_TAG_ID,
        ],
        "from": int(start.timestamp() * 1000),
        "to": int(end.timestamp() * 1000),
    }

    response = requests.get(
        f"{FUXA_URL}/api/daq",
        params={
            "query": json.dumps(
                query,
                separators=(",", ":"),
            ),
        },
        timeout=10,
    )
    response.raise_for_status()

    # sidsを1件だけ指定しているため、先頭の配列を取得する
    history = response.json()[0]

    if not history:
        return []

    stack_items = []

    # 最初のDAQ値を最初の区間として扱う
    category = str(
        history[0]["value"]
    )
    segment_start = history[0]["dt"]

    # 値が変化した時点で、直前までの区間を確定する
    for item in history[1:]:
        value = str(
            item["value"]
        )

        if value == category:
            continue

        stack_items.append(
            {
                "category": category,
                "start": datetime.fromtimestamp(
                    segment_start / 1000,
                    tz=start.tzinfo,
                ),
                "finish": datetime.fromtimestamp(
                    item["dt"] / 1000,
                    tz=start.tzinfo,
                ),
            }
        )

        category = value
        segment_start = item["dt"]

    # 最後の区間は現在時刻または表示終了時刻まで延長する
    stack_items.append(
        {
            "category": category,
            "start": datetime.fromtimestamp(
                segment_start / 1000,
                tz=start.tzinfo,
            ),
            "finish": min(
                end,
                datetime.now().astimezone(),
            ),
        }
    )

    return stack_items


# ============================================================
# Dash画面構成
#
# GraphはIFRAME全体へ追従させる。
# Intervalは1秒ごとにグラフ更新を発生させる。
# ============================================================

app.layout = html.Div(
    [
        dcc.Graph(
            id="gantt-chart",
            config={
                # Plotlyツールバーを非表示にする
                "displayModeBar": False,

                # IFRAMEのサイズ変更へ追従する
                "responsive": True,
            },
            style={
                "width": "100%",
                "height": "100%",
            },
        ),
        dcc.Interval(
            id="refresh",

            # グラフ更新間隔をミリ秒で指定する
            interval=1000,

            # Dashが更新回数を管理する
            n_intervals=0,
        ),
    ],
    style={
        "width": "100%",
        "height": "100%",
        "margin": "0",
        "padding": "0",
    },
)


# ============================================================
# グラフ更新
#
# 表示範囲を計算し、FUXAからDAQ履歴を取得して、
# PlotlyのFigureを再生成する。
# ============================================================

@app.callback(
    Output(
        "gantt-chart",
        "figure",
    ),
    Input(
        "refresh",
        "n_intervals",
    ),
)
def update_graph(_):
    start, end = get_display_range()

    stack_items = get_stack_items(
        start,
        end,
    )

    return create_figure(
        stack_items,
        start,
        end,
    )


# ============================================================
# Dashサーバー起動
# ============================================================

if __name__ == "__main__":
    app.run(
        host="127.0.0.1",
        port=8050,
        debug=False,
    )