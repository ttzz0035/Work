from __future__ import annotations

from datetime import datetime

from dash import Dash, Input, Output, dcc, html

from fuxa_client import get_daq_history
from gantt import create_figure


# ============================================================
# FUXAタグ設定
# ============================================================

# ガントチャートとして表示するDAQタグ
FUXA_TAG_ID = "t_2b7faf71-ae894970"


# ============================================================
# 表示範囲設定
#
# YYYY-MM-DD HH:MM:SS形式で指定する。
# Dashサーバーが動作しているOSのローカルtimezoneとして扱う。
# ============================================================

# FUXAから取得する開始日時
DISPLAY_START = "2026-08-03 13:00:00"

# FUXAから取得する終了日時
DISPLAY_END = "2026-08-03 13:30:00"


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

            * {
                box-sizing: border-box;
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
# 表示開始時刻・終了時刻取得
#
# DISPLAY_STARTとDISPLAY_ENDをdatetimeへ変換する。
# timezoneを含まない文字列はOSのローカルtimezoneとして扱う。
# ============================================================

def get_display_range() -> tuple[datetime, datetime]:
    start = datetime.fromisoformat(
        DISPLAY_START
    ).astimezone()

    end = datetime.fromisoformat(
        DISPLAY_END
    ).astimezone()

    if start >= end:
        raise ValueError(
            "DISPLAY_STARTはDISPLAY_ENDより前にしてください。"
        )

    return start, end


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
    histories = get_daq_history(
        [
            FUXA_TAG_ID,
        ],
        start,
        end,
    )

    # タグIDを1件だけ指定しているため、対象タグの履歴を取得する
    history = histories[FUXA_TAG_ID]

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
# 定数で指定した表示範囲を使用してFUXAからDAQ履歴を取得し、
# PlotlyのFigureを1秒ごとに再生成する。
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