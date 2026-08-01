from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# グラフ表示設定
# ============================================================

# IFRAMEの幅と高さに合わせて自動調整する
GRAPH_AUTOSIZE = True

# Categoryの凡例を表示するか
GRAPH_SHOW_LEGEND = False

# マウス位置に最も近いバーのホバーを表示する
GRAPH_HOVER_MODE = "closest"

# ドラッグによるズームや範囲選択を無効にする
GRAPH_DRAG_MODE = False

# IFRAME内に収まるよう余白を小さくする
GRAPH_MARGIN = {
    "l": 40,
    "r": 10,
    "t": 10,
    "b": 40,
}


# ============================================================
# X軸表示設定
# ============================================================

# X軸タイトル
X_AXIS_TITLE = "時刻"

# X軸へ表示する日時形式
X_AXIS_TICK_FORMAT = "%m-%d %H:%M:%S"

# 自動でX軸範囲を変更するか
X_AXIS_AUTORANGE = False

# ズームや移動によるX軸範囲の変更を禁止する
X_AXIS_FIXED_RANGE = True


# ============================================================
# Y軸表示設定
# ============================================================

# Y軸タイトル
Y_AXIS_TITLE = ""

# 工程行のグリッドを表示するか
Y_AXIS_SHOW_GRID = False

# ズームや移動によるY軸範囲の変更を禁止する
Y_AXIS_FIXED_RANGE = True

# 空グラフでY軸ラベルを表示するか
EMPTY_Y_AXIS_SHOW_TICK_LABELS = False


# ============================================================
# バー表示設定
# ============================================================

# すべての区間を表示する行名
BAR_ROW_NAME = "工程"

# バー内部の文字位置
BAR_TEXT_POSITION = "inside"

# バー内部の文字寄せ
BAR_INSIDE_TEXT_ANCHOR = "middle"


# ============================================================
# ホバー表示設定
#
# customdataの位置:
# 0: Category
# 1: Start
# 2: Finish
# 3: DurationSeconds
#
# 不要な行はこの定数から削除する。
# ============================================================

HOVER_TEMPLATE = (
    "状態: %{customdata[0]}<br>"
    "開始: %{customdata[1]|%Y-%m-%d %H:%M:%S}<br>"
    "終了: %{customdata[2]|%Y-%m-%d %H:%M:%S}<br>"
    "継続時間: %{customdata[3]:.1f} 秒"
    "<extra></extra>"
)


# ============================================================
# 空データ用Figure生成
#
# FUXAからDAQ履歴を取得できなかった場合でも、
# 指定された表示時間範囲を持つ空のグラフを表示する。
# ============================================================

def create_empty_figure(
    display_start: datetime,
    display_end: datetime,
) -> go.Figure:
    figure = go.Figure()

    # グラフ全体の表示と操作を設定する
    figure.update_layout(
        autosize=GRAPH_AUTOSIZE,
        margin=GRAPH_MARGIN,
        showlegend=GRAPH_SHOW_LEGEND,
        hovermode=GRAPH_HOVER_MODE,
        dragmode=GRAPH_DRAG_MODE,
        xaxis_title=X_AXIS_TITLE,
        yaxis_title=Y_AXIS_TITLE,
    )

    # データがなくても指定された時間範囲をX軸へ表示する
    figure.update_xaxes(
        range=[
            display_start,
            display_end,
        ],
        autorange=X_AXIS_AUTORANGE,
        tickformat=X_AXIS_TICK_FORMAT,
        fixedrange=X_AXIS_FIXED_RANGE,
    )

    # 空データ時は工程名などのY軸ラベルを表示しない
    figure.update_yaxes(
        showticklabels=EMPTY_Y_AXIS_SHOW_TICK_LABELS,
        fixedrange=Y_AXIS_FIXED_RANGE,
    )

    return figure


# ============================================================
# FUXA区間データからDataFrame生成
#
# app.pyで生成した次の形式をPlotly用へ変換する。
#
# {
#     "category": "0",
#     "start": datetime,
#     "finish": datetime,
# }
# ============================================================

def create_dataframe(
    stack_items: list[dict],
) -> pd.DataFrame:
    rows = []

    for item in stack_items:
        start = item["start"]
        finish = item["finish"]

        rows.append(
            {
                # すべての区間を同じ横一列へ表示する
                "Row": BAR_ROW_NAME,

                # FUXAタグの値を区間種別として使用する
                "Category": item["category"],

                # 区間の開始時刻
                "Start": start,

                # 区間の終了時刻
                "Finish": finish,

                # バー内部へ表示する文字
                "Label": item["category"],

                # ホバーへ表示する継続時間
                "DurationSeconds": (
                    finish - start
                ).total_seconds(),
            }
        )

    return pd.DataFrame(rows)


# ============================================================
# ガントチャート生成
#
# FUXAの値が連続していた区間を横方向へ並べ、
# 指定された時間範囲内に表示する。
# ============================================================

def create_figure(
    stack_items: list[dict],
    display_start: datetime,
    display_end: datetime,
) -> go.Figure:
    # 表示対象がない場合は時間軸だけの空グラフを返す
    if not stack_items:
        return create_empty_figure(
            display_start,
            display_end,
        )

    dataframe = create_dataframe(
        stack_items
    )

    # StartからFinishまでを横棒として描画する
    figure = px.timeline(
        dataframe,
        x_start="Start",
        x_end="Finish",
        y="Row",
        color="Category",
        text="Label",
        custom_data=[
            "Category",
            "Start",
            "Finish",
            "DurationSeconds",
        ],
    )

    # バー内部の文字とホバー内容を設定する
    figure.update_traces(
        textposition=BAR_TEXT_POSITION,
        insidetextanchor=BAR_INSIDE_TEXT_ANCHOR,
        hovertemplate=HOVER_TEMPLATE,
    )

    # IFRAMEサイズへの追従とグラフ操作を設定する
    figure.update_layout(
        xaxis_title=X_AXIS_TITLE,
        yaxis_title=Y_AXIS_TITLE,
        autosize=GRAPH_AUTOSIZE,
        margin=GRAPH_MARGIN,
        showlegend=GRAPH_SHOW_LEGEND,
        hovermode=GRAPH_HOVER_MODE,
        dragmode=GRAPH_DRAG_MODE,
    )

    # app.pyで計算した時間範囲を固定して表示する
    figure.update_xaxes(
        range=[
            display_start,
            display_end,
        ],
        autorange=X_AXIS_AUTORANGE,
        tickformat=X_AXIS_TICK_FORMAT,
        fixedrange=X_AXIS_FIXED_RANGE,
    )

    # Y軸の表示と操作を設定する
    figure.update_yaxes(
        showgrid=Y_AXIS_SHOW_GRID,
        fixedrange=Y_AXIS_FIXED_RANGE,
    )

    return figure