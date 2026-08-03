from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# 表示設定
# ============================================================

# 上部の状態凡例全体を表示するか
# 凡例全体は空文字では非表示にできないため、この設定だけ残す
SHOW_STATE_LEGEND = True

# バー内部へ表示する文字
# "{category}":
#   バー内部へ「1」「2」などを表示する
# "":
#   バー内部の文字を非表示にする
BAR_LABEL_TEMPLATE = "{category}"

# Y軸へ表示する工程名
# "工程":
#   左側へ「工程」を表示する
# "":
#   左側の工程名を非表示にする
BAR_ROW_NAME = "工程"

# 凡例タイトル
# "状態比率":
#   凡例タイトルを表示する
# "":
#   凡例タイトルを非表示にする
LEGEND_TITLE = "状態比率"

# 凡例へ表示する状態文字列
# "{category}: {ratio:.1f}%":
#   「1: 77.6%」の形式で状態と比率を表示する
# "{category}":
#   状態だけを表示する
# "":
#   凡例の文字を空にする
LEGEND_LABEL_TEMPLATE = "{category}: {ratio:.1f}%"

# ホバーへ追加する状態比率表示
# 空文字にすると状態合計時間と状態比率を非表示にする
HOVER_RATIO_TEMPLATE = (
    "<br>"
    "状態合計時間: %{customdata[4]:.1f} 秒<br>"
    "状態比率: %{customdata[5]:.1f}%"
)


# ============================================================
# グラフ表示設定
# ============================================================

# IFRAMEの幅と高さに合わせて自動調整する
GRAPH_AUTOSIZE = True

# マウス位置に最も近いバーのホバーを表示する
GRAPH_HOVER_MODE = "closest"

# ドラッグによるズームや範囲選択を無効にする
GRAPH_DRAG_MODE = False

# 凡例を表示する場合の余白
GRAPH_MARGIN_WITH_LEGEND = {
    "l": 40,
    "r": 10,
    "t": 60,
    "b": 40,
}

# 凡例を表示しない場合の余白
GRAPH_MARGIN_WITHOUT_LEGEND = {
    "l": 40,
    "r": 10,
    "t": 10,
    "b": 40,
}


# ============================================================
# 状態表示設定
#
# FUXAタグから取得する状態値ごとに色を固定する。
# 使用する状態はすべてここへ登録する。
#
# 未登録の状態を受信した場合は、自動色を割り当てず、
# 設定漏れとしてValueErrorを発生させる。
# ============================================================

STATE_COLOR_MAP = {
    "0": "#6B7280",
    "1": "#16A34A",
    "2": "#F59E0B",
    "3": "#DC2626",
}


# ============================================================
# X軸表示設定
# ============================================================

# X軸タイトル
# 空文字にすると非表示になる
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
# 空文字にすると非表示になる
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

# バー内部の文字位置
BAR_TEXT_POSITION = "inside"

# バー内部の文字寄せ
BAR_INSIDE_TEXT_ANCHOR = "middle"


# ============================================================
# グラフ余白取得
# ============================================================

def get_graph_margin() -> dict:
    if SHOW_STATE_LEGEND:
        return GRAPH_MARGIN_WITH_LEGEND

    return GRAPH_MARGIN_WITHOUT_LEGEND


# ============================================================
# ホバー表示設定生成
#
# customdataの位置:
# 0: Category
# 1: Start
# 2: Finish
# 3: DurationSeconds
# 4: StateDurationSeconds
# 5: StateRatio
# ============================================================

def create_hover_template() -> str:
    return (
        "状態: %{customdata[0]}<br>"
        "開始: %{customdata[1]|%Y-%m-%d %H:%M:%S}<br>"
        "終了: %{customdata[2]|%Y-%m-%d %H:%M:%S}<br>"
        "区間継続時間: %{customdata[3]:.1f} 秒"
        + HOVER_RATIO_TEMPLATE
        + "<extra></extra>"
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
        margin=get_graph_margin(),
        showlegend=False,
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
# 状態色設定検証
#
# FUXAから取得した状態がSTATE_COLOR_MAPへ登録されているか確認する。
# ============================================================

def validate_state_colors(
    stack_items: list[dict],
) -> None:
    categories = {
        str(item["category"])
        for item in stack_items
    }

    undefined_categories = sorted(
        categories - set(STATE_COLOR_MAP)
    )

    if undefined_categories:
        raise ValueError(
            "STATE_COLOR_MAPに未登録の状態があります: "
            + ", ".join(undefined_categories)
        )


# ============================================================
# バー内部表示文字生成
# ============================================================

def create_bar_label(
    category: str,
) -> str:
    return BAR_LABEL_TEMPLATE.format(
        category=category,
    )


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
#
# 状態比率は、表示対象として取得した全区間の合計時間に対する、
# 各状態の合計時間の割合として計算する。
# ============================================================

def create_dataframe(
    stack_items: list[dict],
) -> pd.DataFrame:
    rows = []

    for item in stack_items:
        start = item["start"]
        finish = item["finish"]
        category = str(item["category"])

        rows.append(
            {
                # すべての区間を同じ横一列へ表示する
                "Row": BAR_ROW_NAME,

                # FUXAタグの値を区間種別として使用する
                "Category": category,

                # 区間の開始時刻
                "Start": start,

                # 区間の終了時刻
                "Finish": finish,

                # バー内部へ表示する文字
                "Label": create_bar_label(
                    category
                ),

                # ホバーと状態比率計算へ使用する継続時間
                "DurationSeconds": (
                    finish - start
                ).total_seconds(),
            }
        )

    dataframe = pd.DataFrame(rows)

    state_duration_seconds = (
        dataframe.groupby(
            "Category",
            sort=False,
        )["DurationSeconds"]
        .sum()
        .to_dict()
    )

    total_duration_seconds = dataframe[
        "DurationSeconds"
    ].sum()

    if total_duration_seconds <= 0:
        raise ValueError(
            "状態比率を計算できる有効な時間区間がありません。"
        )

    dataframe["StateDurationSeconds"] = dataframe[
        "Category"
    ].map(
        state_duration_seconds
    )

    dataframe["StateRatio"] = (
        dataframe["StateDurationSeconds"]
        / total_duration_seconds
        * 100
    )

    return dataframe


# ============================================================
# 凡例表示更新
#
# LEGEND_LABEL_TEMPLATEの例:
#
# "{category}: {ratio:.1f}%"
#   1: 77.6%
#
# "{category}"
#   1
#
# ""
#   空文字
# ============================================================

def update_legend_labels(
    figure: go.Figure,
    dataframe: pd.DataFrame,
) -> None:
    state_ratios = (
        dataframe.groupby(
            "Category",
            sort=False,
        )["StateRatio"]
        .first()
        .to_dict()
    )

    for trace in figure.data:
        category = str(trace.name)

        trace.name = LEGEND_LABEL_TEMPLATE.format(
            category=category,
            ratio=state_ratios[category],
        )


# ============================================================
# ガントチャート生成
#
# FUXAの値が連続していた区間を横方向へ並べ、
# 指定された時間範囲内に表示する。
#
# 状態ごとの色はSTATE_COLOR_MAPで固定する。
# 空文字で非表示にできる項目は文字列設定で制御する。
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

    validate_state_colors(
        stack_items
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
            "StateDurationSeconds",
            "StateRatio",
        ],
        color_discrete_map=STATE_COLOR_MAP,
        category_orders={
            "Category": list(
                STATE_COLOR_MAP.keys()
            ),
        },
    )

    # 凡例名へ状態比率を付与する
    update_legend_labels(
        figure,
        dataframe,
    )

    # バー内部の文字とホバー内容を設定する
    figure.update_traces(
        textposition=BAR_TEXT_POSITION,
        insidetextanchor=BAR_INSIDE_TEXT_ANCHOR,
        hovertemplate=create_hover_template(),
    )

    # IFRAMEサイズへの追従とグラフ操作を設定する
    figure.update_layout(
        xaxis_title=X_AXIS_TITLE,
        yaxis_title=Y_AXIS_TITLE,
        autosize=GRAPH_AUTOSIZE,
        margin=get_graph_margin(),
        showlegend=SHOW_STATE_LEGEND,
        hovermode=GRAPH_HOVER_MODE,
        dragmode=GRAPH_DRAG_MODE,
        legend={
            "title": {
                "text": LEGEND_TITLE,
            },
            "orientation": "h",
            "x": 0,
            "xanchor": "left",
            "y": 1.02,
            "yanchor": "bottom",
        },
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
        showticklabels=bool(
            BAR_ROW_NAME
        ),
        fixedrange=Y_AXIS_FIXED_RANGE,
    )

    return figure