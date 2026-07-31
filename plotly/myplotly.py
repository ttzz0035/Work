from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px


# ========================================
# 入力データ
# A → B → A → C の順番で積み上げる
# duration_days は各区間の日数
# ========================================
stack_items = [
    {"category": "A", "duration_days": 3},
    {"category": "B", "duration_days": 2},
    {"category": "A", "duration_days": 4},
    {"category": "C", "duration_days": 3},
    {"category": "B", "duration_days": 2},
]


# ========================================
# 開始日
# ========================================
current_start = datetime(2026, 8, 1)


# ========================================
# Timeline用データを生成
# ========================================
timeline_rows = []

for index, item in enumerate(stack_items, start=1):
    start = current_start
    finish = start + timedelta(days=item["duration_days"])

    timeline_rows.append(
        {
            "Row": "工程",
            "Category": item["category"],
            "Start": start,
            "Finish": finish,
            "Duration": item["duration_days"],
            "Sequence": index,
            "Label": item["category"],
        }
    )

    # 次の区間は直前の終了日時から開始する
    current_start = finish


df = pd.DataFrame(timeline_rows)


# ========================================
# ガントチャート生成
# ========================================
fig = px.timeline(
    df,
    x_start="Start",
    x_end="Finish",
    y="Row",
    color="Category",
    text="Label",
    hover_data={
        "Row": False,
        "Category": True,
        "Start": "|%Y-%m-%d",
        "Finish": "|%Y-%m-%d",
        "Duration": True,
        "Sequence": True,
        "Label": False,
    },
    color_discrete_map={
        "A": "#4C78A8",
        "B": "#F58518",
        "C": "#54A24B",
    },
    title="A → B → A → C の積み上げガントチャート",
)


# ========================================
# 表示設定
# ========================================
fig.update_traces(
    textposition="inside",
    insidetextanchor="middle",
    textfont={
        "size": 16,
        "color": "white",
    },
    marker={
        "line": {
            "color": "white",
            "width": 1,
        }
    },
)


fig.update_layout(
    xaxis_title="日付",
    yaxis_title="",
    legend_title="区分",
    height=350,
    bargap=0.15,
    hoverlabel={
        "namelength": -1,
    },
)


fig.update_xaxes(
    tickformat="%Y-%m-%d",
    showgrid=True,
)


fig.update_yaxes(
    showgrid=False,
)


fig.show()