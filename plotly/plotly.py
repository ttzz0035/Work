import pandas as pd
import plotly.express as px

df = pd.DataFrame(
    [
        {"Task": "要件整理", "Start": "2026-08-01", "Finish": "2026-08-05"},
        {"Task": "基本設計", "Start": "2026-08-04", "Finish": "2026-08-12"},
        {"Task": "実装", "Start": "2026-08-10", "Finish": "2026-08-25"},
        {"Task": "試験", "Start": "2026-08-20", "Finish": "2026-08-31"},
    ]
)

fig = px.timeline(
    df,
    x_start="Start",
    x_end="Finish",
    y="Task",
    title="開発スケジュール",
)

fig.update_yaxes(autorange="reversed")
fig.show()