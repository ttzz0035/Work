
"""Node Dashboard 共通UIログSchema。"""

import json
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, TypeAdapter, model_validator


class UIBase(BaseModel):
    """すべてのUIログに共通するSchema。"""

    model_config = ConfigDict(extra="forbid", strict=True)
    version: Literal[1] = 1


class UILayout(UIBase):
    """Dashboardのグリッド列数を定義する。"""

    type: Literal["layout"] = "layout"
    columns: int = Field(ge=1, le=12)


class UICard(UIBase):
    """カードの識別子・表示名・グリッド配置を定義する。"""

    id: str = Field(min_length=1)
    label: str | None = None
    row: int | None = Field(default=None, ge=0)
    column: int | None = Field(default=None, ge=0)
    columnspan: int | None = Field(default=None, ge=1)


class UIValue(UICard):
    """数値・文字列・状態を表示するカード。"""

    type: Literal["value"] = "value"
    value: str | int | float | bool | None = None
    unit: str | None = None
    color: str | None = Field(default=None, pattern=r"^#[0-9A-Fa-f]{6}$")


class UIProgress(UICard):
    """通常の進捗、または基準値を中央とする偏差を表示するカード。"""

    type: Literal["progress"] = "progress"
    value: int | float | None = Field(default=None, ge=0)
    reference: int | float | None = Field(default=None, ge=0)
    range: tuple[float, float] | None = None
    unit: str | None = None

    @model_validator(mode="after")
    def validate_progress(self):
        # 通常モードで完結するメッセージは0〜100%に制限する。
        # 値だけの更新はDashboardが既存カードのモードに基づいて検証する。
        if self.reference is None and self.range is None and (
            self.label is not None or self.row is not None or self.column is not None
        ):
            if self.value is not None and self.value > 100:
                raise ValueError("Normal progress must be between 0 and 100")

        # 基準値より小さい側と大きい側の表示範囲を検証する。
        if self.range is not None:
            if self.range[0] >= 0 or self.range[1] <= 0:
                raise ValueError("range must be (negative, positive)")

        return self
    
class UILamp(UICard):
    """カラーコードと任意の文字で状態を表示するカード。"""

    type: Literal["lamp"] = "lamp"
    color: str | None = Field(default=None, pattern=r"^#[0-9A-Fa-f]{6}$")
    text: str | None = None


# typeフィールドでメッセージのSchemaを判別する。
UIMessage = Annotated[
    UILayout | UIValue | UIProgress | UILamp,
    Field(discriminator="type"),
]

UI_ADAPTER = TypeAdapter(UIMessage)
UI_CLASSES = (UILayout, UIValue, UIProgress, UILamp)
PREFIX = "[UI]"


def emit_ui(message: UIMessage):
    """SchemaインスタンスからUIログを標準出力へ送信する。"""

    # Taskから未定義のUIメッセージを送信させない。
    if not isinstance(message, UI_CLASSES):
        raise TypeError("UI Schema instance required")

    # 出力時にもSchema検証を実施する。
    data = UI_ADAPTER.validate_python(message).model_dump(exclude_none=True)
    print(PREFIX + json.dumps(data, ensure_ascii=False, allow_nan=False), flush=True)


def parse_ui(line: str):
    """Process ComposeのログからUIメッセージを復元する。"""

    if not isinstance(line, str):
        return None

    # Process ComposeがJSONでログを包んでいる場合はmessageを取得する。
    try:
        wrapped = json.loads(line)
        if isinstance(wrapped, dict) and isinstance(wrapped.get("message"), str):
            line = wrapped["message"]
    except json.JSONDecodeError:
        pass

    # 通常ログはUI処理の対象外とする。
    index = line.find(PREFIX)
    if index < 0:
        return None

    # 受信したJSONも共通Schemaで検証する。
    return UI_ADAPTER.validate_json(line[index + len(PREFIX):].strip())


def json_schema():
    """他言語のTaskでも利用できるJSON Schemaを取得する。"""

    return UI_ADAPTER.json_schema()
