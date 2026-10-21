
"""Node Dashboard 共通UIログSchema。"""

import json
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, TypeAdapter, model_validator


class UIBase(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    version: Literal[1] = 1


class UILayout(UIBase):
    """Dashboardの列数。"""
    type: Literal["layout"] = "layout"
    columns: int = Field(ge=1, le=12)


class UICard(UIBase):
    """カード部品の共通フィールド。"""
    id: str = Field(min_length=1)
    label: str | None = None
    row: int | None = Field(default=None, ge=0)
    column: int | None = Field(default=None, ge=0)
    columnspan: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_position(self):
        # 実際の列数との整合性はDashboard側で検証する。
        return self


class UIValue(UICard):
    type: Literal["value"] = "value"
    value: str | int | float | bool | None = None
    unit: str | None = None


class UIStatus(UICard):
    type: Literal["status"] = "status"
    value: str | None = None


class UIProgress(UICard):
    type: Literal["progress"] = "progress"
    value: int | float | None = Field(default=None, ge=0, le=100)


UIMessage = Annotated[
    UILayout | UIValue | UIStatus | UIProgress,
    Field(discriminator="type"),
]
UI_ADAPTER = TypeAdapter(UIMessage)
UI_CLASSES = (UILayout, UIValue, UIStatus, UIProgress)
PREFIX = "[UI]"


def emit_ui(message: UIMessage):
    """SchemaインスタンスからUIログを出力する。"""
    if not isinstance(message, UI_CLASSES):
        raise TypeError("UI Schema instance required")
    data = message.model_dump(exclude_none=True)
    print(PREFIX + json.dumps(data, ensure_ascii=False, allow_nan=False), flush=True)


def parse_ui(line: str):
    """Process ComposeログからUIイベントを復元する。"""
    if not isinstance(line, str):
        return None

    try:
        wrapped = json.loads(line)
        if isinstance(wrapped, dict) and isinstance(wrapped.get("message"), str):
            line = wrapped["message"]
    except json.JSONDecodeError:
        pass

    index = line.find(PREFIX)
    if index < 0:
        return None

    return UI_ADAPTER.validate_json(line[index + len(PREFIX):].strip())


def json_schema():
    """共通JSON Schemaを取得する。"""
    return UI_ADAPTER.json_schema()
