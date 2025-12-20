from dataclasses import dataclass
from typing import Optional, Literal


NodeKind = Literal["folder", "file", "sheet"]


@dataclass(frozen=True)
class NodeTag:
    kind: NodeKind
    path: str
    sheet: Optional[str] = None

    def __str__(self) -> str:
        if self.kind == "sheet":
            return f"{self.path}::{self.sheet}"
        return self.path
