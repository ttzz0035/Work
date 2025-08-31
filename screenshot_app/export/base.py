# export/base.py
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Protocol, Iterable, Callable, Any

from core.model import ExportBundle, load_bundle_from_folder

DEFAULT_TITLE = "Captures Export"

@dataclass
class ExportOptions:
    title: Optional[str] = None        # タブ名 / シート名
    filename: Optional[Path] = None    # 出力ファイル名（拡張子は各Exporter側で付与）

class Exporter(Protocol):
    name: str
    ext: str
    def export_bundle(self, bundle: ExportBundle, options: ExportOptions) -> Path: ...
    def export(self, folder: Path, options: ExportOptions) -> Path: ...

class BaseExporter:
    """共通抽象基底。共通の“既存タブ/シート削除”ヘルパを提供"""
    name: str = ""
    ext: str = ""

    def export_bundle(self, bundle: ExportBundle, options: ExportOptions) -> Path:
        raise NotImplementedError

    def export(self, folder: Path, options: ExportOptions) -> Path:
        bundle = load_bundle_from_folder(folder, title=(options.title or DEFAULT_TITLE))
        return self.export_bundle(bundle, options)

    @staticmethod
    def remove_existing_by_title(
        targets: Iterable[Any],
        match_title: str,
        get_name: Callable[[Any], str],
        delete: Callable[[Any], None],
    ) -> int:
        """
        既存タブ/シートをタイトル一致で削除（共通化）。
        returns: 削除件数
        """
        cnt = 0
        for t in list(targets):
            try:
                if get_name(t) == match_title:
                    delete(t)
                    cnt += 1
            except Exception:
                continue
        return cnt
