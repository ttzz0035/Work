# Common/filelist.py
from __future__ import annotations
from pathlib import Path
import json
from typing import Iterable, List, Optional

__all__ = ["FileListManager", "DEFAULT_CFG_DIR", "DEFAULT_CFG_PATH"]

APP_NAME = "File Selector"
DEFAULT_CFG_DIR = Path.home() / ".file_selector"
DEFAULT_CFG_PATH = DEFAULT_CFG_DIR / "config.json"


class FileListManager:
    """
    複数ファイル選択ロジック（UI非依存）
    - 重複排除
    - 存在チェック
    - 拡張子フィルタ（任意・デフォルト無制限）
    - last_dir の永続化（config.json）
    """
    def __init__(
        self,
        allowed_exts: Optional[Iterable[str]] = None,  # None=全許可
        cfg_path: Path = DEFAULT_CFG_PATH,
        persist: bool = True,
    ) -> None:
        self._files: List[Path] = []
        self.allowed_exts = tuple(allowed_exts) if allowed_exts else None
        self.cfg_path = cfg_path
        self.persist = persist
        self._cfg = self._load_config() if persist else {}
        self._last_dir = Path(self._cfg.get("last_dir", str(Path.home())))

    # ---------- Public API ----------
    def add(self, paths: Iterable[Path]) -> int:
        before = len(self._files)
        seen = {p.resolve() for p in self._files}

        added_any = False
        for raw in paths:
            p = Path(raw).resolve()
            if not p.exists() or not p.is_file():
                continue
            if self.allowed_exts and p.suffix not in self.allowed_exts:
                continue
            if p in seen:
                continue
            self._files.append(p)
            seen.add(p)
            added_any = True

        if added_any:
            try:
                self._last_dir = self._files[-1].parent
                self._persist()
            except Exception:
                pass

        return len(self._files) - before

    def remove_by_indices(self, indices: Iterable[int]) -> int:
        idxs = sorted({int(i) for i in indices if i is not None}, reverse=True)
        removed = 0
        for i in idxs:
            if 0 <= i < len(self._files):
                del self._files[i]
                removed += 1
        if removed:
            self._persist()
        return removed

    def clear(self) -> None:
        self._files.clear()
        self._persist()

    def list(self) -> List[Path]:
        return list(self._files)

    def count(self) -> int:
        return len(self._files)

    def last_dir(self) -> Path:
        return Path(self._last_dir)

    def set_last_dir(self, path: Path) -> None:
        p = Path(path).resolve()
        self._last_dir = p if p.exists() and p.is_dir() else Path.home()
        self._persist()

    # ---------- Internal ----------
    def _load_config(self) -> dict:
        if self.cfg_path.exists():
            try:
                return json.loads(self.cfg_path.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def _persist(self) -> None:
        if not self.persist:
            return
        try:
            DEFAULT_CFG_DIR.mkdir(parents=True, exist_ok=True)
            cfg = dict(self._cfg)
            cfg["last_dir"] = str(self._last_dir)
            self.cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
