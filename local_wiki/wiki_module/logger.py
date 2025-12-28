from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional


class AppLogger:
    """
    ルートロガーは汚さない（propagate=False）。
    ファイル＋コンソールに出す。ローテーションあり。
    """

    def __init__(
        self,
        name: str,
        log_path: str = "logs/localwiki.log",
        level: str = "INFO",
        max_bytes: int = 2_000_000,
        backup_count: int = 3,
    ) -> None:
        self.name = name
        self.log_path = log_path
        self.level = level
        self.max_bytes = max_bytes
        self.backup_count = backup_count
        self._logger: Optional[logging.Logger] = None

    def get(self) -> logging.Logger:
        if self._logger is not None:
            return self._logger

        logger = logging.getLogger(self.name)
        logger.setLevel(self._parse_level(self.level))
        logger.propagate = False

        if not logger.handlers:
            Path(self.log_path).parent.mkdir(parents=True, exist_ok=True)

            fmt = logging.Formatter(
                fmt="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
            )

            fh = RotatingFileHandler(
                self.log_path,
                maxBytes=self.max_bytes,
                backupCount=self.backup_count,
                encoding="utf-8",
            )
            fh.setFormatter(fmt)
            fh.setLevel(self._parse_level(self.level))

            ch = logging.StreamHandler()
            ch.setFormatter(fmt)
            ch.setLevel(self._parse_level(self.level))

            logger.addHandler(fh)
            logger.addHandler(ch)

        self._logger = logger
        return logger

    def _parse_level(self, s: str) -> int:
        u = (s or "").strip().upper()
        if u == "DEBUG":
            return logging.DEBUG
        if u == "WARNING":
            return logging.WARNING
        if u == "ERROR":
            return logging.ERROR
        if u == "CRITICAL":
            return logging.CRITICAL
        return logging.INFO
