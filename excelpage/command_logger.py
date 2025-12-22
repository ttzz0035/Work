from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, List

from Logger import Logger
logger = Logger(
    name="CommandLogger",
    log_file_path="logs/app.log",
    level="DEBUG",
)

@dataclass(frozen=True)
class CommandRecord:
    ts: str
    command: str
    args: List[str]


class CommandLogger:
    """
    操作をコマンド形式で記録し、復元（リプレイ）可能にする。
    - 追記のみ（append-only）
    - 1行 = 1コマンド
    - 先頭にタイムスタンプ
    """
    def __init__(self, path: str = "logs/commands.log"):
        self.path = path
        os.makedirs(os.path.dirname(self.path), exist_ok=True)

    def write(self, command: str, *args: str) -> None:
        try:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            line = f"{ts} {command}"
            for a in args:
                line += " " + self._quote(str(a))
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
            logger.info("[CMD] %s", line)
        except Exception as ex:
            logger.error("[CMD] write failed: %s", ex, exc_info=True)

    def read(self) -> List[CommandRecord]:
        records: List[CommandRecord] = []
        if not os.path.exists(self.path):
            return records
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                for raw in f:
                    raw = raw.strip()
                    if not raw:
                        continue
                    rec = self._parse_line(raw)
                    if rec:
                        records.append(rec)
        except Exception as ex:
            logger.error("[CMD] read failed: %s", ex, exc_info=True)
        return records

    def _quote(self, s: str) -> str:
        # 空、スペース、引用符を含むなら "..." で囲う
        if s == "" or (" " in s) or ('"' in s) or ("\t" in s):
            return '"' + s.replace('"', '\\"') + '"'
        return s

    def _parse_line(self, line: str) -> CommandRecord | None:
        # 形式: "YYYY-mm-dd HH:MM:SS COMMAND arg..."
        # 引数はダブルクォート対応
        try:
            if len(line) < 20:
                return None
            ts = line[:19]
            rest = line[20:].strip()
            if not rest:
                return None
            parts = self._split_args(rest)
            if not parts:
                return None
            command = parts[0]
            args = parts[1:]
            return CommandRecord(ts=ts, command=command, args=args)
        except Exception:
            logger.error("[CMD] parse failed: %s", line, exc_info=True)
            return None

    def _split_args(self, s: str) -> List[str]:
        # シンプルなシェル風split（"..." と \" を対応）
        out: List[str] = []
        buf: List[str] = []
        in_q = False
        i = 0
        while i < len(s):
            ch = s[i]
            if in_q:
                if ch == "\\" and i + 1 < len(s) and s[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                if ch == '"':
                    in_q = False
                    i += 1
                    continue
                buf.append(ch)
                i += 1
                continue
            else:
                if ch.isspace():
                    if buf:
                        out.append("".join(buf))
                        buf.clear()
                    i += 1
                    continue
                if ch == '"':
                    in_q = True
                    i += 1
                    continue
                buf.append(ch)
                i += 1
                continue
        if buf:
            out.append("".join(buf))
        return out
