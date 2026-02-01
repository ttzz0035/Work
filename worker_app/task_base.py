# SPDX-License-Identifier: MIT
import logging
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Callable, Dict, Any


class TaskBase(ABC):
    def __init__(
        self,
        runtime: dict,
        ui_state: Dict[str, Any],
        logger: logging.Logger,
        ui_call: Callable,
        append_logs: Callable[[], None],
        update_status: Callable[[], None],
    ):
        self.runtime = runtime
        self.ui_state = ui_state
        self.logger = logger
        self.ui_call = ui_call
        self.append_logs = append_logs
        self.update_status = update_status

    # =================================================
    # UI 通知（低レベル）
    # =================================================
    def ui_log(self):
        self.ui_call(self.append_logs)

    def ui_status(self):
        self.ui_call(self.update_status)

    # =================================================
    # ★ 正規ログAPI（TaskImpl はこれだけ使う）
    # =================================================
    def log(self, message: str):
        """
        Task 用ログ出力（必須API）
        - logger に出す
        - UI に即反映
        """
        self.logger.info(message)
        self.ui_log()

    # =================================================
    # ★ 入力データ取得（UI 由来）
    # =================================================
    def get_input(self, key: str, default=None):
        return self.ui_state.get(key, default)

    # =================================================
    # ★ runtime データ取得（Worker 管理）
    # =================================================
    def get_runtime(self, key: str, default=None):
        return self.runtime.get(key, default)

    # =================================================
    # ★ 進捗ステップ（ticks / 時刻 / UI）
    # =================================================
    def step(self, message: str | None = None) -> None:
        self.runtime["ticks"] = self.runtime.get("ticks", 0) + 1
        self.runtime["last_tick_at"] = datetime.now()

        if message:
            self.logger.info(message)

        self.ui_log()
        self.ui_status()

    # =================================================
    # ★ 停止判定
    # =================================================
    def check_stop(self) -> bool:
        return not self.runtime.get("running", False)

    # =================================================
    # ★ 終了通知
    # =================================================
    def finish(self):
        self.logger.info("[TASK] finished")
        self.ui_log()
        self.ui_status()

    def fail(self, exc: Exception):
        self.logger.exception(f"[TASK] failed: {exc}")
        self.ui_log()
        self.ui_status()

    # =================================================
    # 実装強制
    # =================================================
    @abstractmethod
    def run(self) -> None:
        pass
