# SPDX-License-Identifier: MIT
import logging
from abc import ABC, abstractmethod
from typing import Callable


class TaskBase(ABC):
    """
    実行タスクの抽象基底
    worker はこのインターフェースのみを扱う
    """

    @abstractmethod
    def run(
        self,
        runtime: dict,
        logger: logging.Logger,
        ui_call: Callable[[Callable], None],
        append_logs: Callable[[], None],
        update_status: Callable[[], None],
    ) -> None:
        pass
