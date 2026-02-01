# SPDX-License-Identifier: MIT
import time
import logging
from typing import Callable
from playwright.sync_api import sync_playwright

from task_base import TaskBase

HEADELESS = False


class TaskImpl(TaskBase):
    def __init__(
        self,
        runtime: dict,
        logger: logging.Logger,
        ui_call: Callable,
        append_logs: Callable[[], None],
        update_status: Callable[[], None],
    ):
        self.runtime = runtime
        self.logger = logger
        self.ui_call = ui_call
        self.append_logs = append_logs
        self.update_status = update_status

    def run(self) -> None:
        mode = self.runtime.get("mode")

        if mode == "register":
            self._run_register()
        elif mode == "verify":
            self._run_verify()
        else:
            self.logger.error(f"[TASK] unknown mode={mode}")
            self.ui_call(self.append_logs)

    # -------------------------------------------------
    # 登録処理
    # -------------------------------------------------
    def _run_register(self) -> None:
        self.logger.info("[TASK] register start")
        self.ui_call(self.append_logs)
        self.ui_call(self.update_status)

        print(self.runtime["start"], self.runtime["end"])

        with sync_playwright() as p:
            self.logger.info("[TASK] launch chromium")
            self.ui_call(self.append_logs)

            browser = p.chromium.launch(headless=HEADELESS)
            context = browser.new_context()
            page = context.new_page()

            page.goto("https://example.com")

            for i in range(15):
                if not self.runtime.get("running"):
                    break

                self.runtime["ticks"] += 1
                self.logger.info(f"[TASK] working {i+1}/15")
                self.ui_call(self.append_logs)
                self.ui_call(self.update_status)
                time.sleep(1)

            context.close()
            browser.close()

        self.logger.info("[TASK] register end")
        self.ui_call(self.append_logs)
        self.ui_call(self.update_status)

    # -------------------------------------------------
    # 照合処理
    # -------------------------------------------------
    def _run_verify(self) -> None:
        self.logger.info("[TASK] verify start")
        self.ui_call(self.append_logs)
        self.ui_call(self.update_status)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=HEADELESS)
            page = browser.new_page()
            page.goto("https://httpbin.org/get")

            for i in range(5):
                if not self.runtime.get("running"):
                    break

                self.runtime["ticks"] += 1
                self.logger.info(f"[TASK] verifying {i+1}/5")
                self.ui_call(self.append_logs)
                self.ui_call(self.update_status)
                time.sleep(1)

            browser.close()

        self.logger.info("[TASK] verify end")
        self.ui_call(self.append_logs)
        self.ui_call(self.update_status)
