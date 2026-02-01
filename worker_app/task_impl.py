# SPDX-License-Identifier: MIT
import time
from task_base import TaskBase


class TaskImpl(TaskBase):
    def run(self) -> None:
        # -------------------------------
        # UI → Task 入力（Base経由）
        # -------------------------------
        mode = self.get_input("mode")
        job_id = self.get_input("job_id")
        start = self.get_input("start_date")
        end = self.get_input("end_date")

        print("=== TASK INPUT (via TaskBase) ===")
        print("mode :", mode)
        print("job  :", job_id)
        print("start:", start)
        print("end  :", end)
        print("================================")

        # Base が提供する正規ログAPI
        self.log(
            f"[TASK] start mode={mode} job={job_id} "
            f"period={start}～{end}"
        )

        try:
            if mode == "register":
                self._run_register()
            elif mode == "verify":
                self._run_verify()
            else:
                raise ValueError(f"unknown mode={mode}")

            self.finish()

        except Exception as e:
            self.fail(e)

    # -------------------------------
    # 登録処理（ダミー）
    # -------------------------------
    def _run_register(self):
        self.log("[TASK] register start")
        ##############################
        # データロード
        #
        #
        ##############################

        # データのサイズを実装する
        total = len(range(10))

        for i in range(total):
            if self.check_stop():
                self.log("[TASK] stop requested")
                return
            
            ###############################
            # ここに実装
            #
            #
            #
            #
            #
            ###############################


            # ticks++, last_tick_at 更新、UI通知を全部 Base がやる
            self.step(f"[TASK] register step {i + 1}/{total}")
            time.sleep(0.5)

        self.log("[TASK] register end")

    # -------------------------------
    # 照合処理（ダミー）
    # -------------------------------
    def _run_verify(self):
        self.log("[TASK] verify start")
        ##############################
        # データロード
        #
        #
        ##############################

        # データのサイズを実装する
        total = len(range(10))

        for i in range(total):
            if self.check_stop():
                self.log("[TASK] stop requested")
                return

            ###############################
            # ここに実装
            #
            #
            #
            #
            #
            ###############################

            self.step(f"[TASK] verify step {i + 1}/{total}")
            time.sleep(0.5)

        self.log("[TASK] verify end")

if __name__ == "__main__":
    import logging

    # ダミー runtime / ui_state
    runtime = {
        "running": True,
        "ticks": 0,
    }

    ui_state = {
        "mode": "register",
        "job_id": 123,
        "start_date": "2026/02/01",
        "end_date": "2026/02/01",
    }

    # ダミー UI コール
    def ui_call(fn): fn()
    def append_logs(): print("  [UI] log updated")
    def update_status(): print("  [UI] status updated")

    logger = logging.getLogger("TASK")
    logger.setLevel(logging.INFO)
    logger.addHandler(logging.StreamHandler())

    task = TaskImpl(
        runtime=runtime,
        ui_state=ui_state,
        logger=logger,
        ui_call=ui_call,
        append_logs=append_logs,
        update_status=update_status,
    )

    task.run()

    print("final ticks =", runtime["ticks"])
