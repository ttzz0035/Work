
class Task:
    """Workerから呼び出される業務処理。単独デバッグにも対応する。"""

    def __init__(self):
        self.running = False

    def run(self, config):
        """Workerと単独デバッグで共通の実行入口。"""
        self.running = True

        try:
            print(f"[TASK CONFIG] {config}", flush=True)

            # -----------------------
            # 業務処理はここに実装する。
            # -----------------------
            print(f"Mode: {config.get('mode')}", flush=True)
            print(f"Interval: {config.get('interval')}", flush=True)

        finally:
            self.running = False


if __name__ == "__main__":
    # VS CodeからF5で単独デバッグする。
    # WorkerやProcess Composeの起動は不要。
    DEBUG_CONFIG = {
        "task": "sample",
        "interval": 1000,
        "mode": "collect",
    }

    Task().run(DEBUG_CONFIG)
