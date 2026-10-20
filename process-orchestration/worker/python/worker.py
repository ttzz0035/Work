
import argparse
import importlib.util
import re
import sys
from pathlib import Path

import yaml


# プロジェクトのルートディレクトリ。
BASE = Path(__file__).resolve().parents[2]
TASKS_DIR = BASE / "tasks"


class Worker:
    """設定の取得とTaskの読み込み・実行を担当する。"""

    def __init__(self, config_path=None, key=None, config=None):
        self.config_path = Path(config_path).resolve() if config_path else None
        self.key = key

        # 通常実行ではYAMLを読み込み、デバッグでは辞書を直接受け取る。
        if config is not None:
            if config_path is not None or key is not None:
                raise ValueError("Specify config or config_path/key, not both")
            if not isinstance(config, dict):
                raise TypeError("config must be a dictionary")
            self.config = config
        else:
            if self.config_path is None or not self.key:
                raise ValueError("config_path and key are required")
            self.config = self.load_config()

    def load_config(self):
        """Composeから指定されたYAMLの設定キーを読み出す。"""
        with self.config_path.open("r", encoding="utf-8") as file:
            data = yaml.safe_load(file)

        for part in self.key.split("."):
            if not isinstance(data, dict) or part not in data:
                raise KeyError(f"Configuration key not found: {self.key}")
            data = data[part]

        if not isinstance(data, dict):
            raise ValueError(f"Configuration must be a dictionary: {self.key}")

        return data

    def load_task(self, task_name):
        """tasks/<task_name>/task.pyからTaskクラスを読み込む。"""
        if not isinstance(task_name, str) or not re.fullmatch(r"[A-Za-z0-9_-]+", task_name):
            raise ValueError(f"Invalid task name: {task_name}")

        task_file = TASKS_DIR / task_name / "task.py"

        if not task_file.is_file():
            raise FileNotFoundError(f"Task not found: {task_file}")

        # Taskファイルを動的に読み込む。
        spec = importlib.util.spec_from_file_location(
            f"worker_task_{task_name}",
            task_file,
        )

        if spec is None or spec.loader is None:
            raise ImportError(f"Cannot load task: {task_name}")

        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)

        # 全Taskの公開インターフェースはTask.run(config)に統一する。
        task_class = getattr(module, "Task", None)

        if not isinstance(task_class, type):
            raise TypeError(f"Task class not found: {task_file}")

        return task_class()

    def run(self):
        """通常実行と単独デバッグで共通のTask実行処理。"""
        print(f"[KEY] {self.key}", flush=True)
        print(f"[CONFIG] {self.config}", flush=True)

        # 設定から実行するTaskを決定する。
        task_name = self.config["task"]
        task = self.load_task(task_name)

        print(f"[TASK START] {task_name}", flush=True)

        # Task側に業務処理を委譲する。
        task.run(self.config)

        print(f"[TASK END] {task_name}", flush=True)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Process Composeから起動する場合。
        # python worker.py --config nodes.yaml --key nodes.XSASA.parameters
        parser = argparse.ArgumentParser()
        parser.add_argument("--config", required=True)
        parser.add_argument("--key", required=True)
        args = parser.parse_args()

        worker = Worker(config_path=args.config, key=args.key)

    else:
        # VS CodeからF5で起動する場合。
        # コマンドライン引数やProcess Composeは不要。
        DEBUG_CONFIG = {
            "task": "sample",
            "interval": 1000,
            "mode": "collect",
        }

        worker = Worker(config=DEBUG_CONFIG)

    worker.run()
