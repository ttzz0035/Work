
import argparse
import importlib.util
import re
import sys
from pathlib import Path

import yaml


BASE = Path(__file__).resolve().parents[2]
TASKS_DIR = BASE / "tasks"

# プロジェクト共通モジュールをTaskから参照できるようにする。
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

class Worker:
    """Node設定に従ってTaskを読み込み、実行する。"""

    def __init__(self, config_path=None, key=None, config=None, task=None):
        self.config_path = Path(config_path).resolve() if config_path else None
        self.key = key

        # Compose起動とPython直接起動の設定方法を分離する。
        if config is not None:
            if config_path is not None or key is not None:
                raise ValueError("Specify config or config_path/key, not both")
            if not isinstance(config, dict) or not isinstance(task, str) or not task:
                raise ValueError("config and task are required")
            self.config = config
            self.task_name = task
        else:
            if self.config_path is None or not self.key:
                raise ValueError("config_path and key are required")
            self.config = self.load_config()

    def load_config(self):
        """Node直下のtaskとparametersを読み込む。"""
        with self.config_path.open("r", encoding="utf-8") as file:
            data = yaml.safe_load(file)

        # 既存のnodes.XXX.parametersから親Nodeを特定する。
        parts = self.key.split(".")
        if len(parts) != 3 or parts[0] != "nodes" or parts[2] != "parameters":
            raise ValueError(f"Invalid configuration key: {self.key}")

        for part in parts[:2]:
            if not isinstance(data, dict) or part not in data:
                raise KeyError(f"Configuration key not found: {self.key}")
            data = data[part]

        if not isinstance(data, dict):
            raise ValueError(f"Invalid node configuration: {self.key}")

        self.task_name = data["task"]
        parameters = data["parameters"]

        if not isinstance(parameters, dict):
            raise ValueError(f"Invalid parameters: {self.key}")

        return parameters

    def load_task(self, task_name):
        """tasks/<task名>/task.pyからTaskクラスを読み込む。"""
        if not isinstance(task_name, str) or not re.fullmatch(r"[A-Za-z0-9_-]+", task_name):
            raise ValueError(f"Invalid task name: {task_name}")

        task_file = TASKS_DIR / task_name / "task.py"

        if not task_file.is_file():
            raise FileNotFoundError(f"Task not found: {task_file}")

        # Taskモジュールを動的ロードする。
        spec = importlib.util.spec_from_file_location(f"worker_task_{task_name}", task_file)

        if spec is None or spec.loader is None:
            raise ImportError(f"Cannot load task: {task_name}")

        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)

        task_class = getattr(module, "Task", None)

        if not isinstance(task_class, type):
            raise TypeError(f"Task class not found: {task_file}")

        return task_class()

    def run(self):
        """設定されたTaskを実行する。"""
        print(f"[KEY] {self.key}", flush=True)
        print(f"[TASK] {self.task_name}", flush=True)
        print(f"[CONFIG] {self.config}", flush=True)

        task = self.load_task(self.task_name)

        print(f"[TASK START] {self.task_name}", flush=True)
        task.run(self.config)
        print(f"[TASK END] {self.task_name}", flush=True)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Process Composeからの既存引数起動。
        parser = argparse.ArgumentParser()
        parser.add_argument("--config", required=True)
        parser.add_argument("--key", required=True)
        args = parser.parse_args()

        worker = Worker(config_path=args.config, key=args.key)

    else:
        # VS Codeから引数なしで直接デバッグする。
        DEBUG_TASK = "sample"
        DEBUG_CONFIG = {
            "interval": 1000,
            "mode": "collect",
        }

        worker = Worker(config=DEBUG_CONFIG, task=DEBUG_TASK)

    worker.run()
