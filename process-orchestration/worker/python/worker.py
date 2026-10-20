import argparse
from pathlib import Path

import yaml


class Worker:
    def __init__(self, config_path, key):
        self.config_path = Path(config_path).resolve()
        self.key = key
        self.config = self.load_config()

    def load_config(self):
        with self.config_path.open("r", encoding="utf-8") as file:
            data = yaml.safe_load(file)

        for part in self.key.split("."):
            if not isinstance(data, dict) or part not in data:
                raise KeyError(f"Configuration key not found: {self.key}")
            data = data[part]

        return data

    def run(self):
        print(f"[KEY] {self.key}", flush=True)
        print(f"[CONFIG] {self.config}", flush=True)

        # -----------------------
        # 業務処理はここに実装する。
        # -----------------------
        pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    parser.add_argument("--key", required=True)
    args = parser.parse_args()

    Worker(args.config, args.key).run()