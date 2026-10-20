
import os
import re
import shlex
import subprocess
from pathlib import Path

import yaml


BASE = Path(__file__).resolve().parent
CONFIG = BASE / "nodes.yaml"


class ProcessConfig:
    """Node設定とProcess Composeキャッシュを管理する。"""

    def __init__(self):
        self.cfg = {}
        self.nodes = {}
        self.reload()

    def reload(self):
        """nodes.yamlを読み込み、設定の妥当性を検証する。"""
        data = yaml.safe_load(CONFIG.read_text(encoding="utf-8"))

        if not isinstance(data, dict) or not isinstance(data.get("nodes"), dict):
            raise ValueError("Invalid nodes.yaml")

        cfg = data.get("manager", {})
        nodes = data["nodes"]

        if not isinstance(cfg, dict):
            raise ValueError("Invalid manager settings")

        for name, node in nodes.items():
            if not isinstance(name, str) or not re.fullmatch(r"[A-Za-z0-9_-]+", name):
                raise ValueError(f"Invalid node name: {name}")

            if not isinstance(node, dict):
                raise ValueError(f"Invalid node: {name}")

            executable = node.get("executable")
            if not isinstance(executable, str) or not re.fullmatch(r"[A-Za-z0-9_.-]+", executable):
                raise ValueError(f"Invalid executable: {name}")

            if not isinstance(node.get("file"), str) or not node["file"]:
                raise ValueError(f"Invalid worker file: {name}")

            if not isinstance(node.get("enabled", True), bool):
                raise ValueError(f"Invalid enabled setting: {name}")

            if not isinstance(node.get("parameters", {}), dict):
                raise ValueError(f"Invalid parameters: {name}")

            if node.get("enabled", True):
                self.worker_path(node)

        self.cfg = cfg
        self.nodes = nodes

    def worker_path(self, node):
        """Workerファイルを解決し、指定ディレクトリ外への参照を防ぐ。"""
        directory = (BASE / "worker" / node["executable"]).resolve()
        path = (directory / node["file"]).resolve()

        if not path.is_relative_to(directory) or not path.is_file():
            raise ValueError(f"Invalid worker path: {path}")

        return path

    def cache_path(self):
        """Composeキャッシュの保存先を取得する。"""
        return (BASE / self.cfg.get("cache", "process-compose.yaml")).resolve()

    def process_config(self, name):
        """Node定義からREST API登録用の設定を生成する。"""
        node = self.nodes[name]

        args = [
            node["executable"],
            str(self.worker_path(node)),
            "--config",
            str(CONFIG),
            "--key",
            f"nodes.{name}.parameters",
        ]

        command = subprocess.list2cmdline(args) if os.name == "nt" else shlex.join(args)

        if os.name == "nt":
            executable = "cmd"
            arguments = ["/C", command]
        else:
            executable = "sh"
            arguments = ["-c", command]

        return {
            "name": name,
            "replicaName": name,
            "command": command,
            "executable": executable,
            "args": arguments,
            "workingDir": str(BASE),
            "disabled": True,
        }

    def save_cache(self, processes):
        """取得済みのProject定義を最小限のCompose YAMLとして保存する。"""
        result = {}

        for name, info in processes.items():
            if not isinstance(info, dict) or not info.get("command"):
                raise RuntimeError(f"Invalid process configuration: {name}")

            result[name] = {
                "command": info["command"],
                "working_dir": info.get("workingDir", str(BASE)),
                "disabled": info.get("disabled", True),
            }

        path = self.cache_path()
        path.parent.mkdir(parents=True, exist_ok=True)

        temp = path.with_suffix(path.suffix + ".tmp")

        try:
            temp.write_text(
                yaml.safe_dump(
                    {"version": "0.5", "processes": result},
                    sort_keys=False,
                    allow_unicode=True,
                ),
                encoding="utf-8",
            )
            os.replace(temp, path)
        finally:
            temp.unlink(missing_ok=True)
