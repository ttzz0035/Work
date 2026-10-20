
import os
import subprocess
import tempfile
import time
from pathlib import Path
from urllib import error

import yaml

from process_api import ProcessComposeAPI
from process_config import BASE, ProcessConfig


class ProcessService:
    """Process Composeの操作とライフサイクルを管理する。"""

    def __init__(self):
        self.config = ProcessConfig()
        self.api = ProcessComposeAPI(self.config)
        self.server = None

    def snapshot(self):
        """サーバーとNodeの状態を取得する。設定は変更しない。"""
        try:
            states = self.api.states()
            server = "Running"
        except error.URLError:
            states = {}
            server = "Stopped"

        nodes = {
            name: states.get(name, "Not Applied" if node.get("enabled", True) else "Disabled")
            for name, node in self.config.nodes.items()
        }

        return {"server": server, "nodes": nodes}

    def apply_missing(self):
        """未登録の有効Nodeを一括登録する。"""
        states = self.api.states()
        missing = [
            name for name, node in self.config.nodes.items()
            if node.get("enabled", True) and name not in states
        ]

        if not missing:
            return

        # 既存Nodeの登録情報を保持する。
        processes = {name: self.api.process_info(name) for name in states}

        # 未登録Nodeの実行設定を追加する。
        for name in missing:
            processes[name] = self.config.process_config(name)

        result = self.api.update_project(processes)

        # API適用後、Nodeが登録されたことを確認する。
        deadline = time.monotonic() + 5

        while time.monotonic() < deadline:
            current = self.api.states()
            pending = [name for name in missing if name not in current]

            if not pending:
                return

            time.sleep(0.2)

        raise RuntimeError(f"Project update failed: {pending}; response={result}")

    def save_cache(self):
        """サーバーの登録情報を取得してComposeキャッシュへ保存する。"""
        processes = {
            name: self.api.process_info(name)
            for name in self.api.states()
        }
        self.config.save_cache(processes)

    def start_server(self):
        """ComposeキャッシュからTUIを起動する。なければ空Projectで起動する。"""
        if self.server and self.server.poll() is None:
            raise RuntimeError("Managed server is already running")

        self.config.reload()

        if self.snapshot()["server"] == "Running":
            raise RuntimeError("Another Process Compose server is running")

        cache = self.config.cache_path()
        bootstrap = None

        # キャッシュが存在しない場合は、一時的な空Projectを用意する。
        if not cache.exists():
            with tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", dir=BASE, delete=False, encoding="utf-8") as file:
                yaml.safe_dump({"version": "0.5", "processes": {}}, file)
                bootstrap = Path(file.name)

        try:
            tui = self.config.cfg.get("tui", True)

            if not isinstance(tui, bool):
                raise ValueError("manager.tui must be boolean")

            command = [
                self.config.cfg.get("executable", "process-compose"),
                "up", "-f", str(bootstrap or cache),
                "--keep-project", "--no-watch",
            ]

            if not tui:
                command.append("-t=false")

            # WindowsではTUIを独立したコンソールで起動する。
            flags = (subprocess.CREATE_NEW_CONSOLE if tui else subprocess.CREATE_NO_WINDOW) if os.name == "nt" else 0
            streams = {} if tui else {
                "stdin": subprocess.DEVNULL,
                "stdout": subprocess.DEVNULL,
                "stderr": subprocess.DEVNULL,
            }

            self.server = subprocess.Popen(command, cwd=BASE, creationflags=flags, **streams)
            deadline = time.monotonic() + self.config.cfg.get("startup_timeout", 10)

            # TUI起動後、REST APIが応答可能になるまで待機する。
            while time.monotonic() < deadline:
                if self.server.poll() is not None:
                    raise RuntimeError(f"Server exited: {self.server.returncode}")

                try:
                    self.api.states()
                    return
                except error.URLError:
                    time.sleep(0.2)

            raise TimeoutError("Server startup timeout")

        finally:
            # 初回起動用の一時ファイルはキャッシュとして残さない。
            if bootstrap:
                bootstrap.unlink(missing_ok=True)

    def stop_server(self):
        """Project保存後、TUIと配下のWorkerを終了する。"""
        if not self.server or self.server.poll() is not None:
            raise RuntimeError("Managed server is not running")

        # 保存失敗時はサーバーを停止しない。
        self.save_cache()

        if os.name == "nt":
            result = subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(self.server.pid)],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )

            if result.returncode:
                raise RuntimeError(result.stderr or result.stdout)
        else:
            self.server.terminate()

        self.server.wait(timeout=10)
        self.server = None

    def node(self, action, name):
        """NodeをAPI経由で操作する。Composeキャッシュは変更しない。"""
        if name not in self.config.nodes:
            raise ValueError(f"Unknown node: {name}")

        if action not in ("start", "stop", "restart"):
            raise ValueError(f"Unknown action: {action}")

        if action != "stop" and not self.config.nodes[name].get("enabled", True):
            raise ValueError(f"Disabled node: {name}")

        states = self.api.states()

        # 起動・再起動時のみ、未登録Nodeを一括登録する。
        if name not in states:
            if action == "stop":
                return

            self.apply_missing()
            states = self.api.states()

        if name not in states:
            raise RuntimeError(f"Node not registered: {name}")

        state = states[name].strip().lower()
        active = {"running", "launching", "launched", "restarting"}

        if action == "stop":
            if state not in active:
                return
            self.api.stop_node(name)

        elif action == "start":
            if state == "running":
                return
            self.api.start_node(name)

        elif action == "restart":
            # 再起動前に対象Nodeの最新定義をサーバーへ適用する。
            processes = {key: self.api.process_info(key) for key in states}
            processes[name] = self.config.process_config(name)
            self.api.update_project(processes)

            # Project更新で自動起動した場合は二重起動を避ける。
            current = self.api.states().get(name, "").strip().lower()

            if current not in active:
                self.api.start_node(name)

    def execute(self, action, name=None):
        """GUIからの操作を実行し、画面表示用の状態を返す。"""
        if action == "refresh":
            self.config.reload()
        elif action == "server_start":
            self.start_server()
        elif action == "server_stop":
            self.stop_server()
        elif action == "server_restart":
            self.stop_server()
            self.start_server()
        elif action in ("start", "stop", "restart"):
            self.node(action, name)
        else:
            raise ValueError(f"Unknown action: {action}")

        return self.snapshot()
