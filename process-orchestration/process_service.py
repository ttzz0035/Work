
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

    ACTIVE = {"running", "pending", "launching", "launched", "restarting"}

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

    def wait_node(self, name, running, timeout=30):
        """Nodeが目的の状態になるまで待機する。"""
        deadline = time.monotonic() + timeout

        while time.monotonic() < deadline:
            state = self.api.states().get(name)

            if state is None:
                raise RuntimeError(f"Node not registered: {name}")

            normalized = state.strip().lower()

            if running and normalized == "running":
                return

            if not running and normalized not in self.ACTIVE:
                return

            if running and normalized in {"error", "failed", "completed"}:
                raise RuntimeError(f"Node startup failed: {name} ({state})")

            time.sleep(0.2)

        raise TimeoutError(f"Node state timeout: {name}, running={running}")

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

    def remove_missing(self):
        """YAMLから削除された停止済みNodeをProjectから削除する。"""
        states = self.api.states()
        removed = []

        for name, state in states.items():
            if name in self.config.nodes:
                continue

            # Running/Pendingなどの稼働中Nodeは削除しない。
            if state.strip().lower() in self.ACTIVE:
                print(f"[NODE SKIP] {name}: {state}", flush=True)
                continue

            removed.append(name)

        if not removed:
            return []

        # 削除対象以外の登録情報を保持する。
        processes = {
            name: self.api.process_info(name)
            for name in states
            if name not in removed
        }

        result = self.api.update_project(processes)

        # Project更新後、削除が反映されたことを確認する。
        deadline = time.monotonic() + 5

        while time.monotonic() < deadline:
            current = self.api.states()
            remaining = [name for name in removed if name in current]

            if not remaining:
                for name in removed:
                    print(f"[NODE REMOVE] {name}", flush=True)
                return removed

            time.sleep(0.2)

        raise RuntimeError(f"Node removal failed: {remaining}; response={result}")

    def sync_nodes(self):
        """Nodeの追加と削除を同期する。"""
        self.remove_missing()
        self.apply_missing()

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

        if action == "stop":
            if state not in self.ACTIVE:
                return
            self.api.stop_node(name)

        elif action == "start":
            if state in self.ACTIVE:
                return
            self.api.start_node(name)

        elif action == "restart":
            # 再起動前に対象Nodeの最新定義をサーバーへ適用する。
            processes = {key: self.api.process_info(key) for key in states}
            processes[name] = self.config.process_config(name)
            self.api.update_project(processes)

            # Project更新で自動起動した場合は二重起動を避ける。
            current = self.api.states().get(name, "").strip().lower()

            if current not in self.ACTIVE:
                self.api.start_node(name)

    def start_all(self):
        """有効NodeをYAML定義順に起動する。"""
        self.sync_nodes()

        for name, node in self.config.nodes.items():
            if not node.get("enabled", True):
                print(f"[NODE SKIP] {name}: Disabled", flush=True)
                continue

            state = self.api.states().get(name, "").strip().lower()

            if state == "running":
                print(f"[NODE SKIP] {name}: Running", flush=True)
                continue

            print(f"[NODE START] {name}", flush=True)

            # 個別操作と同じAPIを使用する。
            self.node("start", name)

            # Runningになるまで待ってから次のNodeを起動する。
            self.wait_node(name, running=True)
            print(f"[NODE RUNNING] {name}", flush=True)

    def stop_all(self):
        """登録済みNodeを逆順に停止する。"""
        states = self.api.states()

        # YAMLから削除済みのNodeも停止対象とする。
        names = list(states)

        for name in reversed(names):
            state = self.api.states().get(name, "").strip().lower()

            if state not in self.ACTIVE:
                print(f"[NODE SKIP] {name}: {state}", flush=True)
                continue

            print(f"[NODE STOP] {name}", flush=True)

            # YAMLから削除済みのNodeも停止できるようAPIを直接使用する。
            self.api.stop_node(name)

            # 停止完了を確認してから次のNodeを停止する。
            self.wait_node(name, running=False)
            print(f"[NODE STOPPED] {name}", flush=True)

        # 停止後、削除可能になったNodeを反映する。
        self.sync_nodes()

    def restart_all(self):
        """全Node停止後、YAML定義順に起動する。"""
        self.stop_all()
        self.start_all()

    def execute(self, action, name=None):
        """GUIからの操作を実行し、画面表示用の状態を返す。"""
        print(f"[SERVICE] action={action} node={name}", flush=True)

        if action == "refresh":
            self.config.reload()

            # サーバー稼働中のみNodeの追加・削除を同期する。
            if self.snapshot()["server"] == "Running":
                self.sync_nodes()

        elif action == "server_start":
            self.start_server()
            self.sync_nodes()

        elif action == "server_stop":
            self.stop_server()

        elif action == "server_restart":
            self.stop_server()
            self.start_server()
            self.sync_nodes()

        elif action in ("start", "stop", "restart"):
            self.config.reload()
            self.node(action, name)

            # 停止時には削除待ちNodeも同期する。
            if action == "stop":
                self.sync_nodes()

        elif action == "start_all":
            self.config.reload()
            self.start_all()

        elif action == "stop_all":
            self.config.reload()
            self.stop_all()

        elif action == "restart_all":
            self.config.reload()
            self.restart_all()

        else:
            raise ValueError(f"Unknown action: {action}")

        print(f"[SERVICE DONE] action={action}", flush=True)
        return self.snapshot()
