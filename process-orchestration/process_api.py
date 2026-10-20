
import json
import os
from urllib import error, request
from urllib.parse import quote


class ProcessComposeAPI:
    """Process Compose REST APIへの通信を担当する。"""

    def __init__(self, config):
        # 接続設定はProcessConfigから参照する。
        self.config = config

    def request(self, method, path, payload=None):
        """HTTPリクエストを実行し、レスポンスをJSONとして返す。"""
        headers = {}

        # APIトークン認証が設定されている場合に使用する。
        if token := os.getenv("PC_API_TOKEN"):
            headers["X-PC-Token-Key"] = token

        body = None

        # POST/PATCHなどのJSONリクエストを生成する。
        if payload is not None:
            headers["Content-Type"] = "application/json"
            body = json.dumps(payload).encode("utf-8")

        url = self.config.cfg.get("api", "http://localhost:8080").rstrip("/") + path
        req = request.Request(url, data=body, headers=headers, method=method)

        try:
            with request.urlopen(req, timeout=self.config.cfg.get("request_timeout", 3)) as response:
                raw = response.read()
                return json.loads(raw) if raw else {}

        except error.HTTPError as exc:
            message = exc.read().decode(errors="replace")
            raise RuntimeError(f"HTTP {exc.code}: {message}") from exc

    def states(self):
        """全Nodeの名前と実行状態を取得する。"""
        result = self.request("GET", "/processes")

        return {
            process["name"]: process["status"]
            for process in result["data"]
            if process.get("name")
        }

    def process_info(self, name):
        """指定Nodeの登録済み設定を取得する。"""
        return self.request("GET", f"/process/info/{quote(name, safe='')}")

    def update_project(self, processes):
        """複数Nodeの定義をProjectへ一括適用する。"""
        return self.request("POST", "/project", {
            "version": "0.5",
            "processes": processes,
        })

    def start_node(self, name):
        """登録済みNodeを起動する。"""
        return self.request("POST", f"/process/start/{quote(name, safe='')}")

    def stop_node(self, name):
        """実行中Nodeを停止する。"""
        return self.request("PATCH", f"/process/stop/{quote(name, safe='')}")

    def restart_node(self, name):
        """登録済みNodeを再起動する。"""
        return self.request("POST", f"/process/restart/{quote(name, safe='')}")
