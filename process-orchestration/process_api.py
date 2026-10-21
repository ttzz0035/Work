
import json
import os
from urllib import error, request
from urllib.parse import quote


class ProcessComposeAPI:
    """Process ComposeのREST API通信のみを担当する。"""

    def __init__(self, config):
        # API接続先やタイムアウトなどの設定を保持する。
        self.config = config

    def request(self, method, path, payload=None):
        """指定されたHTTPメソッドでProcess Compose APIを呼び出す。"""

        # 環境変数にAPIトークンがある場合のみ認証ヘッダーを設定する。
        headers = {}
        if token := os.getenv("PC_API_TOKEN"):
            headers["X-PC-Token-Key"] = token

        # リクエストデータがある場合はJSON形式に変換する。
        body = json.dumps(payload).encode("utf-8") if payload is not None else None
        if body is not None:
            headers["Content-Type"] = "application/json"

        # 設定されたAPI接続先とエンドポイントからURLを生成する。
        url = self.config.cfg.get("api", "http://localhost:8080").rstrip("/") + path
        req = request.Request(url, data=body, headers=headers, method=method)

        try:
            # APIを呼び出し、JSONレスポンスをPythonオブジェクトへ変換する。
            with request.urlopen(req, timeout=self.config.cfg.get("request_timeout", 3)) as response:
                raw = response.read()
                return json.loads(raw) if raw else {}
        except error.HTTPError as exc:
            # HTTPエラーにはステータスコードとレスポンス本文を含める。
            raise RuntimeError(f"HTTP {exc.code}: {exc.read().decode(errors='replace')}") from exc

    def states(self):
        """登録されている全Nodeの実行状態を取得する。"""

        # プロセス名をキー、実行状態を値とする辞書を返す。
        result = self.request("GET", "/processes")
        return {p["name"]: p["status"] for p in result["data"] if p.get("name")}

    def process_info(self, name):
        """指定Nodeの詳細情報を取得する。"""
        return self.request("GET", f"/process/info/{quote(name, safe='')}")

    def process_logs(self, name):
        """指定Nodeのログを取得する。"""

        # UI表示用ログも通常ログと同じエンドポイントから取得する。
        result = self.request("GET", f"/process/logs/{quote(name, safe='')}/0/0")
        return result["logs"]

    def update_project(self, processes):
        """Process Composeのプロジェクト定義を更新する。"""
        return self.request("POST", "/project", {"version": "0.5", "processes": processes})

    def start_node(self, name):
        """指定Nodeを起動する。"""
        return self.request("POST", f"/process/start/{quote(name, safe='')}")

    def stop_node(self, name):
        """指定Nodeを停止する。"""
        return self.request("PATCH", f"/process/stop/{quote(name, safe='')}")

    def restart_node(self, name):
        """指定Nodeを再起動する。"""
        return self.request("POST", f"/process/restart/{quote(name, safe='')}")
