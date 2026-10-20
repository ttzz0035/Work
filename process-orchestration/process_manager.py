
import queue
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from process_service import ProcessService


class ProcessManagerApp:
    """Process Composeの管理GUI。UI操作と状態表示を担当する。"""

    def __init__(self, root):
        self.root = root
        self.service = ProcessService()
        self.events = queue.Queue()
        self.busy = False
        self.closing = False

        # Tkinterの表示状態はメインスレッドでのみ操作する。
        self.status = tk.StringVar(value="Ready")
        self.server_status = tk.StringVar(value="Stopped")
        self.buttons = []

        self.build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after(100, self.poll)
        self.submit("refresh")

    def build_ui(self):
        """画面を構築する。API通信や業務処理は行わない。"""
        self.root.title("Process Manager")
        self.root.geometry("640x400")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        # Server操作エリア。
        server = ttk.LabelFrame(self.root, text="Server", padding=10)
        server.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(server, textvariable=self.server_status, width=15).grid(row=0, column=0)

        for index, action in enumerate(("start", "stop", "restart"), 1):
            button = ttk.Button(server, text=action.title(), command=lambda a=action: self.submit("server_" + a))
            button.grid(row=0, column=index, padx=4)
            self.buttons.append(button)

        # Node一覧。
        ttk.Label(self.root, text="Nodes", padding=5).grid(row=1, column=0, sticky="w")

        self.tree = ttk.Treeview(self.root, columns=("node", "status"), show="headings", selectmode="browse")

        for column in ("node", "status"):
            self.tree.heading(column, text=column.title())
            self.tree.column(column, width=250)

        self.tree.grid(row=2, column=0, sticky="nsew", padx=10)

        # Node操作エリア。
        bar = ttk.Frame(self.root)
        bar.grid(row=3, column=0, pady=10)

        for index, action in enumerate(("start", "stop", "restart", "refresh")):
            button = ttk.Button(bar, text=action.title(), command=lambda a=action: self.submit(a))
            button.grid(row=0, column=index, padx=5)
            self.buttons.append(button)

        # アプリケーション状態。
        ttk.Label(self.root, textvariable=self.status, padding=8).grid(row=4, column=0, sticky="w")

    def set_busy(self, busy):
        """処理中の二重操作を防止する。"""
        self.busy = busy

        for button in self.buttons:
            button.configure(state="disabled" if busy else "normal")

    def submit(self, action, confirmed=False):
        """操作を受け付け、バックグラウンド処理へ渡す。"""
        if self.busy or self.closing:
            return

        selection = self.tree.selection()
        node_actions = ("start", "stop", "restart")

        # Node操作には対象Nodeの選択が必要。
        if action in node_actions and not selection:
            return

        # Serverの停止・再起動は確認する。
        if action in ("server_stop", "server_restart") and not confirmed:
            if not messagebox.askyesno("Confirm", "Process Composeと全Nodeを停止しますか？"):
                return

        name = selection[0] if action in node_actions else None

        self.set_busy(True)
        self.status.set(f"Executing: {action}")

        threading.Thread(target=self.execute, args=(action, name), daemon=True).start()

    def execute(self, action, name):
        """Serviceを別スレッドで実行し、結果をQueueへ返す。"""
        try:
            result = self.service.execute(action, name)
            self.events.put((result, None))
        except Exception as exc:
            self.events.put((None, str(exc)))

    def render_snapshot(self, snapshot):
        """Serviceから取得した状態をNode一覧に反映する。"""
        self.server_status.set(snapshot["server"])
        nodes = snapshot["nodes"]

        # 更新前のNode選択を保持する。
        selection = self.tree.selection()

        # 設定から削除されたNodeを画面から除去する。
        for name in self.tree.get_children():
            if name not in nodes:
                self.tree.delete(name)

        # 既存Nodeを更新し、新規Nodeのみ追加する。
        for name, state in nodes.items():
            if self.tree.exists(name):
                self.tree.item(name, values=(name, state))
            else:
                self.tree.insert("", "end", iid=name, values=(name, state))

        # 有効な選択だけを復元する。
        for name in selection:
            if self.tree.exists(name):
                self.tree.selection_add(name)

    def poll(self):
        """Queueを監視し、メインスレッドで画面を更新する。"""
        while not self.events.empty():
            result, failure = self.events.get_nowait()
            self.set_busy(False)

            if failure:
                self.status.set("Error")

                # 終了処理に失敗した場合はWindowを再表示する。
                if self.closing:
                    self.closing = False
                    self.root.deiconify()

                messagebox.showerror("Process Manager", failure)
                continue

            # TUI終了が完了した場合にGUIを破棄する。
            if self.closing:
                self.root.destroy()
                return

            self.render_snapshot(result)
            self.status.set("Ready")

        self.root.after(100, self.poll)

    def close(self):
        """終了前にComposeキャッシュ保存とTUI停止を実行する。"""
        if self.busy or self.closing:
            return

        server = self.service.server

        # 管理対象Serverが停止していれば、そのまま終了する。
        if server is None or server.poll() is not None:
            self.root.destroy()
            return

        if not messagebox.askyesno("Exit", "Process Composeを停止して終了しますか？"):
            return

        # Windowを即時非表示にし、終了処理は別スレッドで実行する。
        self.root.withdraw()
        self.submit("server_stop", confirmed=True)
        self.closing = True


if __name__ == "__main__":
    root = tk.Tk()
    ProcessManagerApp(root)
    root.mainloop()
