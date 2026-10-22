
import queue
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from node_window import NodeWindow
from process_service import ProcessService


COLORS = {
    "running": "#22C55E", "busy": "#F59E0B",
    "error": "#EF4444", "stopped": "#94A3B8",
}
ICONS = {"start": "▶", "stop": "■", "restart": "↻", "dashboard": "▣"}


class ProcessManagerApp:
    """Process ComposeとNodeの操作GUI。"""

    def __init__(self, root):
        self.root = root
        self.service = ProcessService()
        self.events = queue.Queue()
        self.windows = {}
        self.buttons = []
        self.busy = False
        self.closing = False
        self.status = tk.StringVar(value="Ready")
        self.server_status = tk.StringVar(value="Stopped")
        self.build_ui()
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(100, self.poll)
        self.submit("refresh")

    @staticmethod
    def status_color(state):
        state = str(state).lower()
        if state == "running":
            return COLORS["running"]
        if state in ("pending", "launching", "launched", "restarting", "terminating"):
            return COLORS["busy"]
        if state in ("error", "failed"):
            return COLORS["error"]
        return COLORS["stopped"]

    def add_button(self, parent, symbol, description, command, column):
        """アイコンボタンとホバー説明を設定する。"""
        button = ttk.Button(parent, text=symbol, width=4, command=command)
        button.grid(row=0, column=column, padx=3)
        self.buttons.append(button)
        tooltip = {"window": None}

        def show(event):
            if tooltip["window"] or not button.winfo_exists():
                return
            popup = tk.Toplevel(button)
            popup.overrideredirect(True)
            popup.attributes("-topmost", True)
            ttk.Label(popup, text=description, padding=(8, 4)).grid()
            popup.geometry(f"+{button.winfo_rootx()}+{button.winfo_rooty() + button.winfo_height() + 4}")
            tooltip["window"] = popup

        def hide(event):
            if tooltip["window"]:
                tooltip["window"].destroy()
                tooltip["window"] = None

        button.bind("<Enter>", show)
        button.bind("<Leave>", hide)
        button.bind("<Destroy>", hide)
        return button

    def build_ui(self):
        root = self.root
        root.title("Process Manager")
        root.geometry("740x450")
        root.minsize(540, 300)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        server = ttk.LabelFrame(root, text="Process Compose", padding=12)
        server.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        server.columnconfigure(2, weight=1)

        self.server_lamp = tk.Label(server, text="●", font=("Segoe UI", 15),
                                    fg=COLORS["stopped"])
        self.server_lamp.grid(row=0, column=0)
        ttk.Label(server, textvariable=self.server_status, width=14).grid(
            row=0, column=1, padx=6
        )

        controls = ttk.Frame(server)
        controls.grid(row=0, column=3)
        for index, action in enumerate(("start", "stop", "restart")):
            self.add_button(
                controls, ICONS[action], f"Process Compose: {action.title()}",
                lambda a=action: self.submit(f"server_{a}"), index
            )

        panel = ttk.LabelFrame(root, text="Nodes", padding=10)
        panel.grid(row=1, column=0, sticky="nsew", padx=12, pady=6)
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(2, weight=1)

        # Node全体の一括操作ボタンを配置する。
        bulk = ttk.Frame(panel)
        bulk.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        bulk.columnconfigure(0, weight=1)

        bulk_controls = ttk.Frame(bulk)
        bulk_controls.grid(row=0, column=1, sticky="e")

        for index, action in enumerate(("start", "stop", "restart")):
            self.add_button(
                bulk_controls, ICONS[action], f"All Nodes: {action.title()}",
                lambda a=action: self.submit(f"{a}_all"), index
            )

        # 既存のNode一覧ヘッダーを維持する。
        header = ttk.Frame(panel)
        header.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        header.columnconfigure(1, weight=1)
        ttk.Label(header, text="STATUS", width=16).grid(row=0, column=0)
        ttk.Label(header, text="NODE").grid(row=0, column=1, sticky="w")
        ttk.Label(header, text="ACTIONS").grid(row=0, column=2, padx=8)
        self.add_button(header, "↻", "Refresh", lambda: self.submit("refresh"), 3)

        container = ttk.Frame(panel)
        container.grid(row=2, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.node_frame = ttk.Frame(canvas)
        self.node_frame.columnconfigure(0, weight=1)
        window_id = canvas.create_window((0, 0), window=self.node_frame, anchor="nw")
        self.node_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))
        ttk.Label(root, textvariable=self.status, padding=8).grid(
            row=2, column=0, sticky="w", padx=10
        )

    def open_dashboard(self, name):
        """同一NodeのDashboardは重複して開かない。"""
        existing = self.windows.get(name)
        if existing and existing.window.winfo_exists():
            existing.window.lift()
            existing.window.focus_force()
            return
        self.windows[name] = NodeWindow(self.root, self.service.api, name)

    def render_snapshot(self, snapshot):
        """サーバー状態とNode一覧を更新する。"""
        server = snapshot["server"]
        self.server_status.set(server)
        self.server_lamp.configure(fg=self.status_color(server))

        for widget in self.node_frame.winfo_children():
            widget.destroy()
        self.buttons = [button for button in self.buttons if button.winfo_exists()]

        for index, (name, state) in enumerate(snapshot["nodes"].items()):
            row = ttk.Frame(self.node_frame, padding=(4, 7))
            row.grid(row=index * 2, column=0, sticky="ew")
            row.columnconfigure(2, weight=1)

            tk.Label(row, text="●", fg=self.status_color(state),
                     font=("Segoe UI", 12)).grid(row=0, column=0)
            ttk.Label(row, text=state, width=15).grid(row=0, column=1, sticky="w")
            ttk.Label(row, text=name).grid(row=0, column=2, sticky="w", padx=10)

            controls = ttk.Frame(row)
            controls.grid(row=0, column=3, sticky="e")
            for col, action in enumerate(("start", "stop", "restart")):
                self.add_button(
                    controls, ICONS[action], f"{name}: {action.title()}",
                    lambda a=action, n=name: self.submit(a, n), col
                )
            self.add_button(
                controls, ICONS["dashboard"], f"{name}: Dashboard",
                lambda n=name: self.open_dashboard(n), 3
            )
            ttk.Separator(self.node_frame, orient="horizontal").grid(
                row=index * 2 + 1, column=0, sticky="ew"
            )

    def set_busy(self, busy):
        self.busy = busy
        for button in self.buttons:
            if button.winfo_exists():
                button.configure(state="disabled" if busy else "normal")

    def submit(self, action, name=None, confirmed=False):
        """ProcessServiceへの操作を別スレッドで実行する。"""
        if self.busy or self.closing:
            return

        # Server停止・再起動に加え、Nodeの一括停止・再起動も確認する。
        if action in ("server_stop", "server_restart", "stop_all", "restart_all") and not confirmed:
            message = (
                "Process Composeと全Nodeを停止しますか？"
                if action.startswith("server_")
                else "全Nodeを停止しますか？"
                if action == "stop_all"
                else "全Nodeを停止してから再起動しますか？"
            )
            if not messagebox.askyesno("Confirm", message):
                return

        self.set_busy(True)
        self.status.set(f"Executing: {action}" + (f" ({name})" if name else ""))
        threading.Thread(target=self.execute, args=(action, name), daemon=True).start()

    def execute(self, action, name):
        try:
            self.events.put((self.service.execute(action, name), None))
        except Exception as exc:
            self.events.put((None, str(exc)))

    def poll(self):
        """結果をGUIスレッドで受け取る。"""
        while not self.events.empty():
            result, failure = self.events.get_nowait()
            self.set_busy(False)
            if failure:
                self.status.set("Error")
                if self.closing:
                    self.closing = False
                    self.root.deiconify()
                messagebox.showerror("Process Manager", failure)
                continue
            if self.closing:
                self.root.destroy()
                return
            self.render_snapshot(result)
            self.status.set("Ready")
        self.root.after(100, self.poll)

    def close(self):
        """キャッシュ保存とTUI停止を既存Serviceに委譲する。"""
        if self.busy or self.closing:
            return
        server = self.service.server
        if server is None or server.poll() is not None:
            self.root.destroy()
            return
        if not messagebox.askyesno("Exit", "Process Composeを停止して終了しますか？"):
            return
        self.root.withdraw()
        self.submit("server_stop", confirmed=True)
        self.closing = True


if __name__ == "__main__":
    root = tk.Tk()
    ProcessManagerApp(root)
    root.mainloop()
