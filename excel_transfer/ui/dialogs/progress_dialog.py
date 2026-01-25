# ============================================
# ui/dialogs/progress_dialog.py
# ============================================
import tkinter as tk
from tkinter import ttk


class ProgressDialog:
    def __init__(self, parent: tk.Tk, ctx, message: str | None = None):
        self._closed = False
        self.parent = parent
        self.ctx = ctx

        if message is None:
            message = self.ctx.labels["progress_message_default"]

        self.win = tk.Toplevel(parent)
        self.win.title(self.ctx.labels["progress_title"])
        self.win.resizable(False, False)
        self.win.transient(parent)
        self.win.grab_set()

        self.win.protocol("WM_DELETE_WINDOW", self._disable_close)

        # ----------------------------------------
        # UI
        # ----------------------------------------
        frm = ttk.Frame(self.win, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text=message).pack(anchor="w")

        self.pb = ttk.Progressbar(
            frm,
            mode="indeterminate",
            length=240,
        )
        self.pb.pack(pady=(12, 0))
        self.pb.start(10)

        # ----------------------------------------
        # 中央配置
        # ----------------------------------------
        self._center_to_parent()

    # ----------------------------------------
    # Utils
    # ----------------------------------------
    def _center_to_parent(self):
        # レイアウト確定
        self.win.update_idletasks()
        self.parent.update_idletasks()

        # ダイアログサイズ
        w = self.win.winfo_width()
        h = self.win.winfo_height()

        # 親Window位置・サイズ
        px = self.parent.winfo_rootx()
        py = self.parent.winfo_rooty()
        pw = self.parent.winfo_width()
        ph = self.parent.winfo_height()

        # 中央座標
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2

        self.win.geometry(f"{w}x{h}+{x}+{y}")

    def _disable_close(self):
        pass  # 閉じさせない

    def close(self):
        if self._closed:
            return
        self._closed = True
        try:
            self.pb.stop()
            self.win.grab_release()
            self.win.destroy()
        except Exception:
            pass
