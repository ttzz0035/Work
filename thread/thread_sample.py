import tkinter as tk
from tkinter import ttk


class BusyWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        self.geometry("320x140")
        self.title("処理中")
        self.progress = ttk.Progressbar(self, mode="determinate", maximum=100, length=260)
        self.progress.pack(pady=(20, 20))
        self.label = tk.Label(self, text="待機中")
        self.label.pack(pady=(0, 10))

    def show(self, text="処理中...", value=0):
        self.label.config(text=text)
        self.progress.configure(value=value)
        self.deiconify()
        self.lift()

    def hide(self):
        self.withdraw()

    def set_text(self, text):
        self.label.config(text=text)

    def set_progress(self, value):
        self.progress.configure(value=value)


if __name__ == '__main__':
    import threading
    import time
    def on_done():
        busy.set_text("完了")
        if auto_close_var.get():
            busy.hide()

    def worker():
        for i in range(101):
            time.sleep(0.03)
            root.after(0, lambda n=i: busy.set_progress(n))
            root.after(0, lambda n=i: busy.set_text(f"処理中 {n}%"))
        root.after(0, on_done)

    def on_click():
        busy.show("開始", 0)
        threading.Thread(target=worker, daemon=True).start()

    root = tk.Tk()
    root.geometry("320x200")

    busy = BusyWindow(root)
    auto_close_var = tk.BooleanVar(value=False)

    tk.Checkbutton(
        root,
        text="完了時に自動で閉じる",
        variable=auto_close_var,
    ).pack(pady=(20, 10))

    tk.Button(root, text="開始", command=on_click).pack(pady=20)

    root.mainloop()