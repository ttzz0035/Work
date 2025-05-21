import tkinter as tk
from tkinter import ttk, filedialog

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("メイン画面")
        self.root.geometry("600x400")
        self.prev_geom = ""  # 前回のサイズ・位置記録用
        self.root.bind("<Configure>", self.on_resize_or_move)  # リサイズや移動イベント監視

        self.init_menu()
        self.create_main_layout()

        # 初回ログ出力
        self.log_geometry("起動時")

    def log_geometry(self, prefix=""):
        geom = self.root.winfo_geometry()
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        print(f"[{prefix}] geometry={geom} pos=({x},{y}) size=({width},{height})")

    def on_resize_or_move(self, event):
        geom = self.root.winfo_geometry()
        if geom != self.prev_geom:
            self.prev_geom = geom
            self.log_geometry("変化検出")

    def init_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="終了", command=self.root.quit)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        self.root.config(menu=menubar)

    def create_main_layout(self):
        self.root.grid_rowconfigure(0, weight=2)
        self.root.grid_rowconfigure(1, weight=6)
        self.root.grid_rowconfigure(2, weight=2)
        self.root.grid_columnconfigure(0, weight=1)

        frame1 = tk.Frame(self.root, bg="skyblue")
        frame1.grid(row=0, column=0, sticky="nsew")

        frame2 = tk.Frame(self.root, bg="lightgreen")
        frame2.grid(row=1, column=0, sticky="nsew")
        frame2.grid_columnconfigure(0, weight=1)
        frame2.grid_columnconfigure(1, weight=1)
        frame2.grid_rowconfigure(0, weight=1)

        self.lf1_widgets = self.create_file_widget_group(frame2, "LabelFrame1", 0)
        self.lf2_widgets = self.create_file_widget_group(frame2, "LabelFrame2", 1)

        frame3 = tk.Frame(self.root, bg="orange")
        frame3.grid(row=2, column=0, sticky="nsew")
        run_btn = ttk.Button(frame3, text="実行", command=self.open_modal_window)
        run_btn.pack(pady=10)

    def create_file_widget_group(self, parent, label, col):
        lf = ttk.LabelFrame(parent, text=label)
        lf.grid(row=0, column=col, sticky="nsew", padx=5, pady=5)

        path_var = tk.StringVar()
        entry = ttk.Entry(lf, textvariable=path_var)
        entry.pack(fill="x", padx=5, pady=(5, 0))

        listbox = tk.Listbox(lf)
        listbox.pack(fill="both", expand=True, padx=5, pady=5)

        def select_file():
            filepath = filedialog.askopenfilename()
            if filepath:
                path_var.set(filepath)
                listbox.insert(tk.END, filepath)

        button = ttk.Button(lf, text="ファイル選択", command=select_file)
        button.pack(padx=5, pady=5)

        return {"entry": entry, "button": button, "listbox": listbox, "path_var": path_var}

    def open_modal_window(self):
        modal = tk.Toplevel(self.root)
        modal.title("モーダル画面")
        modal.resizable(False, False)

        # 中央配置
        w, h = 300, 150
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 2) - (h // 2)
        modal.geometry(f"{w}x{h}+{x}+{y}")

        modal.grab_set()
        modal.focus_force()
        modal.protocol("WM_DELETE_WINDOW", lambda: self.cancel_modal(modal))

        label = tk.Label(modal, text="10秒後に比較結果画面を開きます", font=("Arial", 11))
        label.pack(pady=20)

        cancel_btn = ttk.Button(modal, text="キャンセル", command=lambda: self.cancel_modal(modal))
        cancel_btn.pack(pady=5)

        modal.after(10000, lambda: self.auto_close_modal(modal))

    def cancel_modal(self, modal):
        modal.grab_release()
        modal.destroy()

    def auto_close_modal(self, modal):
        if not modal.winfo_exists():
            return
        modal.grab_release()
        modal.destroy()
        self.show_result_screen()

    def show_result_screen(self):
        result_win = tk.Toplevel(self.root)
        result_win.title("比較結果画面")
        result_win.geometry("700x400")

        result_frame = tk.Frame(result_win)
        result_frame.pack(fill="both", expand=True, padx=10, pady=10)
        result_frame.grid_columnconfigure(0, weight=1)
        result_frame.grid_columnconfigure(1, weight=1)
        result_frame.grid_rowconfigure(0, weight=1)

        lf1 = ttk.LabelFrame(result_frame, text="LabelFrame1 結果")
        lf1.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        lb1 = tk.Listbox(lf1)
        lb1.pack(fill="both", expand=True, padx=5, pady=5)

        lf2 = ttk.LabelFrame(result_frame, text="LabelFrame2 結果")
        lf2.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        lb2 = tk.Listbox(lf2)
        lb2.pack(fill="both", expand=True, padx=5, pady=5)

        # ✅ 比較結果（色付き）
        lb1.insert(tk.END, "OK: 同じ行 A")
        lb1.itemconfig(tk.END, {'fg': 'black'})

        lb1.insert(tk.END, "NG: 差分あり A")
        lb1.itemconfig(tk.END, {'fg': 'red'})

        lb2.insert(tk.END, "OK: 同じ行 B")
        lb2.itemconfig(tk.END, {'fg': 'black'})

        lb2.insert(tk.END, "NG: 差分あり B")
        lb2.itemconfig(tk.END, {'fg': 'red'})

        close_btn = ttk.Button(result_win, text="閉じる", command=result_win.destroy)
        close_btn.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
