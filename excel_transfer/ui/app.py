# excel_transfer/ui/app.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from models.dto import GrepRequest, DiffRequest, TransferRequest, CountRequest
from services.transfer import run_transfer_from_csvs
from services.grep import run_grep
from services.diff import run_diff
from services.count import run_count


# ==================================================
# Base Tab
# ==================================================
class BaseTab:
    def __init__(self, app, title: str):
        self.app = app
        self.ctx = app.ctx
        self.logger = app.logger
        self.tab = ttk.Frame(app.nb, padding=(4, 2, 4, 2))
        app.nb.add(self.tab, text=title)

    def log(self, msg: str):
        self.app.append_log(msg)


# ==================================================
# Transfer Tab
# ==================================================
class TransferTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_transfer"])
        self.build()

    def build(self):
        ttk.Label(self.tab, text=self.ctx.labels["label_transfer_config"]).grid(row=0, column=0, sticky="w")

        self.entry = tk.Entry(self.tab, width=70)
        self.entry.grid(row=0, column=1, sticky="we", padx=4)
        self.entry.insert(0, self.ctx.user_paths.get("transfer_config", ""))

        ttk.Button(
            self.tab, text="...", width=3,
            command=lambda: self.app.choose_files(
                self.entry, "transfer_config", [("CSV Files", "*.csv")]
            )
        ).grid(row=0, column=2)

        self.var_skip = tk.BooleanVar(value=bool(self.ctx.user_paths.get("transfer_skip_oor", False)))
        ttk.Checkbutton(
            self.tab, text="範囲外セルはスキップ（警告のみ）", variable=self.var_skip
        ).grid(row=1, column=1, sticky="w")

        ttk.Button(self.tab, text="実行", command=self.run).grid(row=0, column=3, rowspan=2, padx=4)
        self.tab.grid_columnconfigure(1, weight=1)

    def run(self):
        self.ctx.save_user_path("transfer_config", self.entry.get())
        self.ctx.save_user_path("transfer_skip_oor", bool(self.var_skip.get()))

        paths = [p for p in self.entry.get().split("?") if p]
        mode = "skip" if self.var_skip.get() else "error"
        req = TransferRequest(csv_paths=paths, out_of_range_mode=mode)

        try:
            note = run_transfer_from_csvs(req, self.ctx, self.logger, self.log)
            self.log(f"[OK] 転記完了: {note}")
        except Exception as e:
            self.log(f"[ERR] 転記: {e}")


# ==================================================
# Grep Tab
# ==================================================
class GrepTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_grep"])
        self.build()

    def build(self):
        ttk.Label(self.tab, text=self.ctx.labels["label_grep_root"]).grid(row=0, column=0, sticky="w")
        self.root_entry = tk.Entry(self.tab, width=70)
        self.root_entry.grid(row=0, column=1, sticky="we", padx=4)
        self.root_entry.insert(0, self.ctx.user_paths.get("grep_root", ""))

        ttk.Button(
            self.tab, text="...", width=3,
            command=lambda: self.app.choose_dir(self.root_entry, "grep_root")
        ).grid(row=0, column=2)

        ttk.Label(self.tab, text=self.ctx.labels["label_grep_keyword"]).grid(row=1, column=0, sticky="w")
        self.kw_entry = tk.Entry(self.tab, width=40)
        self.kw_entry.grid(row=1, column=1, sticky="we", padx=4)
        self.kw_entry.insert(0, self.ctx.user_paths.get("grep_keyword", ""))

        self.var_ic = tk.BooleanVar(value=bool(self.ctx.user_paths.get("grep_ignore_case", True)))
        self.var_rx = tk.BooleanVar(value=bool(self.ctx.user_paths.get("grep_use_regex", False)))

        ttk.Checkbutton(self.tab, text=self.ctx.labels["check_ignore_case"], variable=self.var_ic).grid(row=1, column=2)
        ttk.Checkbutton(self.tab, text="正規表現", variable=self.var_rx).grid(row=1, column=3)

        ttk.Button(self.tab, text="実行", command=self.run).grid(row=0, column=4, rowspan=2, padx=4)
        self.tab.grid_columnconfigure(1, weight=1)

    def run(self):
        self.ctx.save_user_path("grep_root", self.root_entry.get())
        self.ctx.save_user_path("grep_keyword", self.kw_entry.get())
        self.ctx.save_user_path("grep_ignore_case", bool(self.var_ic.get()))
        self.ctx.save_user_path("grep_use_regex", bool(self.var_rx.get()))

        req = GrepRequest(
            root_dir=self.root_entry.get().strip(),
            keyword=self.kw_entry.get().strip(),
            ignore_case=self.var_ic.get(),
            use_regex=self.var_rx.get(),
        )

        try:
            out, cnt = run_grep(req, self.ctx, self.logger, self.log)
            self.log(f"[OK] Grep結果: {out} / {cnt}件")
        except Exception as e:
            self.log(f"[ERR] Grep: {e}")


# ==================================================
# Diff Tab（範囲指定対応）
# ==================================================
class DiffTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_diff"])
        self.build()

    def build(self):
        # --- File A ---
        ttk.Label(self.tab, text=self.ctx.labels["label_diff_file_a"])\
            .grid(row=0, column=0, sticky="w")

        self.file_a = tk.Entry(self.tab, width=70)
        self.file_a.grid(row=0, column=1, sticky="we", padx=4)
        self.file_a.insert(0, self.ctx.user_paths.get("diff_file_a", ""))

        ttk.Button(
            self.tab, text="...", width=3,
            command=lambda: self.app.choose_file(self.file_a, "diff_file_a")
        ).grid(row=0, column=2)

        # --- File B ---
        ttk.Label(self.tab, text=self.ctx.labels["label_diff_file_b"])\
            .grid(row=1, column=0, sticky="w")

        self.file_b = tk.Entry(self.tab, width=70)
        self.file_b.grid(row=1, column=1, sticky="we", padx=4)
        self.file_b.insert(0, self.ctx.user_paths.get("diff_file_b", ""))

        ttk.Button(
            self.tab, text="...", width=3,
            command=lambda: self.app.choose_file(self.file_b, "diff_file_b")
        ).grid(row=1, column=2)

        # --- Range A ---
        ttk.Label(self.tab, text="比較範囲 A（空=全体）")\
            .grid(row=2, column=0, sticky="w")

        self.range_a = tk.Entry(self.tab, width=20)
        self.range_a.grid(row=2, column=1, sticky="w", padx=4)
        self.range_a.insert(0, self.ctx.user_paths.get("diff_range_a", ""))

        # --- Range B ---
        ttk.Label(self.tab, text="比較範囲 B（空=全体）")\
            .grid(row=3, column=0, sticky="w")

        self.range_b = tk.Entry(self.tab, width=20)
        self.range_b.grid(row=3, column=1, sticky="w", padx=4)
        self.range_b.insert(0, self.ctx.user_paths.get("diff_range_b", ""))

        # --- Options ---
        self.var_formula = tk.BooleanVar(
            value=bool(self.ctx.user_paths.get("diff_compare_formula", False))
        )
        self.var_ctx = tk.BooleanVar(
            value=bool(self.ctx.user_paths.get("diff_include_context", True))
        )
        self.var_shapes = tk.BooleanVar(
            value=bool(self.ctx.user_paths.get("diff_compare_shapes", False))
        )

        ttk.Checkbutton(
            self.tab, text="数式比較", variable=self.var_formula
        ).grid(row=2, column=2, sticky="w")

        ttk.Checkbutton(
            self.tab, text="ジャンプリンク/コンテキスト", variable=self.var_ctx
        ).grid(row=3, column=2, sticky="w")

        ttk.Checkbutton(
            self.tab, text="図・画像も比較", variable=self.var_shapes
        ).grid(row=4, column=2, sticky="w")

        # --- Run ---
        ttk.Button(
            self.tab, text="実行", command=self.run
        ).grid(row=0, column=3, rowspan=5, padx=6)

        self.tab.grid_columnconfigure(1, weight=1)

    def run(self):
        # 保存
        self.ctx.save_user_path("diff_file_a", self.file_a.get())
        self.ctx.save_user_path("diff_file_b", self.file_b.get())
        self.ctx.save_user_path("diff_range_a", self.range_a.get())
        self.ctx.save_user_path("diff_range_b", self.range_b.get())
        self.ctx.save_user_path("diff_compare_formula", bool(self.var_formula.get()))
        self.ctx.save_user_path("diff_include_context", bool(self.var_ctx.get()))
        self.ctx.save_user_path("diff_compare_shapes", bool(self.var_shapes.get()))

        # Request
        req = DiffRequest(
            file_a=self.file_a.get().strip(),
            file_b=self.file_b.get().strip(),
            key_cols=[],
            compare_formula=self.var_formula.get(),
            include_context=self.var_ctx.get(),
            compare_shapes=self.var_shapes.get(),
        )

        # 動的に範囲を付与（後方互換）
        req.range_a = self.range_a.get().strip()
        req.range_b = self.range_b.get().strip()

        try:
            out = run_diff(req, self.ctx, self.logger, self.log)
            self.log(f"[OK] 差分レポート: {out}")
        except Exception as e:
            self.log(f"[ERR] Diff: {e}")

# ==================================================
# Count Tab
# ==================================================
class CountTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, "Count")
        self.build()

    def build(self):
        ttk.Label(self.tab, text="対象Excel").grid(row=0, column=0, sticky="w")
        self.files = tk.Entry(self.tab, width=70)
        self.files.grid(row=0, column=1, sticky="we", padx=4)
        self.files.insert(0, self.ctx.user_paths.get("count_files", ""))

        ttk.Button(
            self.tab, text="...", width=3,
            command=lambda: self.app.choose_files(
                self.files, "count_files",
                [("Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls")]
            )
        ).grid(row=0, column=2)

        ttk.Label(self.tab, text="シート名（空=先頭）").grid(row=1, column=0, sticky="w")
        self.sheet = tk.Entry(self.tab, width=30)
        self.sheet.grid(row=1, column=1, sticky="w", padx=4)
        self.sheet.insert(0, self.ctx.user_paths.get("count_sheet", ""))

        ttk.Label(self.tab, text="開始セル").grid(row=2, column=0, sticky="w")
        self.start = tk.Entry(self.tab, width=12)
        self.start.grid(row=2, column=1, sticky="w", padx=4)
        self.start.insert(0, self.ctx.user_paths.get("count_start", "B2"))

        self.dir = tk.StringVar(value=self.ctx.user_paths.get("count_dir", "row"))
        ttk.Radiobutton(self.tab, text="行方向", variable=self.dir, value="row").grid(row=3, column=1, sticky="w")
        ttk.Radiobutton(self.tab, text="列方向", variable=self.dir, value="col").grid(row=3, column=1, padx=80)

        ttk.Label(self.tab, text="許容空白数").grid(row=4, column=0, sticky="w")
        self.tol = tk.Spinbox(self.tab, from_=0, to=1000, width=6)
        self.tol.grid(row=4, column=1, sticky="w", padx=4)
        self.tol.delete(0, "end")
        self.tol.insert(0, str(self.ctx.user_paths.get("count_tolerate_blanks", 0)))

        self.mode = tk.StringVar(value=self.ctx.user_paths.get("count_mode", "jump"))
        ttk.Radiobutton(self.tab, text="高速", variable=self.mode, value="jump").grid(row=5, column=1, sticky="w")
        ttk.Radiobutton(self.tab, text="精密", variable=self.mode, value="scan").grid(row=5, column=1, padx=80)

        ttk.Button(self.tab, text="実行", command=self.run).grid(row=7, column=3, sticky="se", padx=6, pady=6)

        self.tab.grid_columnconfigure(1, weight=1)

    def run(self):
        self.ctx.save_user_path("count_files", self.files.get())
        self.ctx.save_user_path("count_sheet", self.sheet.get())
        self.ctx.save_user_path("count_start", self.start.get())
        self.ctx.save_user_path("count_dir", self.dir.get())
        self.ctx.save_user_path("count_tolerate_blanks", int(self.tol.get() or 0))
        self.ctx.save_user_path("count_mode", self.mode.get())

        req = CountRequest(
            files=[p for p in self.files.get().split("?") if p],
            sheet=self.sheet.get().strip(),
            start_cell=self.start.get().strip(),
            direction=self.dir.get(),
            tolerate_blanks=int(self.tol.get() or 0),
            mode=self.mode.get(),
        )

        try:
            out = run_count(req, self.ctx, self.logger, self.log)
            self.log(f"[OK] Count結果: {out}")
        except Exception as e:
            self.log(f"[ERR] Count: {e}")


# ==================================================
# App
# ==================================================
class ExcelApp:
    def __init__(self, ctx, logger):
        self.ctx, self.logger = ctx, logger

        self.root = tk.Tk()
        self.root.title(ctx.labels["app_title"])

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=4, pady=2)

        self.transfer_tab = TransferTab(self)
        self.grep_tab = GrepTab(self)
        self.diff_tab = DiffTab(self)
        self.count_tab = CountTab(self)

        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=False, padx=4, pady=(2, 4))
        ttk.Label(log_frame, text="ログ").pack(anchor="w")
        self.log = tk.Text(log_frame, height=8)
        self.log.pack(fill="both", expand=True)
        ttk.Button(log_frame, text="コピー", command=self.copy_log).pack(anchor="e")

        self.root.update_idletasks()
        self.root.minsize(self.root.winfo_reqwidth(), self.root.winfo_reqheight())

    # --------------------------
    # utilities
    # --------------------------
    def append_log(self, msg: str):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)

    def copy_log(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.log.get("1.0", tk.END))
        messagebox.showinfo("クリップボード", "ログをコピーしました。")

    def choose_files(self, entry, key, filetypes):
        paths = filedialog.askopenfilenames(
            initialdir=self.ctx.default_dir_for(entry.get()),
            filetypes=filetypes
        )
        if paths:
            entry.delete(0, tk.END)
            entry.insert(0, "?".join(paths))
            self.ctx.save_user_path(key, entry.get())

    def choose_file(self, entry, key):
        path = filedialog.askopenfilename(
            initialdir=self.ctx.default_dir_for(entry.get()),
            filetypes=[("Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls")]
        )
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            self.ctx.save_user_path(key, path)

    def choose_dir(self, entry, key):
        path = filedialog.askdirectory(
            initialdir=self.ctx.default_dir_for(entry.get())
        )
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            self.ctx.save_user_path(key, path)

    def run(self):
        self.root.mainloop()
