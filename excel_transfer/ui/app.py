# ============================================
# excel_transfer/ui/app.py
# ============================================
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from models.dto import GrepRequest, DiffRequest, TransferRequest, CountRequest
from services.transfer import run_transfer_from_csvs
from services.grep import run_grep
from services.diff import run_diff
from services.count import run_count
from outputs.excel_diff_html import generate_html_report
from outputs.transfer_csv_builder import TransferCsvBuilderDialog


class BaseTab:
    def __init__(self, app, title: str):
        self.app = app
        self.ctx = app.ctx
        self.logger = app.logger
        self.tab = ttk.Frame(app.nb, padding=(6, 6, 6, 6))
        app.nb.add(self.tab, text=title)

    def log(self, msg: str):
        self.app.append_log(msg)


# =========================================================
# Transfer Tab
# =========================================================
class TransferTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_transfer"])
        self.build()

    def build(self):
        frm = ttk.LabelFrame(self.tab, text="転記")
        frm.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text=self.ctx.labels["label_transfer_config"]).grid(row=0, column=0, sticky="w")
        self.entry = tk.Entry(frm, width=70)
        self.entry.grid(row=0, column=1, sticky="we", padx=4)
        self.entry.insert(0, self.ctx.user_paths.get("transfer_config", ""))

        ttk.Button(
            frm,
            text="...",
            width=3,
            command=lambda: self.app.choose_files(self.entry, "transfer_config", [("CSV Files", "*.csv")])
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Button(
            frm,
            text="作成",
            width=5,
            command=self.open_builder,
        ).grid(row=0, column=3, padx=(0, 4))

        self.var_skip = tk.BooleanVar(value=bool(self.ctx.user_paths.get("transfer_skip_oor", False)))
        ttk.Checkbutton(
            frm,
            text="範囲外セルはスキップ（警告のみ）",
            variable=self.var_skip,
        ).grid(row=1, column=1, sticky="w", padx=4, pady=(2, 0))

        ttk.Button(self.tab, text="実行", command=self.run).grid(row=1, column=1, sticky="ne", padx=6, pady=6)

        self.tab.grid_columnconfigure(0, weight=1)

    def open_builder(self):
        self.logger.info("[UI] open transfer csv builder")

        def on_created(path: str):
            self.entry.delete(0, tk.END)
            self.entry.insert(0, path)
            self.ctx.save_user_path("transfer_config", path)

        TransferCsvBuilderDialog(
            parent=self.tab,
            ctx=self.ctx,
            logger=self.logger,
            on_created=on_created,
        )

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


# =========================================================
# Grep Tab
# =========================================================
class GrepTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_grep"])
        self.build()

    def build(self):
        frm = ttk.LabelFrame(self.tab, text="Grep")
        frm.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text=self.ctx.labels["label_grep_root"]).grid(row=0, column=0, sticky="w")
        self.root_entry = tk.Entry(frm, width=70)
        self.root_entry.grid(row=0, column=1, sticky="we", padx=4)
        self.root_entry.insert(0, self.ctx.user_paths.get("grep_root", ""))

        ttk.Button(
            frm, text="...", width=3,
            command=lambda: self.app.choose_dir(self.root_entry, "grep_root")
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Label(frm, text=self.ctx.labels["label_grep_keyword"]).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.kw_entry = tk.Entry(frm, width=40)
        self.kw_entry.grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))
        self.kw_entry.insert(0, self.ctx.user_paths.get("grep_keyword", ""))

        self.var_ic = tk.BooleanVar(value=bool(self.ctx.user_paths.get("grep_ignore_case", True)))
        self.var_rx = tk.BooleanVar(value=bool(self.ctx.user_paths.get("grep_use_regex", False)))

        opt = ttk.Frame(frm)
        opt.grid(row=2, column=1, sticky="w", padx=4, pady=(4, 0))
        ttk.Checkbutton(opt, text=self.ctx.labels["check_ignore_case"], variable=self.var_ic).pack(side="left")
        ttk.Checkbutton(opt, text="正規表現", variable=self.var_rx).pack(side="left", padx=(8, 0))

        btn = ttk.Button(self.tab, text="実行", command=self.run)
        btn.grid(row=1, column=1, sticky="ne", padx=6, pady=6)

        self.tab.grid_columnconfigure(0, weight=1)

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


# =========================================================
# Diff Tab
# =========================================================
class DiffTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_diff"])
        self.build()

    def build(self):
        self.tab.grid_columnconfigure(0, weight=1)

        files = ttk.LabelFrame(self.tab, text="ファイル")
        files.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        files.grid_columnconfigure(1, weight=1)

        ttk.Label(files, text=self.ctx.labels["label_diff_file_a"]).grid(row=0, column=0, sticky="w")
        self.file_a = tk.Entry(files, width=70)
        self.file_a.grid(row=0, column=1, sticky="we", padx=4)
        self.file_a.insert(0, self.ctx.user_paths.get("diff_file_a", ""))

        ttk.Button(
            files, text="...", width=3,
            command=lambda: self.app.choose_file(self.file_a, "diff_file_a")
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Label(files, text=self.ctx.labels["label_diff_file_b"]).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.file_b = tk.Entry(files, width=70)
        self.file_b.grid(row=1, column=1, sticky="we", padx=4, pady=(4, 0))
        self.file_b.insert(0, self.ctx.user_paths.get("diff_file_b", ""))

        ttk.Button(
            files, text="...", width=3,
            command=lambda: self.app.choose_file(self.file_b, "diff_file_b")
        ).grid(row=1, column=2, padx=(0, 4), pady=(4, 0))

        # --- ranges + base ---
        mid = ttk.Frame(self.tab)
        mid.grid(row=1, column=0, sticky="we", padx=4, pady=4)
        mid.grid_columnconfigure(0, weight=1)
        mid.grid_columnconfigure(1, weight=1)

        rng = ttk.LabelFrame(mid, text="比較範囲（必須）")
        rng.grid(row=0, column=0, sticky="we", padx=(0, 4))
        rng.grid_columnconfigure(1, weight=1)

        ttk.Label(rng, text="範囲 A").grid(row=0, column=0, sticky="w")
        self.range_a = tk.Entry(rng, width=20)
        self.range_a.grid(row=0, column=1, sticky="w", padx=4)
        self.range_a.insert(0, self.ctx.user_paths.get("diff_range_a", ""))

        ttk.Label(rng, text="範囲 B").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.range_b = tk.Entry(rng, width=20)
        self.range_b.grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))
        self.range_b.insert(0, self.ctx.user_paths.get("diff_range_b", ""))

        base = ttk.LabelFrame(mid, text="差分ベース（どちらをDIFFにするか）")
        base.grid(row=0, column=1, sticky="we", padx=(4, 0))
        self.diff_base = tk.StringVar(value=self.ctx.user_paths.get("diff_base_file", "B"))
        ttk.Radiobutton(
            base, text="比較先（B）をベース（既定）",
            variable=self.diff_base, value="B"
        ).pack(anchor="w", padx=6, pady=(2, 0))
        ttk.Radiobutton(
            base, text="比較元（A）をベース",
            variable=self.diff_base, value="A"
        ).pack(anchor="w", padx=6, pady=(2, 6))

        # --- options ---
        opt = ttk.LabelFrame(self.tab, text="オプション")
        opt.grid(row=2, column=0, sticky="we", padx=4, pady=4)

        self.var_formula = tk.BooleanVar(value=bool(self.ctx.user_paths.get("diff_compare_formula", False)))
        self.var_ctx = tk.BooleanVar(value=bool(self.ctx.user_paths.get("diff_include_context", True)))
        self.var_shapes = tk.BooleanVar(value=bool(self.ctx.user_paths.get("diff_compare_shapes", False)))

        ttk.Checkbutton(opt, text="数式比較", variable=self.var_formula).pack(side="left", padx=6, pady=6)
        ttk.Checkbutton(opt, text="ジャンプリンク/コンテキスト", variable=self.var_ctx).pack(side="left", padx=6, pady=6)
        ttk.Checkbutton(opt, text="図形/画像も比較", variable=self.var_shapes).pack(side="left", padx=6, pady=6)

        # --- buttons ---
        btns = ttk.Frame(self.tab)
        btns.grid(row=3, column=0, sticky="we", padx=4, pady=(4, 0))
        btns.grid_columnconfigure(0, weight=1)

        self.report_btn = ttk.Button(btns, text="HTMLレポート", command=self.make_report)
        self.report_btn.grid(row=0, column=0, sticky="w", padx=6, pady=6)

        run_btn = ttk.Button(btns, text="差分を作成", command=self.run)
        run_btn.grid(row=0, column=1, sticky="e", padx=6, pady=6)

    def run(self):
        self.ctx.save_user_path("diff_file_a", self.file_a.get())
        self.ctx.save_user_path("diff_file_b", self.file_b.get())
        self.ctx.save_user_path("diff_range_a", self.range_a.get())
        self.ctx.save_user_path("diff_range_b", self.range_b.get())
        self.ctx.save_user_path("diff_base_file", self.diff_base.get())
        self.ctx.save_user_path("diff_compare_formula", bool(self.var_formula.get()))
        self.ctx.save_user_path("diff_include_context", bool(self.var_ctx.get()))
        self.ctx.save_user_path("diff_compare_shapes", bool(self.var_shapes.get()))

        req = DiffRequest(
            file_a=self.file_a.get().strip(),
            file_b=self.file_b.get().strip(),
            range_a=self.range_a.get().strip(),
            range_b=self.range_b.get().strip(),
            base_file=self.diff_base.get().strip(),
            key_cols=[],
            compare_formula=self.var_formula.get(),
            include_context=self.var_ctx.get(),
            compare_shapes=self.var_shapes.get(),
        )

        try:
            out = run_diff(req, self.ctx, self.logger, self.log)
            self.log(f"[OK] 差分レポート: {out}")
        except Exception as e:
            self.log(f"[ERR] Diff: {e}")

    def make_report(self):
        self.log("[UI] select diff json")

        json_path = filedialog.askopenfilename(
            title="Select diff JSON",
            filetypes=[("JSON files", "*.json")],
        )

        if not json_path:
            self.log("[UI] canceled (json)")
            return

        self.log("[UI] select output html")

        out_path = filedialog.asksaveasfilename(
            title="Save HTML report",
            defaultextension=".html",
            filetypes=[("HTML files", "*.html")],
        )

        if not out_path:
            self.log("[UI] canceled (html)")
            return

        try:
            generate_html_report(Path(json_path), Path(out_path))
            self.log(f"[OK] HTMLレポート: {out_path}")
            messagebox.showinfo("レポート", f"HTMLレポートを出力しました。\n{out_path}")
        except Exception as e:
            self.log(f"[ERR] HTMLレポート: {e}")
            messagebox.showerror("レポート", f"HTML生成に失敗しました。\n{e}")

# =========================================================
# Count
# =========================================================
class CountTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, "Count")
        self.build()

    def build(self):
        frm = ttk.LabelFrame(self.tab, text="Count")
        frm.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="対象Excel").grid(row=0, column=0, sticky="w")
        self.files = tk.Entry(frm, width=70)
        self.files.grid(row=0, column=1, sticky="we", padx=4)
        self.files.insert(0, self.ctx.user_paths.get("count_files", ""))

        ttk.Button(
            frm, text="...", width=3,
            command=lambda: self.app.choose_files(
                self.files, "count_files",
                [("Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls")]
            )
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Label(frm, text="シート名（空=先頭）").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.sheet = tk.Entry(frm, width=30)
        self.sheet.grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))
        self.sheet.insert(0, self.ctx.user_paths.get("count_sheet", ""))

        ttk.Label(frm, text="開始セル").grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.start = tk.Entry(frm, width=12)
        self.start.grid(row=2, column=1, sticky="w", padx=4, pady=(4, 0))
        self.start.insert(0, self.ctx.user_paths.get("count_start", "B2"))

        self.dir = tk.StringVar(value=self.ctx.user_paths.get("count_dir", "row"))
        dfrm = ttk.Frame(frm)
        dfrm.grid(row=3, column=1, sticky="w", padx=4, pady=(4, 0))
        ttk.Radiobutton(dfrm, text="行方向", variable=self.dir, value="row").pack(side="left")
        ttk.Radiobutton(dfrm, text="列方向", variable=self.dir, value="col").pack(side="left", padx=(12, 0))

        ttk.Label(frm, text="許容空白数").grid(row=4, column=0, sticky="w", pady=(4, 0))
        self.tol = tk.Spinbox(frm, from_=0, to=1000, width=6)
        self.tol.grid(row=4, column=1, sticky="w", padx=4, pady=(4, 0))
        self.tol.delete(0, "end")
        self.tol.insert(0, str(self.ctx.user_paths.get("count_tolerate_blanks", 0)))

        self.mode = tk.StringVar(value=self.ctx.user_paths.get("count_mode", "jump"))
        mfrm = ttk.Frame(frm)
        mfrm.grid(row=5, column=1, sticky="w", padx=4, pady=(4, 0))
        ttk.Radiobutton(mfrm, text="高速", variable=self.mode, value="jump").pack(side="left")
        ttk.Radiobutton(mfrm, text="精密", variable=self.mode, value="scan").pack(side="left", padx=(12, 0))

        ttk.Button(self.tab, text="実行", command=self.run).grid(
            row=0, column=1, sticky="ne", padx=6, pady=6
        )

        self.tab.grid_columnconfigure(0, weight=1)

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
            filetypes=[("Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls")],
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
