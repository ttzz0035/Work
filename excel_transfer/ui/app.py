# excel_transfer/ui/app.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from models.dto import GrepRequest, DiffRequest, TransferRequest, CountRequest
from services.transfer import run_transfer_from_csvs
from services.grep import run_grep
from services.diff import run_diff
from services.count import run_count


class ExcelApp:
    def __init__(self, ctx, logger):
        self.ctx, self.logger = ctx, logger

        self.root = tk.Tk()
        self.root.title(ctx.labels["app_title"])

        # Notebook（タブ）
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=4, pady=2)

        self._build_tab_transfer()
        self._build_tab_grep()
        self._build_tab_diff()
        self._build_tab_count()

        # ログ領域（上詰め）
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=False, padx=4, pady=(2, 4))
        ttk.Label(log_frame, text="ログ").pack(anchor="w", pady=(0, 1))
        self.log = tk.Text(log_frame, height=8)
        self.log.pack(fill="both", expand=True)
        ttk.Button(log_frame, text="コピー", command=self._copy_log).pack(anchor="e", pady=(1, 0))

        # ウィンドウの最小サイズを現状に合わせる（最小限表示）
        self.root.update_idletasks()
        self.root.minsize(self.root.winfo_reqwidth(), self.root.winfo_reqheight())

    # ==========================
    # 転記タブ
    # ==========================
    def _build_tab_transfer(self):
        tab = ttk.Frame(self.nb, padding=(4, 2, 4, 2))
        self.nb.add(tab, text=self.ctx.labels["section_transfer"])

        ttk.Label(tab, text=self.ctx.labels["label_transfer_config"]).grid(row=0, column=0, sticky="w", padx=4, pady=2)

        width = self.ctx.app_settings["components"].get("transfer_config", {}).get("width", 70)
        self.transfer_entry = tk.Entry(tab, width=width)
        self.transfer_entry.grid(row=0, column=1, padx=4, pady=2, sticky="we")
        self.transfer_entry.insert(0, self.ctx.user_paths.get("transfer_config", ""))

        ttk.Button(
            tab,
            text="...",
            width=3,
            command=lambda: self._choose_files(self.transfer_entry, "transfer_config", [("CSV Files", "*.csv")]),
        ).grid(row=0, column=2, padx=2, pady=2)

        # 範囲外セルの扱い: チェックON=skip（継続保存）、OFF=error（即中止）
        init_skip = bool(self.ctx.user_paths.get("transfer_skip_oor", False))
        self.var_skip_oor = tk.BooleanVar(value=init_skip)
        ttk.Checkbutton(tab, text="範囲外セルはスキップ（警告のみ）", variable=self.var_skip_oor)\
            .grid(row=1, column=1, sticky="w", padx=4, pady=(0,2))

        # 実行ボタン
        ttk.Button(tab, text="実行", command=self.on_transfer).grid(row=0, column=3, rowspan=2, padx=4, pady=2)

        tab.grid_columnconfigure(1, weight=1)

    # ==========================
    # Grepタブ
    # ==========================
    def _build_tab_grep(self):
        tab = ttk.Frame(self.nb, padding=(4, 2, 4, 2))
        self.nb.add(tab, text=self.ctx.labels["section_grep"])

        ttk.Label(tab, text=self.ctx.labels["label_grep_root"]).grid(row=0, column=0, sticky="w", padx=4, pady=2)

        width_root = self.ctx.app_settings["components"].get("grep_root", {}).get("width", 70)
        self.grep_root = tk.Entry(tab, width=width_root)
        self.grep_root.grid(row=0, column=1, padx=4, pady=2, sticky="we")
        self.grep_root.insert(0, self.ctx.user_paths.get("grep_root", ""))

        ttk.Button(
            tab, text="...", width=3, command=lambda: self._choose_dir(self.grep_root, "grep_root")
        ).grid(row=0, column=2, padx=2, pady=2)

        ttk.Label(tab, text=self.ctx.labels["label_grep_keyword"]).grid(row=1, column=0, sticky="w", padx=4, pady=2)

        width_kw = self.ctx.app_settings["components"].get("grep_keyword", {}).get("width", 40)
        self.grep_kw = tk.Entry(tab, width=width_kw)
        self.grep_kw.grid(row=1, column=1, padx=4, pady=2, sticky="we")
        self.grep_kw.insert(0, self.ctx.user_paths.get("grep_keyword", ""))

        init_ic = bool(self.ctx.user_paths.get("grep_ignore_case", True))
        init_rx = bool(self.ctx.user_paths.get("grep_use_regex", False))
        self.var_ic = tk.BooleanVar(value=init_ic)
        self.var_rx = tk.BooleanVar(value=init_rx)
        ttk.Checkbutton(tab, text=self.ctx.labels["check_ignore_case"], variable=self.var_ic).grid(
            row=1, column=2, padx=2, pady=2
        )
        ttk.Checkbutton(tab, text="正規表現", variable=self.var_rx).grid(row=1, column=3, padx=2, pady=2)

        ttk.Button(tab, text="実行", command=self.on_grep).grid(row=0, column=4, rowspan=2, padx=4, pady=2)
        tab.grid_columnconfigure(1, weight=1)

    # ==========================
    # Diffタブ
    # ==========================
    def _build_tab_diff(self):
        tab = ttk.Frame(self.nb, padding=(4, 2, 4, 0))
        self.nb.add(tab, text=self.ctx.labels["section_diff"])

        ttk.Label(tab, text=self.ctx.labels["label_diff_file_a"]).grid(row=0, column=0, sticky="w", padx=4, pady=2)

        width_a = self.ctx.app_settings["components"].get("diff_file_a", {}).get("width", 70)
        self.diff_a = tk.Entry(tab, width=width_a)
        self.diff_a.grid(row=0, column=1, padx=4, pady=2, sticky="we")
        self.diff_a.insert(0, self.ctx.user_paths.get("diff_file_a", ""))

        ttk.Button(tab, text="...", width=3, command=lambda: self._choose_file(self.diff_a, "diff_file_a")).grid(
            row=0, column=2, padx=2, pady=2
        )

        ttk.Label(tab, text=self.ctx.labels["label_diff_file_b"]).grid(row=1, column=0, sticky="w", padx=4, pady=2)

        width_b = self.ctx.app_settings["components"].get("diff_file_b", {}).get("width", 70)
        self.diff_b = tk.Entry(tab, width=width_b)
        self.diff_b.grid(row=1, column=1, padx=4, pady=2, sticky="we")
        self.diff_b.insert(0, self.ctx.user_paths.get("diff_file_b", ""))

        ttk.Button(tab, text="...", width=3, command=lambda: self._choose_file(self.diff_b, "diff_file_b")).grid(
            row=1, column=2, padx=2, pady=2
        )

        ttk.Label(tab, text=self.ctx.labels["label_diff_key_cols"]).grid(row=2, column=0, sticky="w", padx=4, pady=2)

        width_keys = self.ctx.app_settings["components"].get("diff_key_cols", {}).get("width", 40)
        self.diff_keys = tk.Entry(tab, width=width_keys)
        self.diff_keys.grid(row=2, column=1, padx=4, pady=2, sticky="we")
        self.diff_keys.insert(0, self.ctx.user_paths.get("diff_key_cols", ""))

        init_formula = bool(self.ctx.user_paths.get("diff_compare_formula", False))
        init_ctx = bool(self.ctx.user_paths.get("diff_include_context", True))
        init_shapes = bool(self.ctx.user_paths.get("diff_compare_shapes", False))
        self.var_formula = tk.BooleanVar(value=init_formula)
        self.var_ctx = tk.BooleanVar(value=init_ctx)
        self.var_shapes = tk.BooleanVar(value=init_shapes)

        ttk.Checkbutton(tab, text=self.ctx.labels["check_compare_formula"], variable=self.var_formula).grid(
            row=2, column=2, padx=2, pady=2
        )
        ttk.Checkbutton(tab, text="ジャンプリンク/コンテキスト", variable=self.var_ctx).grid(
            row=3, column=1, sticky="w", padx=4, pady=(0, 2)
        )
        ttk.Checkbutton(tab, text="図・画像も比較", variable=self.var_shapes).grid(
            row=3, column=2, sticky="w", padx=4, pady=(0, 2)
        )

        ttk.Button(tab, text="実行", command=self.on_diff).grid(row=0, column=3, rowspan=4, padx=4, pady=2)
        tab.grid_columnconfigure(1, weight=1)

    # ==========================
    # Countタブ
    # ==========================
    def _build_tab_count(self):
        tab = ttk.Frame(self.nb, padding=(4, 2, 4, 2))
        self.nb.add(tab, text="Count")

        ttk.Label(tab, text="対象Excel").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.count_files = tk.Entry(tab, width=70)
        self.count_files.grid(row=0, column=1, padx=4, pady=2, sticky="we")
        self.count_files.insert(0, self.ctx.user_paths.get("count_files", ""))

        ttk.Button(
            tab,
            text="...",
            width=3,
            command=lambda: self._choose_files(
                self.count_files, "count_files", [("Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls")]
            ),
        ).grid(row=0, column=2, padx=2, pady=2)

        ttk.Label(tab, text="シート名（空=先頭）").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self.count_sheet = tk.Entry(tab, width=30)
        self.count_sheet.grid(row=1, column=1, padx=4, pady=2, sticky="w")
        self.count_sheet.insert(0, self.ctx.user_paths.get("count_sheet", ""))

        ttk.Label(tab, text="開始セル").grid(row=2, column=0, sticky="w", padx=4, pady=2)
        self.count_start = tk.Entry(tab, width=12)
        self.count_start.grid(row=2, column=1, padx=4, pady=2, sticky="w")
        self.count_start.insert(0, self.ctx.user_paths.get("count_start", "B2"))

        ttk.Label(tab, text="方向").grid(row=3, column=0, sticky="w", padx=4, pady=2)
        self.count_dir = tk.StringVar(value=self.ctx.user_paths.get("count_dir", "row"))
        ttk.Radiobutton(tab, text="行方向", variable=self.count_dir, value="row").grid(
            row=3, column=1, sticky="w", padx=4, pady=2
        )
        ttk.Radiobutton(tab, text="列方向", variable=self.count_dir, value="col").grid(
            row=3, column=1, sticky="w", padx=80, pady=2
        )

        ttk.Label(tab, text="許容空白数（高速/精密 共通）").grid(row=4, column=0, sticky="w", padx=4, pady=2)
        self.count_tol = tk.Spinbox(tab, from_=0, to=1000, width=6)
        self.count_tol.grid(row=4, column=1, padx=4, pady=2, sticky="w")
        self.count_tol.delete(0, "end")
        self.count_tol.insert(0, str(self.ctx.user_paths.get("count_tolerate_blanks", 0)))

        self.count_mode = tk.StringVar(value=self.ctx.user_paths.get("count_mode", "jump"))
        ttk.Radiobutton(tab, text="高速（Ctrl+矢印）", variable=self.count_mode, value="jump").grid(
            row=5, column=1, sticky="w", padx=4, pady=2
        )
        ttk.Radiobutton(tab, text="精密（逐次スキャン）", variable=self.count_mode, value="scan").grid(
            row=5, column=1, sticky="w", padx=160, pady=2
        )

        # 右下固定
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_columnconfigure(3, weight=1)
        tab.grid_rowconfigure(6, weight=1)

        run_btn = ttk.Button(tab, text="実行", command=self.on_count)
        run_btn.grid(row=7, column=3, padx=6, pady=6, sticky="se")

    # ==========================
    # 共通ユーティリティ
    # ==========================
    def _append_log(self, msg: str):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)

    def _copy_log(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.log.get("1.0", tk.END))
        messagebox.showinfo("クリップボード", "ログをコピーしました。")

    def _choose_files(self, entry: tk.Entry, key: str, filetypes):
        initdir = self.ctx.default_dir_for(entry.get())
        paths = filedialog.askopenfilenames(initialdir=initdir, filetypes=filetypes)
        if paths:
            entry.delete(0, tk.END)
            entry.insert(0, "?".join(paths))
            self.ctx.save_user_path(key, entry.get())

    def _choose_file(self, entry: tk.Entry, key: str):
        initdir = self.ctx.default_dir_for(entry.get())
        path = filedialog.askopenfilename(
            initialdir=initdir, filetypes=[("Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls")]
        )
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            self.ctx.save_user_path(key, path)

    def _choose_dir(self, entry: tk.Entry, key: str):
        initdir = self.ctx.default_dir_for(entry.get())
        path = filedialog.askdirectory(initialdir=initdir)
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            self.ctx.save_user_path(key, path)

    # ==========================
    # ハンドラ（実行時に設定を保存）
    # ==========================
    def on_transfer(self):
        # 保存
        self.ctx.save_user_path("transfer_config", self.transfer_entry.get())
        self.ctx.save_user_path("transfer_skip_oor", bool(self.var_skip_oor.get()))

        # 実行
        paths = [p for p in self.transfer_entry.get().split("?") if p]
        mode = "skip" if self.var_skip_oor.get() else "error"
        req = TransferRequest(csv_paths=paths, out_of_range_mode=mode)
        try:
            note = run_transfer_from_csvs(req, self.ctx, self.logger, self._append_log)
            self._append_log(f"[OK] 転記完了: {note}")
        except Exception as e:
            self._append_log(f"[ERR] 転記: {e}")

    def on_grep(self):
        # 保存
        self.ctx.save_user_path("grep_root", self.grep_root.get().strip())
        self.ctx.save_user_path("grep_keyword", self.grep_kw.get().strip())
        self.ctx.save_user_path("grep_ignore_case", bool(self.var_ic.get()))
        self.ctx.save_user_path("grep_use_regex", bool(self.var_rx.get()))

        # 実行
        req = GrepRequest(
            root_dir=self.grep_root.get().strip(),
            keyword=self.grep_kw.get().strip(),
            ignore_case=self.var_ic.get(),
            use_regex=self.var_rx.get(),
        )
        try:
            out, cnt = run_grep(req, self.ctx, self.logger, self._append_log)
            self._append_log(f"[OK] Grep結果: {out} / {cnt}件")
        except Exception as e:
            self._append_log(f"[ERR] Grep: {e}")

    def on_diff(self):
        # 保存
        self.ctx.save_user_path("diff_file_a", self.diff_a.get().strip())
        self.ctx.save_user_path("diff_file_b", self.diff_b.get().strip())
        self.ctx.save_user_path("diff_key_cols", self.diff_keys.get().strip())
        self.ctx.save_user_path("diff_compare_formula", bool(self.var_formula.get()))
        self.ctx.save_user_path("diff_include_context", bool(self.var_ctx.get()))
        self.ctx.save_user_path("diff_compare_shapes", bool(self.var_shapes.get()))

        # 実行
        req = DiffRequest(
            file_a=self.diff_a.get().strip(),
            file_b=self.diff_b.get().strip(),
            key_cols=[c.strip() for c in self.diff_keys.get().split(",") if c.strip()],
            compare_formula=self.var_formula.get(),
            include_context=self.var_ctx.get(),
            compare_shapes=self.var_shapes.get(),
        )
        try:
            out = run_diff(req, self.ctx, self.logger, self._append_log)
            self._append_log(f"[OK] 差分レポート: {out}")
        except Exception as e:
            self._append_log(f"[ERR] Diff: {e}")

    def on_count(self):
        # 保存
        self.ctx.save_user_path("count_files", self.count_files.get().strip())
        self.ctx.save_user_path("count_sheet", self.count_sheet.get().strip())
        self.ctx.save_user_path("count_start", self.count_start.get().strip())
        self.ctx.save_user_path("count_dir", self.count_dir.get())
        self.ctx.save_user_path("count_tolerate_blanks", int(self.count_tol.get() or 0))
        self.ctx.save_user_path("count_mode", self.count_mode.get())

        # 実行
        files = [p for p in self.count_files.get().split("?") if p]
        req = CountRequest(
            files=files,
            sheet=self.count_sheet.get().strip(),
            start_cell=self.count_start.get().strip(),
            direction=self.count_dir.get(),
            tolerate_blanks=int(self.count_tol.get() or 0),
            mode=self.count_mode.get(),
        )
        try:
            out = run_count(req, self.ctx, self.logger, self._append_log)
            self._append_log(f"[OK] Count結果: {out}")
        except Exception as e:
            self._append_log(f"[ERR] Count: {e}")

    # ==========================
    # 起動
    # ==========================
    def run(self):
        self.root.mainloop()
