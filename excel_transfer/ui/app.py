# ============================================
# excel_transfer/ui/app.py
# ============================================
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import json

from ui.dialogs.progress_dialog import ProgressDialog
from ui.dialogs.transfer_csv_builder import TransferCsvBuilderDialog
from ui.dialogs.preview_dialog import PreviewReplaceDialog
from ui.components.excel_canvas import ExcelCanvas
from models.dto import GrepRequest, DiffRequest, TransferRequest, CountRequest
from models.dto import DiffResult

from services.transfer import run_transfer_from_csvs
from services.grep import run_grep
from services.diff import run_diff
from services.count import run_count
from services.excel_view_service import ExcelViewService
from outputs.excel_diff_html import generate_html_report
from outputs.excel_grep_html import generate_grep_html_report


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
    # 共通実行エントリ（実行時に必ず再評価）
    # =========================================================
    def run(self):
        name = self.__class__.__name__
        self.log(f"[RUN] {name} start")

        try:
            self._reevaluate_license()
        except Exception as e:
            self.log(f"[LICENSE][BLOCK] {e}")
            if self.logger:
                self.logger.error(f"[LICENSE][BLOCK] {e}", exc_info=True)
            return

        self.app.show_progress(self.ctx.labels["msg_progress_running"].format(name=name))

        try:
            self._run_impl()
            self.log(f"[RUN] {name} done")
        except Exception as e:
            self.log(f"[ERR] {name}: {e}")
            if self.logger:
                self.logger.error(f"[RUN][ERR] {name}: {e}", exc_info=True)
        finally:
            self.app.hide_progress()

    # =========================================================
    # 実行時ライセンス再評価（唯一の追加責務）
    # =========================================================
    def _reevaluate_license(self):
        try:
            from licensing.build_config import DEBUG_BUILD
            debug = bool(DEBUG_BUILD)
        except Exception:
            debug = False

        # license_manager が無い
        if not hasattr(self.ctx, "license_manager") or self.ctx.license_manager is None:
            if debug:
                self.log("[LICENSE][DEBUG] license_manager missing -> skipped")
                if self.logger:
                    self.logger.warning(
                        "[LICENSE] license_manager is missing (debug build)"
                    )
                return
            raise AttributeError("ctx.license_manager is required")

        lm = self.ctx.license_manager

        try:
            state = lm.get_state()
        except Exception as ex:
            if debug:
                self.log(f"[LICENSE][DEBUG] get_state failed -> skipped ({ex})")
                if self.logger:
                    self.logger.warning(
                        "[LICENSE] get_state failed (debug build)", exc_info=True
                    )
                return
            raise

        self.ctx.license_status = state.status
        self.ctx.license_remaining_days = state.remaining_days

        if state.status == "subscribed":
            self.log("[LICENSE] subscribed")
            return

        if state.status == "trial":
            self.log(f"[LICENSE] trial remain={state.remaining_days}")
            return

        raise RuntimeError("License expired")

    # =========================================================
    # 各タブが実装
    # =========================================================
    def _run_impl(self):
        raise NotImplementedError("Tab must implement _run_impl()")


# =========================================================
# Transfer Tab
# =========================================================
class TransferTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_transfer"])
        self.entry = None
        self.var_skip = None
        self.build()

    def build(self):
        frm = ttk.LabelFrame(self.tab, text=self.ctx.labels["transfer_group"])
        frm.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text=self.ctx.labels["label_transfer_config"]).grid(row=0, column=0, sticky="w")
        self.entry = tk.Entry(frm, width=70)
        self.entry.grid(row=0, column=1, sticky="we", padx=4)
        self.entry.insert(0, self.ctx.user_paths.get("transfer_config", ""))

        ttk.Button(
            frm,
            text=self.ctx.labels["button_ellipsis"],
            width=3,
            command=lambda: self.app.choose_files(
                self.entry,
                "transfer_config",
                [(self.ctx.labels["filetype_csv"], "*.csv")],
            ),
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Button(
            frm,
            text=self.ctx.labels["button_transfer_create"],
            width=5,
            command=self.open_builder,
        ).grid(row=0, column=3, padx=(0, 4))

        self.var_skip = tk.BooleanVar(
            value=bool(self.ctx.user_paths.get("transfer_skip_oor", False))
        )
        ttk.Checkbutton(
            frm,
            text=self.ctx.labels["check_transfer_skip_oor"],
            variable=self.var_skip,
        ).grid(row=1, column=1, sticky="w", padx=4, pady=(2, 0))

        ttk.Button(self.tab, text=self.ctx.labels["button_run"], command=self.run).grid(
            row=1, column=1, sticky="ne", padx=6, pady=6
        )

        self.tab.grid_columnconfigure(0, weight=1)

    def _on_created_transfer_csv(self, path: str):
        if not self.entry:
            return
        self.entry.delete(0, tk.END)
        self.entry.insert(0, path)
        self.ctx.save_user_path("transfer_config", path)

    def open_builder(self):
        self.logger.info("[UI] open transfer csv builder")

        TransferCsvBuilderDialog(
            parent=self.tab,
            ctx=self.ctx,
            logger=self.logger,
            on_created=self._on_created_transfer_csv,
        )

    def _run_impl(self):
        self.ctx.save_user_path("transfer_config", self.entry.get())
        self.ctx.save_user_path("transfer_skip_oor", bool(self.var_skip.get()))

        paths = [p for p in self.entry.get().split("?") if p]
        mode = "skip" if self.var_skip.get() else "error"

        req = TransferRequest(
            csv_paths=paths,
            out_of_range_mode=mode,
        )

        note = run_transfer_from_csvs(req, self.ctx, self.logger, self.log)
        self.log(self.ctx.labels["msg_transfer_done"].format(note=note))


# =========================================================
# Grep Tab
# =========================================================
class GrepTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_grep"])
        self.root_entry = None
        self.kw_entry = None
        self.var_ic = None
        self.var_rx = None
        self.file_rx_entry = None
        self.sheet_rx_entry = None
        self.offset_row = None
        self.offset_col = None
        self.var_replace_enabled = None
        self.replace_entry = None
        self.var_mode = None
        self.rb_preview = None
        self.rb_auto = None
        self.build()

    def build(self):
        self.tab.grid_columnconfigure(0, weight=1)

        # --- main ---
        frm = ttk.LabelFrame(self.tab, text=self.ctx.labels["grep_group"])
        frm.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        frm.grid_columnconfigure(1, weight=1)

        # --- Root ---
        ttk.Label(frm, text=self.ctx.labels["label_grep_root"]).grid(row=0, column=0, sticky="w")
        self.root_entry = tk.Entry(frm, width=70)
        self.root_entry.grid(row=0, column=1, sticky="we", padx=4)
        self.root_entry.insert(0, self.ctx.user_paths.get("grep_root", ""))

        ttk.Button(
            frm, text=self.ctx.labels["button_ellipsis"], width=3,
            command=lambda: self.app.choose_dir(self.root_entry, "grep_root"),
        ).grid(row=0, column=2, padx=(0, 4))

        # --- Keyword ---
        ttk.Label(frm, text=self.ctx.labels["label_grep_keyword"]).grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )
        self.kw_entry = tk.Entry(frm, width=40)
        self.kw_entry.grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))
        self.kw_entry.insert(0, self.ctx.user_paths.get("grep_keyword", ""))

        # --- Options ---
        self.var_ic = tk.BooleanVar(value=bool(self.ctx.user_paths.get("grep_ignore_case", True)))
        self.var_rx = tk.BooleanVar(value=bool(self.ctx.user_paths.get("grep_use_regex", False)))

        opt = ttk.Frame(frm)
        opt.grid(row=2, column=1, sticky="w", padx=4, pady=(4, 0))
        ttk.Checkbutton(opt, text=self.ctx.labels["check_ignore_case"], variable=self.var_ic).pack(side="left")
        ttk.Checkbutton(opt, text=self.ctx.labels["check_grep_regex"], variable=self.var_rx).pack(side="left", padx=(8, 0))

        # --- File name regex ---
        ttk.Label(frm, text=self.ctx.labels["label_grep_file_regex"]).grid(row=3, column=0, sticky="w", pady=(8, 0))
        self.file_rx_entry = tk.Entry(frm, width=40)
        self.file_rx_entry.grid(row=3, column=1, sticky="w", padx=4, pady=(8, 0))
        self.file_rx_entry.insert(0, self.ctx.user_paths.get("grep_file_regex", ""))

        # --- Sheet regex ---
        ttk.Label(frm, text=self.ctx.labels["label_grep_sheet_regex"]).grid(
            row=4, column=0, sticky="w", pady=(8, 0)
        )
        self.sheet_rx_entry = tk.Entry(frm, width=30)
        self.sheet_rx_entry.grid(row=4, column=1, sticky="w", padx=4, pady=(8, 0))
        self.sheet_rx_entry.insert(0, self.ctx.user_paths.get("grep_sheet_regex", ""))

        # --- Offset ---
        ttk.Label(frm, text=self.ctx.labels["label_grep_offset"]).grid(row=5, column=0, sticky="w", pady=(8, 0))
        off = ttk.Frame(frm)
        off.grid(row=5, column=1, sticky="w", padx=4, pady=(8, 0))

        ttk.Label(off, text=self.ctx.labels["label_grep_offset_row"]).pack(side="left")
        self.offset_row = tk.Entry(off, width=6)
        self.offset_row.pack(side="left", padx=(4, 12))
        self.offset_row.insert(0, self.ctx.user_paths.get("grep_offset_row", "0"))

        ttk.Label(off, text=self.ctx.labels["label_grep_offset_col"]).pack(side="left")
        self.offset_col = tk.Entry(off, width=6)
        self.offset_col.pack(side="left")
        self.offset_col.insert(0, self.ctx.user_paths.get("grep_offset_col", "0"))

        # =================================================
        # Replace section
        # =================================================
        self.var_replace_enabled = tk.BooleanVar(
            value=bool(self.ctx.user_paths.get("grep_replace_enabled", False))
        )

        ttk.Checkbutton(
            frm,
            text=self.ctx.labels["check_grep_replace_enabled"],
            variable=self.var_replace_enabled,
            command=self._update_replace_state,
        ).grid(row=6, column=0, sticky="w", pady=(10, 0))

        self.replace_entry = tk.Entry(frm, width=40)
        self.replace_entry.grid(row=6, column=1, sticky="w", padx=4, pady=(10, 0))
        self.replace_entry.insert(0, self.ctx.user_paths.get("grep_replace", ""))

        ttk.Label(frm, text=self.ctx.labels["label_grep_replace_mode"]).grid(row=7, column=0, sticky="w", pady=(8, 0))
        self.var_mode = tk.StringVar(value=self.ctx.user_paths.get("grep_replace_mode", "preview"))

        mode = ttk.Frame(frm)
        mode.grid(row=7, column=1, sticky="w", padx=4, pady=(8, 0))
        self.rb_preview = ttk.Radiobutton(mode, text=self.ctx.labels["radio_grep_replace_preview"], value="preview", variable=self.var_mode)
        self.rb_auto = ttk.Radiobutton(mode, text=self.ctx.labels["radio_grep_replace_auto"], value="auto", variable=self.var_mode)
        self.rb_preview.pack(side="left")
        self.rb_auto.pack(side="left", padx=(12, 0))

        # --- buttons（DiffTab と同配置） ---
        btns = ttk.Frame(self.tab)
        btns.grid(row=1, column=0, sticky="we", padx=4, pady=(4, 0))
        btns.grid_columnconfigure(0, weight=1)

        ttk.Button(btns, text=self.ctx.labels["button_html_report"], command=self.make_report).grid(
            row=0, column=0, sticky="w", padx=6, pady=6
        )

        ttk.Button(btns, text=self.ctx.labels["button_run"], command=self.run).grid(
            row=0, column=1, sticky="e", padx=6, pady=6
        )

        self._update_replace_state()

    # -------------------------------------------------
    # Replace state
    # -------------------------------------------------
    def _update_replace_state(self):
        enabled = self.var_replace_enabled.get()
        state = "normal" if enabled else "disabled"
        self.replace_entry.configure(state=state)
        self.rb_preview.configure(state=state)
        self.rb_auto.configure(state=state)
        self.logger.info(f"[GREP][UI] replace_enabled={enabled}")

    # -------------------------------------------------
    # run
    # -------------------------------------------------
    def run(self):
        try:
            self._run_impl()
        except Exception as e:
            self.log(f"[ERR] Grep: {e}")

    def _run_impl(self):
        # save
        self.ctx.save_user_path("grep_root", self.root_entry.get())
        self.ctx.save_user_path("grep_keyword", self.kw_entry.get())
        self.ctx.save_user_path("grep_ignore_case", bool(self.var_ic.get()))
        self.ctx.save_user_path("grep_use_regex", bool(self.var_rx.get()))
        self.ctx.save_user_path("grep_file_regex", self.file_rx_entry.get())
        self.ctx.save_user_path("grep_sheet_regex", self.sheet_rx_entry.get())
        self.ctx.save_user_path("grep_offset_row", self.offset_row.get())
        self.ctx.save_user_path("grep_offset_col", self.offset_col.get())
        self.ctx.save_user_path("grep_replace_enabled", self.var_replace_enabled.get())
        self.ctx.save_user_path("grep_replace", self.replace_entry.get())
        self.ctx.save_user_path("grep_replace_mode", self.var_mode.get())

        req = GrepRequest(
            root_dir=self.root_entry.get().strip(),
            keyword=self.kw_entry.get().strip(),
            ignore_case=self.var_ic.get(),
            use_regex=self.var_rx.get(),
            file_name_regex=self.file_rx_entry.get().strip() or None,
            sheet_name_regex=self.sheet_rx_entry.get().strip() or None,
            offset_row=int(self.offset_row.get()),
            offset_col=int(self.offset_col.get()),
            replace_enabled=self.var_replace_enabled.get(),
            replace_pattern=self.replace_entry.get(),
            replace_mode=self.var_mode.get(),
        )

        # 1st run (preview build / search)
        out, cnt = run_grep(req, self.ctx, self.logger, self.log)

        # --- preview ---
        if self.var_replace_enabled.get() and req.replace_mode == "preview":
            with open(out, "r", encoding="utf-8") as f:
                data = json.load(f)

            items = []
            for fentry in data.get("files", []):
                for sentry in fentry.get("sheets", []):
                    items.extend(sentry.get("items", []))

            if not items:
                self.log("[GREP] preview: no replace items")
                self.log(f"[OK] Grep結果: {out} / {cnt}件")
                return

            dlg = PreviewReplaceDialog(
                self.app.root,
                items,
                req.replace_pattern,
                ctx=self.ctx,
            )
            if dlg.result is None:
                self.log("[GREP] preview canceled")
                self.log(f"[OK] Grep結果: {out} / {cnt}件")
                return

            self.log(f"[GREP] preview accepted items={len(dlg.result)}")

            # ★ 再実行用フラグ
            req.preview_accepted = True

            # ★ 2nd run (apply replace)
            run_grep(req, self.ctx, self.logger, self.log)

        self.log(f"[OK] Grep結果: {out} / {cnt}件")

    # -------------------------------------------------
    # report（HTMLはボタンのみ）
    # -------------------------------------------------
    def make_report(self):
        json_path = filedialog.askopenfilename(
            title=self.ctx.labels["dialog_select_grep_json_title"],
            filetypes=[(self.ctx.labels["filetype_json"], "*.json")],
        )
        if not json_path:
            return

        out_path = filedialog.asksaveasfilename(
            title=self.ctx.labels["dialog_save_html_title"],
            defaultextension=".html",
            filetypes=[(self.ctx.labels["filetype_html"], "*.html")],
        )
        if not out_path:
            return

        generate_grep_html_report(
            Path(json_path),
            Path(out_path),
            self.ctx.labels,
        )


# =========================================================
# Diff Tab（DiffResult 対応）
# =========================================================
class DiffTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_diff"])
        self.file_a = None
        self.file_b = None
        self.range_a = None
        self.range_b = None
        self.diff_base = None
        self.var_formula = None
        self.var_ctx = None
        self.var_shapes = None
        self.sheet_mode = None
        self.report_btn = None
        self._last_result: DiffResult | None = None
        self.build()

    def build(self):
        self.tab.grid_columnconfigure(0, weight=1)

        files = ttk.LabelFrame(self.tab, text=self.ctx.labels["diff_group_files"])
        files.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        files.grid_columnconfigure(1, weight=1)

        ttk.Label(files, text=self.ctx.labels["label_diff_file_a"]).grid(row=0, column=0, sticky="w")
        self.file_a = tk.Entry(files, width=70)
        self.file_a.grid(row=0, column=1, sticky="we", padx=4)
        self.file_a.insert(0, self.ctx.user_paths.get("diff_file_a", ""))

        ttk.Button(
            files,
            text=self.ctx.labels["button_ellipsis"],
            width=3,
            command=lambda: self.app.choose_file(self.file_a, "diff_file_a"),
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Label(files, text=self.ctx.labels["label_diff_file_b"]).grid(row=1, column=0, sticky="w")
        self.file_b = tk.Entry(files, width=70)
        self.file_b.grid(row=1, column=1, sticky="we", padx=4)
        self.file_b.insert(0, self.ctx.user_paths.get("diff_file_b", ""))

        ttk.Button(
            files,
            text=self.ctx.labels["button_ellipsis"],
            width=3,
            command=lambda: self.app.choose_file(self.file_b, "diff_file_b"),
        ).grid(row=1, column=2, padx=(0, 4))

        mid = ttk.Frame(self.tab)
        mid.grid(row=1, column=0, sticky="we", padx=4, pady=4)
        mid.grid_columnconfigure(0, weight=1)
        mid.grid_columnconfigure(1, weight=1)

        rng = ttk.LabelFrame(mid, text=self.ctx.labels["diff_group_range_required"])
        rng.grid(row=0, column=0, sticky="we", padx=(0, 4))

        ttk.Label(rng, text=self.ctx.labels["diff_label_range_a"]).grid(row=0, column=0, sticky="w")
        self.range_a = tk.Entry(rng, width=20)
        self.range_a.grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(rng, text=self.ctx.labels["diff_label_range_b"]).grid(row=1, column=0, sticky="w")
        self.range_b = tk.Entry(rng, width=20)
        self.range_b.grid(row=1, column=1, sticky="w", padx=4)

        base = ttk.LabelFrame(mid, text=self.ctx.labels["diff_group_base"])
        base.grid(row=0, column=1, sticky="we", padx=(4, 0))

        self.diff_base = tk.StringVar(value="B")
        ttk.Radiobutton(
            base,
            text=self.ctx.labels["diff_radio_base_b_default"],
            variable=self.diff_base,
            value="B",
        ).pack(anchor="w")

        ttk.Radiobutton(
            base,
            text=self.ctx.labels["diff_radio_base_a"],
            variable=self.diff_base,
            value="A",
        ).pack(anchor="w")

        opt = ttk.LabelFrame(self.tab, text=self.ctx.labels["diff_group_options"])
        opt.grid(row=2, column=0, sticky="we", padx=4, pady=4)

        self.var_formula = tk.BooleanVar()
        self.var_ctx = tk.BooleanVar(value=True)
        self.var_shapes = tk.BooleanVar()

        ttk.Checkbutton(opt, text=self.ctx.labels["check_compare_formula"], variable=self.var_formula).pack(side="left")
        ttk.Checkbutton(opt, text=self.ctx.labels["check_diff_include_context"], variable=self.var_ctx).pack(side="left")
        ttk.Checkbutton(opt, text=self.ctx.labels["check_diff_compare_shapes"], variable=self.var_shapes).pack(side="left")

        sheet = ttk.LabelFrame(self.tab, text=self.ctx.labels["diff_group_sheet_mode"])
        sheet.grid(row=3, column=0, sticky="we", padx=4, pady=4)

        self.sheet_mode = tk.StringVar(value="index")
        ttk.Radiobutton(
            sheet,
            text=self.ctx.labels["diff_radio_sheet_index_default"],
            variable=self.sheet_mode,
            value="index",
        ).pack(anchor="w")

        ttk.Radiobutton(
            sheet,
            text=self.ctx.labels["diff_radio_sheet_name"],
            variable=self.sheet_mode,
            value="name",
        ).pack(anchor="w")

        btns = ttk.Frame(self.tab)
        btns.grid(row=4, column=0, sticky="we", padx=4, pady=(4, 0))
        btns.grid_columnconfigure(0, weight=1)

        self.report_btn = ttk.Button(
            btns,
            text=self.ctx.labels["button_html_report"],
            command=self.make_report,
        )
        self.report_btn.grid(row=0, column=0, sticky="w", padx=6, pady=6)

        ttk.Button(
            btns,
            text=self.ctx.labels["button_diff_make"],
            command=self.run,
        ).grid(row=0, column=1, sticky="e", padx=6, pady=6)

    def _run_impl(self):
        req = DiffRequest(
            file_a=self.file_a.get().strip(),
            file_b=self.file_b.get().strip(),
            range_a=self.range_a.get().strip(),
            range_b=self.range_b.get().strip(),
            base_file=self.diff_base.get(),
            compare_formula=self.var_formula.get(),
            include_context=self.var_ctx.get(),
            compare_shapes=self.var_shapes.get(),
            sheet_mode=self.sheet_mode.get(),
        )

        result: DiffResult = run_diff(req, self.ctx, self.logger, self.log)
        self._last_result = result

        self.log(f"[OK] Diff file: {result.diff_path}")
        self.log(f"[OK] Diff JSON: {result.json_path}")

    def make_report(self):
        if not self._last_result:
            messagebox.showwarning(
                self.ctx.labels["msg_report_title"],
                self.ctx.labels["msg_diff_not_executed"],
            )
            return

        out_path = filedialog.asksaveasfilename(
            title=self.ctx.labels["dialog_save_html_title"],
            defaultextension=".html",
            filetypes=[(self.ctx.labels["filetype_html"], "*.html")],
        )
        if not out_path:
            return

        generate_html_report(
            Path(self._last_result.json_path),
            Path(out_path),
            self.ctx.labels,
        )

        self.log(f"[OK] HTML report: {out_path}")
        messagebox.showinfo(
            self.ctx.labels["msg_report_title"],
            self.ctx.labels["msg_report_done_html"].format(path=out_path),
        )

    # -------------------------------------------------
    # run
    # -------------------------------------------------
    def run(self):
        try:
            self._run_impl()
        except Exception as e:
            self.log(f"[ERR] Diff: {e}")

    def _run_impl(self):
        # 保存
        self.ctx.save_user_path("diff_file_a", self.file_a.get())
        self.ctx.save_user_path("diff_file_b", self.file_b.get())
        self.ctx.save_user_path("diff_range_a", self.range_a.get())
        self.ctx.save_user_path("diff_range_b", self.range_b.get())
        self.ctx.save_user_path("diff_base_file", self.diff_base.get())
        self.ctx.save_user_path("diff_compare_formula", bool(self.var_formula.get()))
        self.ctx.save_user_path("diff_include_context", bool(self.var_ctx.get()))
        self.ctx.save_user_path("diff_compare_shapes", bool(self.var_shapes.get()))
        self.ctx.save_user_path("diff_sheet_mode", self.sheet_mode.get())

        req = DiffRequest(
            file_a=self.file_a.get().strip(),
            file_b=self.file_b.get().strip(),
            range_a=self.range_a.get().strip(),
            range_b=self.range_b.get().strip(),
            base_file=self.diff_base.get().strip(),
            compare_formula=self.var_formula.get(),
            include_context=self.var_ctx.get(),
            compare_shapes=self.var_shapes.get(),
            sheet_mode=self.sheet_mode.get(),
        )

        out = run_diff(req, self.ctx, self.logger, self.log)
        self.log(f"[OK] 差分レポート: {out}")

    # -------------------------------------------------
    # report
    # -------------------------------------------------
    def make_report(self):
        self.log("[UI] select diff json")

        json_path = filedialog.askopenfilename(
            title=self.ctx.labels["dialog_select_diff_json_title"],
            filetypes=[(self.ctx.labels["filetype_json"], "*.json")],
        )
        if not json_path:
            self.log("[UI] canceled (json)")
            return

        self.log("[UI] select output html")

        out_path = filedialog.asksaveasfilename(
            title=self.ctx.labels["dialog_save_html_title"],
            defaultextension=".html",
            filetypes=[(self.ctx.labels["filetype_html"], "*.html")],
        )
        if not out_path:
            self.log("[UI] canceled (html)")
            return

        generate_html_report(
            Path(json_path),
            Path(out_path),
            self.ctx.labels,
        )
        self.log(f"[OK] HTMLレポート: {out_path}")
        messagebox.showinfo(
            self.ctx.labels["msg_report_title"],
            self.ctx.labels["msg_report_done_html"].format(path=out_path),
        )


# =========================================================
# Count Tab
# =========================================================
class CountTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_count"])
        self.files = None
        self.sheet = None
        self.start = None
        self.dir = None
        self.tol = None
        self.build()

    def build(self):
        frm = ttk.LabelFrame(self.tab, text=self.ctx.labels["count_group"])
        frm.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text=self.ctx.labels["label_count_files"]).grid(row=0, column=0, sticky="w")
        self.files = tk.Entry(frm, width=70)
        self.files.grid(row=0, column=1, sticky="we", padx=4)
        self.files.insert(0, self.ctx.user_paths.get("count_files", ""))

        ttk.Button(
            frm,
            text=self.ctx.labels["button_ellipsis"],
            width=3,
            command=lambda: self.app.choose_files(
                self.files,
                "count_files",
                [(self.ctx.labels["filetype_excel"], "*.xlsx;*.xlsm;*.xlsb;*.xls")],
            ),
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Label(frm, text=self.ctx.labels["label_count_sheet"]).grid(row=1, column=0, sticky="w")
        self.sheet = tk.Entry(frm, width=30)
        self.sheet.grid(row=1, column=1, sticky="w", padx=4)
        self.sheet.insert(0, self.ctx.user_paths.get("count_sheet", ""))

        ttk.Label(frm, text=self.ctx.labels["label_count_start"]).grid(row=2, column=0, sticky="w")
        self.start = tk.Entry(frm, width=12)
        self.start.grid(row=2, column=1, sticky="w", padx=4)
        self.start.insert(0, self.ctx.user_paths.get("count_start", "B2"))

        self.dir = tk.StringVar(value=self.ctx.user_paths.get("count_dir", "row"))
        dfrm = ttk.Frame(frm)
        dfrm.grid(row=3, column=1, sticky="w", padx=4)
        ttk.Radiobutton(dfrm, text=self.ctx.labels["radio_count_dir_row"], variable=self.dir, value="row").pack(side="left")
        ttk.Radiobutton(dfrm, text=self.ctx.labels["radio_count_dir_col"], variable=self.dir, value="col").pack(side="left", padx=(12, 0))

        ttk.Label(frm, text=self.ctx.labels["label_count_tolerate_blanks"]).grid(row=4, column=0, sticky="w")
        self.tol = tk.Spinbox(frm, from_=0, to=1000, width=6)
        self.tol.grid(row=4, column=1, sticky="w", padx=4)
        self.tol.insert(0, str(self.ctx.user_paths.get("count_tolerate_blanks", 0)))

        ttk.Button(self.tab, text=self.ctx.labels["button_run"], command=self.run).grid(
            row=0, column=1, sticky="ne", padx=6, pady=6
        )

    def _run_impl(self):
        self.ctx.save_user_path("count_files", self.files.get())
        self.ctx.save_user_path("count_sheet", self.sheet.get())
        self.ctx.save_user_path("count_start", self.start.get())
        self.ctx.save_user_path("count_dir", self.dir.get())
        self.ctx.save_user_path("count_tolerate_blanks", int(self.tol.get()))

        req = CountRequest(
            files=[p for p in self.files.get().split("?") if p],
            sheet=self.sheet.get().strip(),
            start_cell=self.start.get().strip(),
            direction=self.dir.get(),
            tolerate_blanks=int(self.tol.get()),
            mode="jump",
        )

        out = run_count(req, self.ctx, self.logger, self.log)
        self.log(f"[OK] Count結果: {out}")


# =========================================================
# Excel Viewer Tab（実行なし）
# =========================================================
class ExcelViewerTab(BaseTab):
    def __init__(self, app):
        super().__init__(app, app.ctx.labels["section_excel"])
        self.left_service = ExcelViewService(logger=self.logger)
        self.right_service = ExcelViewService(logger=self.logger)
        self.left_book = None
        self.left_sheet = None
        self.left_canvas = None
        self.right_book = None
        self.right_sheet = None
        self.right_canvas = None
        self.build()

    def _run_impl(self):
        pass

    def build(self):
        self.tab.grid_columnconfigure(0, weight=1)
        self.tab.grid_rowconfigure(0, weight=1)

        paned = ttk.Panedwindow(self.tab, orient="horizontal")
        paned.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        self._build_side(paned, is_left=True)
        self._build_side(paned, is_left=False)

    def _build_side(self, paned, is_left: bool):
        frm = ttk.Frame(paned)
        paned.add(frm, weight=1)
        frm.rowconfigure(1, weight=1)
        frm.columnconfigure(0, weight=1)

        title = self.ctx.labels["viewer_left"] if is_left else self.ctx.labels["viewer_right"]
        lf = ttk.LabelFrame(frm, text=title)
        lf.grid(row=0, column=0, sticky="we")
        lf.columnconfigure(1, weight=1)

        ttk.Label(lf, text=self.ctx.labels["viewer_label_book"]).grid(row=0, column=0, sticky="w")
        cmb_book = ttk.Combobox(lf, state="readonly", width=60)
        cmb_book.grid(row=0, column=1, sticky="we", padx=4)

        ttk.Button(
            lf,
            text=self.ctx.labels["viewer_button_add"],
            command=lambda side=is_left: self._add_books(side),
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Label(lf, text=self.ctx.labels["viewer_label_sheet"]).grid(row=1, column=0, sticky="w", pady=(4, 0))
        cmb_sheet = ttk.Combobox(lf, state="readonly", width=40)
        cmb_sheet.grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))

        svc = self.left_service if is_left else self.right_service
        canvas = ExcelCanvas(frm, service=svc, logger=self.logger)
        canvas.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

        cmb_book.bind("<<ComboboxSelected>>", lambda _e, side=is_left: self._on_book(side))
        cmb_sheet.bind("<<ComboboxSelected>>", lambda _e, side=is_left: self._on_sheet(side))

        if is_left:
            self.left_book = cmb_book
            self.left_sheet = cmb_sheet
            self.left_canvas = canvas
        else:
            self.right_book = cmb_book
            self.right_sheet = cmb_sheet
            self.right_canvas = canvas

    def _add_books(self, is_left: bool):
        paths = filedialog.askopenfilenames(
            initialdir=self.ctx.default_dir_for(""),
            filetypes=[(self.ctx.labels["filetype_excel"], "*.xlsx;*.xlsm;*.xlsb;*.xls")],
        )
        if not paths:
            return

        svc = self.left_service if is_left else self.right_service
        cmb_book = self.left_book if is_left else self.right_book
        cmb_sheet = self.left_sheet if is_left else self.right_sheet
        canvas = self.left_canvas if is_left else self.right_canvas

        svc.add_books(list(paths))
        cmb_book["values"] = svc.get_book_paths()
        if svc.get_current_book_path():
            cmb_book.set(svc.get_current_book_path())

        cmb_sheet["values"] = svc.get_sheet_names()
        if svc.get_current_sheet_name():
            cmb_sheet.set(svc.get_current_sheet_name())

        canvas.refresh()

    def _on_book(self, is_left: bool):
        svc = self.left_service if is_left else self.right_service
        cmb_book = self.left_book if is_left else self.right_book
        cmb_sheet = self.left_sheet if is_left else self.right_sheet
        canvas = self.left_canvas if is_left else self.right_canvas

        p = cmb_book.get().strip()
        if not p:
            return

        svc.select_book(p)
        cmb_sheet["values"] = svc.get_sheet_names()
        if svc.get_current_sheet_name():
            cmb_sheet.set(svc.get_current_sheet_name())

        canvas.refresh()

    def _on_sheet(self, is_left: bool):
        svc = self.left_service if is_left else self.right_service
        cmb_sheet = self.left_sheet if is_left else self.right_sheet
        canvas = self.left_canvas if is_left else self.right_canvas

        s = cmb_sheet.get().strip()
        if not s:
            return

        svc.select_sheet(s)
        canvas.refresh()


# =========================================================
# App
# =========================================================
class ExcelApp:
    def __init__(self, ctx, logger):
        self.ctx, self.logger = ctx, logger
        self._progress = None

        self.root = tk.Tk()
        self.root.title(ctx.labels["app_title"])

        # ----------------------------------------
        # Menu
        # ----------------------------------------
        self._build_menu()

        # ----------------------------------------
        # Tabs
        # ----------------------------------------
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=4, pady=2)

        self.excel_tab = ExcelViewerTab(self)
        self.transfer_tab = TransferTab(self)
        self.grep_tab = GrepTab(self)
        self.diff_tab = DiffTab(self)
        self.count_tab = CountTab(self)

        # ----------------------------------------
        # Log area
        # ----------------------------------------
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=False, padx=4, pady=(2, 4))
        ttk.Label(log_frame, text=self.ctx.labels["label_log"]).pack(anchor="w")
        self.log = tk.Text(log_frame, height=8)
        self.log.pack(fill="both", expand=True)

    # =====================================================
    # Menu
    # =====================================================
    def _build_menu(self):
        menubar = tk.Menu(self.root)

        # -------------------------------
        # License
        # -------------------------------
        license_menu = tk.Menu(menubar, tearoff=0)
        license_menu.add_command(
            label=self.ctx.labels["menu_third_party_licenses"],
            command=self._show_third_party_licenses,
        )
        menubar.add_cascade(
            label=self.ctx.labels["menu_license"],
            menu=license_menu,
        )

        # -------------------------------
        # Language
        # -------------------------------
        lang_menu = tk.Menu(menubar, tearoff=0)
        for lang in ("ja", "en"):
            lang_menu.add_command(
                label=lang.upper(),
                command=lambda l=lang: self._change_language(l),
            )
        menubar.add_cascade(label="Language", menu=lang_menu)
        
        # -------------------------------
        # About
        # -------------------------------
        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(
            label=self.ctx.labels["menu_about"],
            command=self._show_about,
        )
        menubar.add_cascade(
            label=self.ctx.labels["menu_about"],
            menu=about_menu,
        )

        self.root.config(menu=menubar)

    # =====================================================
    # Language
    # =====================================================
    def _change_language(self, lang: str):
        cur = getattr(self.ctx, "lang", "ja")
        if cur == lang:
            return

        self.logger.info(
            f"[UI] language change requested: {cur} -> {lang} (apply on next start)"
        )

        # 保存のみ（即時反映しない）
        self.ctx.lang = lang
        self.ctx.save_user_path("app_lang", lang)

        # ユーザー通知（ログのみ）
        self.append_log(
            self.ctx.labels["msg_language_apply_next_start"].format(lang=lang)
        )

    def _restart_app(self):
        self.logger.info("[UI] restart app for language change")
        self.root.destroy()
        new_app = ExcelApp(self.ctx, self.logger)
        new_app.run()

    # =====================================================
    # License
    # =====================================================
    def _show_third_party_licenses(self):
        self.logger.info("[UI] open Third Party Licenses")

        text_data = getattr(self.ctx, "third_party_licenses_text", None)
        if not text_data:
            raise RuntimeError("third_party_licenses_text is not loaded")

        win = tk.Toplevel(self.root)
        win.title(self.ctx.labels["menu_third_party_licenses"])
        win.geometry(self.ctx.labels["window_third_party_licenses_geometry"])

        txt = tk.Text(win, wrap="word")
        txt.insert("1.0", text_data)
        txt.config(state="disabled")

        yscroll = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        txt.config(yscrollcommand=yscroll.set)

        txt.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

    # =====================================================
    # About
    # =====================================================
    def _show_about(self):
        self.logger.info("[UI] open About")

        win = tk.Toplevel(self.root)
        win.title(self.ctx.labels["about_title"])
        win.geometry(self.ctx.labels["window_about_geometry"])
        win.resizable(False, False)

        frame = ttk.Frame(win, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=self.ctx.labels["app_title"],
            font=("", 14, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        ttk.Label(
            frame,
            text=self.ctx.labels["about_version"].format(
                version=self.ctx.app_version
            ),
        ).pack(anchor="w")

        ttk.Label(
            frame,
            text=self.ctx.labels["about_copyright"],
        ).pack(anchor="w", pady=(10, 0))

    # =====================================================
    # Utilities
    # =====================================================
    def append_log(self, msg: str):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)

    def choose_files(self, entry, key, filetypes):
        paths = filedialog.askopenfilenames(filetypes=filetypes)
        if paths:
            entry.delete(0, tk.END)
            entry.insert(0, "?".join(paths))
            self.ctx.save_user_path(key, entry.get())

    def choose_file(self, entry, key):
        path = filedialog.askopenfilename(
            filetypes=[(self.ctx.labels["filetype_excel"], "*.xlsx;*.xlsm;*.xlsb;*.xls")]
        )
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            self.ctx.save_user_path(key, path)

    def choose_dir(self, entry, key):
        path = filedialog.askdirectory()
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            self.ctx.save_user_path(key, path)

    def run(self):
        self.root.mainloop()

    # -----------------------------------------
    # Progress control
    # -----------------------------------------
    def show_progress(self, message: str):
        if self._progress:
            return
        self.nb.state(["disabled"])
        self._progress = ProgressDialog(self.root, self.ctx, message)

    def hide_progress(self):
        if self._progress:
            self._progress.close()
            self._progress = None
        self.nb.state(["!disabled"])
