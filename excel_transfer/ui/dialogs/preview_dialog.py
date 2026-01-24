# =========================================================
# Preview Replace Dialog
# =========================================================
import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Any


class PreviewReplaceDialog(tk.Toplevel):
    def __init__(
        self,
        master,
        items: List[Dict[str, Any]],
        replace_pattern: str,
    ):
        super().__init__(master)
        self.title("置換プレビュー（Target に適用される文字列）")
        self.geometry("1200x520")
        self.resizable(True, True)

        self.items = items
        self.replace_pattern = replace_pattern
        self.result: List[Dict[str, Any]] | None = None

        self._build()
        self._populate()

        self.transient(master)
        self.grab_set()
        self.wait_window(self)

    # -------------------------------------------------
    # UI
    # -------------------------------------------------
    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        cols = (
            "checked",
            "file",
            "sheet",
            "hit_pos",
            "target_pos",
            "before",
            "replace",
        )

        self.tree = ttk.Treeview(
            self,
            columns=cols,
            show="headings",
            selectmode="none",
        )

        self.tree.heading("checked", text="✔")
        self.tree.heading("file", text="File")
        self.tree.heading("sheet", text="Sheet")
        self.tree.heading("hit_pos", text="Hit Cell")
        self.tree.heading("target_pos", text="Target Cell（置換対象）")
        self.tree.heading("before", text="Before（Target 現在値）")
        self.tree.heading("replace", text="Replace With（置換文字列）")

        self.tree.column("checked", width=40, anchor="center")
        self.tree.column("file", width=220)
        self.tree.column("sheet", width=120)
        self.tree.column("hit_pos", width=120)
        self.tree.column("target_pos", width=150)
        self.tree.column("before", width=260)
        self.tree.column("replace", width=260)

        ysb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")

        btns = ttk.Frame(self)
        btns.grid(row=2, column=0, sticky="e", padx=8, pady=8)

        ttk.Button(btns, text="全選択", command=self._select_all).pack(side="left", padx=4)
        ttk.Button(btns, text="全解除", command=self._clear_all).pack(side="left", padx=4)
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="left", padx=12)
        ttk.Button(btns, text="OK", command=self._ok).pack(side="left")

        self.tree.bind("<Button-1>", self._toggle_check)

    # -------------------------------------------------
    # populate
    # -------------------------------------------------
    def _populate(self):
        for i, it in enumerate(self.items):
            hit = it["hit"]
            target = it["target"]

            self.tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    "✔" if it.get("checked") else "",
                    it.get("file", ""),
                    it.get("sheet", ""),
                    f"R{hit['row']}C{hit['col']}",
                    f"R{target['row']}C{target['col']}",
                    it.get("before", ""),
                    self.replace_pattern,
                ),
            )

    # -------------------------------------------------
    # handlers
    # -------------------------------------------------
    def _toggle_check(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return

        row = self.tree.identify_row(event.y)
        if not row:
            return

        idx = int(row)
        cur = self.items[idx].get("checked", False)
        self.items[idx]["checked"] = not cur
        self.tree.set(row, "checked", "✔" if not cur else "")

    def _select_all(self):
        for i, it in enumerate(self.items):
            it["checked"] = True
            self.tree.set(str(i), "checked", "✔")

    def _clear_all(self):
        for i, it in enumerate(self.items):
            it["checked"] = False
            self.tree.set(str(i), "checked", "")

    def _ok(self):
        self.result = [it for it in self.items if it.get("checked")]
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()
