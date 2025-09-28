# Apps/APAParser/app.py
from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import sys
import json
import csv
import xml.etree.ElementTree as ET

# --- プロジェクトRootを sys.path に追加 ---
ROOT = Path(__file__).resolve().parents[1]  # Root/
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from apa_parser import APAParser, APAFormatError

APP_NAME = "APA Parser (UI)"
PREF_PATH = Path.home() / ".apa_parser_ui.json"   # 終了時保存/起動時復元


class FileListManager:
    """最小限のファイル選択管理（Common なし運用）"""
    def __init__(self):
        self._files: list[Path] = []

    def add(self, paths: list[Path]):
        seen = set(self._files)
        for p in paths:
            p = Path(p).resolve()
            if p.is_file() and p not in seen:
                self._files.append(p)
                seen.add(p)

    def remove_by_indices(self, indices: list[int]):
        for i in sorted(indices, reverse=True):
            if 0 <= i < len(self._files):
                self._files.pop(i)

    def list(self) -> list[Path]:
        return list(self._files)

    def count(self) -> int:
        return len(self._files)


class FileSelectApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(10,10))
        self.master.title(APP_NAME)

        # 状態
        self._manager = FileListManager()
        self._last_results = []  # list[APAParseResult]
        self.output_type  = tk.StringVar(value="xml")    # "xml" | "csv"

        self._build_ui()
        self._bind_shortcuts()
        self._load_prefs()  # 起動時復元
        self._refresh_buttons()
        self.pack(fill="both", expand=True)

        # 終了時に保存
        self.master.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- UI 構築 ----------------
    def _build_ui(self):
        # ツールバー（左：選択操作 / 右：アクション）
        bar = ttk.Frame(self); bar.grid(row=0, column=0, sticky="we", pady=(0,8))
        bar.columnconfigure(0, weight=1)
        left = ttk.Frame(bar); left.grid(row=0, column=0, sticky="w")
        right = ttk.Frame(bar); right.grid(row=0, column=1, sticky="e")

        # 左（選択操作）
        self.btn_add = ttk.Button(left, text="追加", command=self.on_add)
        self.btn_del = ttk.Button(left, text="削除", command=self.on_del)
        self.btn_add.grid(row=0, column=0, padx=(0,6))
        self.btn_del.grid(row=0, column=1)

        # 右（抽出→出力タイプ→各出力ボタン）
        self.btn_extract = ttk.Button(right, text="抽出（全文解析）", command=self.on_extract)
        self.btn_extract.grid(row=0, column=0, padx=(0,12))

        type_box = ttk.LabelFrame(right, text="出力タイプ")
        type_box.grid(row=0, column=1, padx=(0,12))
        ttk.Radiobutton(type_box, text="XML", value="xml", variable=self.output_type, command=self._refresh_buttons)\
            .grid(row=0, column=0, padx=(6,6))
        ttk.Radiobutton(type_box, text="CSV", value="csv", variable=self.output_type, command=self._refresh_buttons)\
            .grid(row=0, column=1, padx=(0,6))

        # 出力範囲はボタンで分岐
        self.btn_output_full = ttk.Button(right, text="全文出力", command=self.on_output_full)
        self.btn_output_filt = ttk.Button(right, text="フィルタ出力", command=self.on_output_filtered)
        self.btn_output_full.grid(row=0, column=2, padx=(0,6))
        self.btn_output_filt.grid(row=0, column=3)

        # ファイル一覧
        self.listbox = tk.Listbox(self, selectmode=tk.EXTENDED, width=96, height=18, activestyle="dotbox")
        self.listbox.grid(row=1, column=0, sticky="nsew")

        # フィルタ設定（フィルタ出力時に適用）
        pane = ttk.LabelFrame(self, text="フィルタ設定（フィルタ出力時に適用）")
        pane.grid(row=2, column=0, sticky="nsew", pady=(8,0))
        col1 = ttk.Frame(pane); col1.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        col2 = ttk.Frame(pane); col2.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)

        ttk.Label(col1, text="include XPaths（残す・1行1パス）").grid(row=0, column=0, sticky="w")
        self.txt_include = tk.Text(col1, width=58, height=5)
        self.txt_include.grid(row=1, column=0, sticky="nsew")

        ttk.Label(col2, text="exclude XPaths（除く・1行1パス）").grid(row=0, column=0, sticky="w")
        self.txt_exclude = tk.Text(col2, width=58, height=5)
        self.txt_exclude.grid(row=1, column=0, sticky="nsew")

        self.var_rm_empty = tk.BooleanVar(value=True)
        ttk.Checkbutton(pane, text="空要素を削除（子・属性・テキストが空）", variable=self.var_rm_empty)\
            .grid(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=(0,8))

        # 既定値（初回起動時のみ。以降は復元で上書き）
        self._set_text(self.txt_include, ".//Header\n.//ETH-IP\n.//Master\n.//DeviceList")

        # レイアウト
        self.grid(row=0, column=0, sticky="nsew")
        self.master.rowconfigure(0, weight=1); self.master.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1); self.columnconfigure(0, weight=1)
        pane.columnconfigure(0, weight=1); pane.columnconfigure(1, weight=1)
        col1.rowconfigure(1, weight=1); col1.columnconfigure(0, weight=1)
        col2.rowconfigure(1, weight=1); col2.columnconfigure(0, weight=1)

    def _bind_shortcuts(self):
        self.master.bind("<Control-o>", lambda e: self.on_add())
        self.master.bind("<Delete>", lambda e: self.on_del())
        self.master.bind("<Control-e>", lambda e: self.on_extract())
        self.master.bind("<Control-Shift-Return>", lambda e: self.on_output_filtered())
        self.master.bind("<Control-Return>", lambda e: self.on_output_full())

    # ---------------- Actions ----------------
    def on_add(self):
        paths = filedialog.askopenfilenames(title="ファイル選択")
        if paths:
            self._manager.add([Path(p) for p in paths])
            self._sync(); self._refresh_buttons()

    def on_del(self):
        sel = list(self.listbox.curselection())
        self._manager.remove_by_indices(sel)
        self._sync(); self._refresh_buttons()

    def on_extract(self):
        files = self._manager.list()
        if not files:
            messagebox.showwarning("警告", "ファイルが未選択です。"); return
        try:
            parser = APAParser()
            self._last_results = parser.parse_many(files)
            messagebox.showinfo("完了", f"抽出（全文解析）完了: {len(self._last_results)} 件")
            self._refresh_buttons()
        except APAFormatError as e:
            messagebox.showerror("形式エラー", str(e))
        except Exception as e:
            messagebox.showerror("解析エラー", f"{type(e).__name__}: {e}")

    # --- 出力（全文） ---
    def on_output_full(self):
        results = self._get_target_results()
        if not results:
            messagebox.showwarning("警告", "出力対象がありません（抽出未実行、または選択なし）。")
            return
        if self.output_type.get() == "xml":
            self._export_full_xml(results)
        else:
            self._export_full_csv(results)

    # --- 出力（フィルタ） ---
    def on_output_filtered(self):
        results = self._get_target_results()
        if not results:
            messagebox.showwarning("警告", "出力対象がありません（抽出未実行、または選択なし）。")
            return
        include_paths = self._read_lines(self.txt_include)
        exclude_paths = self._read_lines(self.txt_exclude)
        rm_empty = self.var_rm_empty.get()

        if self.output_type.get() == "xml":
            self._export_filtered_xml(results, include_paths, exclude_paths, rm_empty)
        else:
            self._export_filtered_csv(results, include_paths, exclude_paths, rm_empty)

    # ---------------- Export Impl ----------------
    def _export_full_xml(self, results):
        out_dir = Path(__file__).parent / "out" / "xml"
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            for r in results:
                (out_dir / f"{Path(r.source_path).stem}.xml").write_text(r.xml_text, encoding="utf-8")
            messagebox.showinfo("XML出力", f"XML出力完了: {len(results)} 件\n出力先: {out_dir}")
        except Exception as e:
            messagebox.showerror("XML出力エラー", f"{type(e).__name__}: {e}")

    def _export_full_csv(self, results):
        out_dir = Path(__file__).parent / "out"
        out_dir.mkdir(parents=True, exist_ok=True)
        csv_path = out_dir / "summary.csv"
        try:
            with csv_path.open("w", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow(["File","Version","IP","Subnet","Gateway","DeviceCount","LocalSlaveCount"])
                for r in results:
                    s = r.summary
                    w.writerow([
                        str(r.source_path), s.get("Version",""), s.get("IPAddress",""),
                        s.get("Subnet",""),  s.get("Gateway",""),
                        s.get("DeviceCount",""), s.get("LocalSlaveCount","")
                    ])
            messagebox.showinfo("CSV出力", f"CSV出力完了: {len(results)} 件\n{csv_path}")
        except Exception as e:
            messagebox.showerror("CSV出力エラー", f"{type(e).__name__}: {e}")

    def _export_filtered_xml(self, results, include_paths, exclude_paths, rm_empty):
        out_dir = Path(__file__).parent / "out" / "xml_filtered"
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            for r in results:
                xml_text = self._filter_xml(r.xml_text, include_paths, exclude_paths, rm_empty)
                (out_dir / f"{Path(r.source_path).stem}.filtered.xml").write_text(xml_text, encoding="utf-8")
            messagebox.showinfo("XML出力", f"フィルタXML出力完了: {len(results)} 件\n出力先: {out_dir}")
        except Exception as e:
            messagebox.showerror("XML出力エラー", f"{type(e).__name__}: {e}")

    def _export_filtered_csv(self, results, include_paths, exclude_paths, rm_empty):
        """
        フィルタ後XMLをテーブル化。
        先頭列 eip_tag は EPATH(HEX) から復元した EIPタグ名（全EPATH 0x91 を結合）。
        originator_ip は最近傍 Device の szIPAddr（なければ空）。
        """
        out_dir = Path(__file__).parent / "out" / "csv_filtered"
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            for r in results:
                filtered_xml = self._filter_xml(r.xml_text, include_paths, exclude_paths, rm_empty)
                root = ET.fromstring(filtered_xml)

                # 1) 行生成 & 属性キー集合
                rows = []
                attr_keys = set()
                for (xp, el) in self._flatten_xml(root):
                    row_attrs = dict(el.attrib)
                    epath_text = self._get_epath_text(el)
                    eip_tag = self._epath_to_eip_tag(epath_text) if epath_text else ""
                    originator_ip = self._get_originator_ip(root, el)

                    rows.append({
                        "eip_tag": eip_tag,
                        "originator_ip": originator_ip,
                        "tag": el.tag,
                        "path": xp,
                        "text": (el.text or "").strip(),
                        "attrs": row_attrs
                    })
                    attr_keys.update(row_attrs.keys())

                # 2) ヘッダ：固定列 + 属性列
                attr_cols = sorted(attr_keys)
                header = ["eip_tag", "originator_ip", "tag", "path", "text"] + attr_cols

                # 3) 書き出し
                csv_path = out_dir / f"{Path(r.source_path).stem}.filtered.csv"
                with csv_path.open("w", encoding="utf-8", newline="") as f:
                    w = csv.writer(f)
                    w.writerow(header)
                    for row in rows:
                        base = [row["eip_tag"], row["originator_ip"], row["tag"], row["path"], row["text"]]
                        base += [row["attrs"].get(k, "") for k in attr_cols]
                        w.writerow(base)

            messagebox.showinfo("CSV出力", f"フィルタCSV出力完了: {len(results)} 件\n出力先: {out_dir}")
        except Exception as e:
            messagebox.showerror("CSV出力エラー", f"{type(e).__name__}: {e}")

    # ---------------- Helpers ----------------
    def _sync(self):
        self.listbox.delete(0, tk.END)
        for p in self._manager.list():
            self.listbox.insert(tk.END, str(p))

    def _refresh_buttons(self):
        has_files = self._manager.count() > 0
        has_results = len(self._last_results) > 0
        self.btn_del.config(state=("normal" if has_files and self.listbox.size() else "disabled"))
        self.btn_extract.config(state=("normal" if has_files else "disabled"))
        self.btn_output_full.config(state=("normal" if has_results else "disabled"))
        self.btn_output_filt.config(state=("normal" if has_results else "disabled"))

    def _get_target_results(self):
        if not self._last_results:
            return []
        sel = list(self.listbox.curselection())
        if not sel:
            return self._last_results
        selected_paths = {str(self._manager.list()[i]) for i in sel}
        return [r for r in self._last_results if str(r.source_path) in selected_paths]

    @staticmethod
    def _read_lines(widget: tk.Text) -> list[str]:
        return [ln.strip() for ln in widget.get("1.0", "end").splitlines() if ln.strip()]

    @staticmethod
    def _set_text(widget: tk.Text, value: str):
        widget.delete("1.0", "end")
        widget.insert("1.0", value)

    # ---- XML処理（フィルタ/フラット化） ----
    def _filter_xml(self, xml_text: str, include_xpaths: list[str], exclude_xpaths: list[str], remove_empty: bool) -> str:
        root = ET.fromstring(xml_text)
        # include：指定ノード＋先祖のみ残す
        if include_xpaths:
            keep = set()
            for xp in include_xpaths:
                for node in root.findall(xp):
                    cur = node
                    while cur is not None:
                        keep.add(cur)
                        cur = self._get_parent(root, cur)
            self._prune_to_keep(root, keep)
        # exclude：指定ノードを削除
        for xp in exclude_xpaths:
            for node in list(root.findall(xp)):
                self._remove_node(root, node)
        # 空要素削除
        if remove_empty:
            self._remove_empty(root)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")

    def _flatten_xml(self, root: ET.Element):
        """XMLをフラット化し (xpath, element) をyield"""
        def walk(el: ET.Element, path: str):
            cur = f"{path}/{el.tag}" if path else f"/{el.tag}"
            yield (cur, el)
            for ch in list(el):
                yield from walk(ch, cur)
        yield from walk(root, "")

    @staticmethod
    def _get_parent(root: ET.Element, node: ET.Element):
        for el in root.iter():
            for ch in list(el):
                if ch is node:
                    return el
        return None

    def _prune_to_keep(self, root: ET.Element, keep: set[ET.Element]):
        def prune(el: ET.Element):
            for ch in list(el):
                if ch not in keep:
                    el.remove(ch)
                else:
                    prune(ch)
        prune(root)

    def _remove_node(self, root: ET.Element, target: ET.Element):
        parent = self._get_parent(root, target)
        if parent is not None:
            parent.remove(target)

    def _remove_empty(self, root: ET.Element):
        def is_empty(el: ET.Element) -> bool:
            return (len(list(el)) == 0) and (len(el.attrib) == 0) and ((el.text or "").strip() == "")
        changed = True
        while changed:
            changed = False
            for el in list(root.iter()):
                if el is root:
                    continue
                if is_empty(el):
                    parent = self._get_parent(root, el)
                    if parent is not None:
                        parent.remove(el)
                        changed = True

    # ---------------- EPATH → EIPタグ名 復元 ----------------
    def _hex_to_bytes(self, s: str) -> bytes:
        """'91 06 61 6E...' / '9106616E...' などのHEX表記をbytesへ。非HEXは無視。"""
        import re, binascii
        s = s.strip()
        s = s.replace("0x", "").replace("0X", "")
        s = re.sub(r"[^0-9A-Fa-f]", "", s)  # 非HEX削除
        if len(s) % 2 == 1:
            s = s[:-1]  # 余りは切り捨て
        try:
            return binascii.unhexlify(s)
        except Exception:
            return b""

    def _epath_to_eip_tag(self, epath_raw: str) -> str:
        """
        EPATH（HEX）からEIPタグ名を復元。
        対応：ANSI拡張シンボリックセグメント 0x91 + len + ascii + (偶数境界までパディング)。
        0x91 以外のバイトは 1 バイトずつスキップして全体を走査し、出現した 0x91 を全て連結する。
        """
        data = self._hex_to_bytes(epath_raw)
        if not data:
            return ""

        i = 0
        n = len(data)
        parts: list[str] = []

        while i < n:
            seg = data[i]
            if seg == 0x91:  # ANSI extended symbolic segment
                if i + 1 >= n:
                    break
                length = data[i + 1]
                start = i + 2
                end = start + length
                if end > n:
                    break
                name_bytes = data[start:end]
                # 偶数境界パディング
                i = end + (1 if (length % 2) == 1 and end < n else 0)
                try:
                    name = name_bytes.decode("ascii", errors="ignore")
                except Exception:
                    name = ""
                if name:
                    parts.append(name)
                continue
            # 0x91 以外は 1 バイトだけ進めてスキャン継続
            i += 1

        return ".".join(parts) if parts else ""

    def _get_epath_text(self, el: ET.Element) -> str:
        """要素の属性 or 子要素からEPATH（HEX表記）文字列を取得。見つからなければ空文字。"""
        # 属性優先（大文字小文字ゆらぎ対応）
        for k in el.attrib.keys():
            if k.lower() in ("epath", "szepath", "e-path", "path_epath"):
                txt = str(el.attrib.get(k, "")).strip()
                if txt:
                    return txt
        # 子要素 EPATH
        child = el.find("EPATH")
        if child is not None and (child.text or "").strip():
            return child.text.strip()
        return ""

    # -------- Originator IP（最近傍 Device の szIPAddr） --------
    def _get_originator_ip(self, root: ET.Element, el: ET.Element) -> str:
        """
        el から親方向に辿り、最初に見つかった Device 要素の szIPAddr を返す。
        大文字/小文字ゆらぎ szIpAddr も許容。無ければ空文字。
        """
        cur = el
        while cur is not None:
            if cur.tag.lower() == "device":
                val = cur.attrib.get("szIPAddr") or cur.attrib.get("szIpAddr")
                if val and str(val).strip():
                    return str(val).strip()
            cur = self._get_parent(root, cur)
        return ""

    # ---------------- 設定保存/復元 ----------------
    def _on_close(self):
        self._save_prefs()
        self.master.destroy()

    def _save_prefs(self):
        prefs = {
            "include": self._read_lines(self.txt_include),
            "exclude": self._read_lines(self.txt_exclude),
            "remove_empty": bool(self.var_rm_empty.get()),
            "output_type": self.output_type.get(),
        }
        try:
            PREF_PATH.write_text(json.dumps(prefs, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass  # 保存失敗は致命ではない

    def _load_prefs(self):
        if not PREF_PATH.exists():
            return
        try:
            data = json.loads(PREF_PATH.read_text(encoding="utf-8"))
            self._set_text(self.txt_include, "\n".join(data.get("include", [])))
            self._set_text(self.txt_exclude, "\n".join(data.get("exclude", [])))
            self.var_rm_empty.set(bool(data.get("remove_empty", True)))
            self.output_type.set(data.get("output_type", "xml"))
        except Exception:
            pass


def main():
    root = tk.Tk()
    app = FileSelectApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
