
import queue
import threading
import tkinter as tk
from tkinter import ttk

from ui_log import UILamp, UILayout, UIProgress, UIStatus, parse_ui


LAMP_COLORS = {
    "green": "#22C55E",
    "yellow": "#F59E0B",
    "red": "#EF4444",
}
LAMP_OFF = "#94A3B8"


class NodeWindow:
    """Node別UIログからカードを動的生成・更新するWindow。"""

    def __init__(self, parent, api, name):
        self.api, self.name = api, name
        self.window = tk.Toplevel(parent)
        self.window.title(f"{name} - Dashboard")
        self.window.geometry("700x470")
        self.window.minsize(400, 260)

        self.closed = threading.Event()
        self.events = queue.Queue()
        self.cards = {}
        self.columns = 1
        self.previous_logs = []

        self.build_ui()
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        threading.Thread(target=self.watch, daemon=True).start()
        self.window.after(100, self.poll)

    def build_ui(self):
        """スクロール可能なDashboard領域を生成する。"""
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(1, weight=1)

        header = ttk.Frame(self.window, padding=12)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text=self.name, font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        self.status = tk.StringVar(value="Connecting...")
        ttk.Label(header, textvariable=self.status).grid(row=0, column=1)

        container = ttk.Frame(self.window)
        container.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0, background="#F1F5F9")
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.content = tk.Frame(canvas, background="#F1F5F9")
        content_id = canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(content_id, width=e.width))
        self.set_columns(1)

    def set_columns(self, columns):
        """既存カードとの整合性を検証して列数を更新する。"""
        for card_id, card in self.cards.items():
            if card["column"] + card["span"] > columns:
                raise ValueError(f"Card exceeds new columns: {card_id}")

        for column in range(max(self.columns, columns)):
            self.content.columnconfigure(column, weight=0, uniform="")
        for column in range(columns):
            self.content.columnconfigure(column, weight=1, uniform="cards")
        self.columns = columns

    def watch(self):
        """別スレッドでProcess Composeのログを監視する。"""
        while not self.closed.is_set():
            try:
                logs = self.api.process_logs(self.name)
                if not isinstance(logs, list):
                    raise ValueError("Invalid logs response")

                # 前回取得したログとの重複を除外する。
                overlap = min(len(self.previous_logs), len(logs))
                while overlap and self.previous_logs[-overlap:] != logs[:overlap]:
                    overlap -= 1

                for line in logs[overlap:]:
                    try:
                        message = parse_ui(line)
                        if message is not None:
                            self.events.put(("ui", message))
                    except ValueError as exc:
                        self.events.put(("error", str(exc)))

                self.previous_logs = logs
                self.events.put(("connection", "Live"))
            except Exception as exc:
                self.events.put(("connection", f"Disconnected: {exc}"))

            self.closed.wait(1)

    def create_card(self, message):
        """新しいUI部品に対応するカードを生成する。"""
        card_id = message.id
        row = message.row if message.row is not None else len(self.cards)
        column = message.column if message.column is not None else 0
        span = message.columnspan if message.columnspan is not None else 1

        if column + span > self.columns:
            raise ValueError(f"Card exceeds grid columns: {card_id}")

        # 同じ行の既存カードと表示領域が重ならないことを確認する。
        for existing in self.cards.values():
            if existing["row"] == row and (
                column < existing["column"] + existing["span"]
                and existing["column"] < column + span
            ):
                raise ValueError(f"Overlapping UI cards: {card_id}")

        frame = tk.Frame(self.content, bg="white", padx=16, pady=14,
                         highlightbackground="#CBD5E1", highlightthickness=1)
        frame.grid(row=row, column=column, columnspan=span, sticky="nsew", padx=6, pady=6)
        frame.columnconfigure(0, weight=1)

        title = tk.Label(frame, text=message.label or card_id, anchor="w",
                         bg="white", fg="#64748B", font=("Segoe UI", 10))
        title.grid(row=0, column=0, sticky="ew")

        # typeに応じて適切なコントロールを生成する。
        lamp = None
        lamp_text = None
        progress = None

        if isinstance(message, UILamp):
            # ランプと状態文字列を同じ行に配置する。
            lamp_frame = tk.Frame(frame, bg="white")
            lamp_frame.grid(row=1, column=0, sticky="w", pady=(2, 0))

            lamp = tk.Label(lamp_frame, text="●", bg="white", fg=LAMP_OFF,
                            font=("Segoe UI", 32))
            lamp.grid(row=0, column=0, sticky="w")

            # ランプ横の状態文字列はUIログのtextで更新する。
            lamp_text = tk.Label(lamp_frame, text="", bg="white", fg="#0F172A",
                                 font=("Segoe UI", 17, "bold"))
            lamp_text.grid(row=0, column=1, sticky="w", padx=(10, 0))
            value = None
        else:
            value = tk.Label(frame, anchor="w", bg="white", fg="#0F172A",
                             font=("Segoe UI", 17, "bold"))
            value.grid(row=1, column=0, sticky="ew", pady=(6, 0))
            if isinstance(message, UIProgress):
                progress = ttk.Progressbar(frame, maximum=100)
                progress.grid(row=2, column=0, sticky="ew", pady=(10, 0))

        self.cards[card_id] = {
            "title": title, "value": value, "lamp": lamp, "lamp_text": lamp_text,
            "progress": progress, "type": message.type, "row": row,
            "column": column, "span": span,
            "unit": getattr(message, "unit", None) or "",
            "color": getattr(message, "color", None) or "green", "on": False,
        }

    def apply(self, message):
        """Layoutまたはカードの表示内容を反映する。"""
        if isinstance(message, UILayout):
            self.set_columns(message.columns)
            return

        if message.id not in self.cards:
            self.create_card(message)

        card = self.cards[message.id]
        if card["type"] != message.type:
            raise ValueError(f"Cannot change UI type: {message.id}")

        # 作成済みカードのGrid配置は変更させない。
        for key, original in (("row", card["row"]), ("column", card["column"]),
                              ("columnspan", card["span"])):
            new_value = getattr(message, key)
            if new_value is not None and new_value != original:
                raise ValueError(f"Cannot change {key}: {message.id}")

        if message.label is not None:
            card["title"].configure(text=message.label)

        if isinstance(message, UILamp):
            # ランプの色とON/OFF状態を更新する。
            if message.color is not None:
                card["color"] = message.color
            if message.value is not None:
                card["on"] = message.value
            card["lamp"].configure(fg=LAMP_COLORS[card["color"]] if card["on"] else LAMP_OFF)

            # textが指定された場合、ランプ横の文字を更新する。
            if message.text is not None:
                card["lamp_text"].configure(text=message.text)
            return

        if getattr(message, "unit", None) is not None:
            card["unit"] = message.unit
        if message.value is None:
            return

        if isinstance(message, UIProgress):
            card["progress"]["value"] = message.value
            card["value"].configure(text=f"{message.value:g}%")
        else:
            card["value"].configure(text=f"{message.value} {card['unit']}".strip())
            if isinstance(message, UIStatus):
                healthy = message.value.lower() in ("connected", "running", "ok")
                card["value"].configure(fg="#16A34A" if healthy else "#DC2626")

    def poll(self):
        """Tkinterのメインスレッドで画面を更新する。"""
        if self.closed.is_set():
            return
        while not self.events.empty():
            kind, value = self.events.get_nowait()
            if kind == "connection":
                self.status.set(value)
            elif kind == "error":
                self.status.set(f"UI Error: {value}")
            else:
                try:
                    self.apply(value)
                except (ValueError, TypeError, tk.TclError) as exc:
                    self.status.set(f"UI Error: {exc}")
        self.window.after(100, self.poll)

    def close(self):
        """ログ監視を終了する。Node自体は停止しない。"""
        self.closed.set()
        self.window.destroy()
