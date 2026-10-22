
import queue
import threading
import tkinter as tk
from tkinter import ttk

from ui_log import UILamp, UILayout, UIProgress, UIValue, parse_ui


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
        canvas = tk.Canvas(container, highlightthickness=0, bg="#F1F5F9")
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)
        self.content = tk.Frame(canvas, bg="#F1F5F9")
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
        """初回ログではカードと通常のコントロールを生成する。"""
        card_id = message.id
        row = message.row if message.row is not None else len(self.cards)
        column = message.column if message.column is not None else 0
        span = message.columnspan if message.columnspan is not None else 1
        if column + span > self.columns:
            raise ValueError(f"Card exceeds grid columns: {card_id}")
        # 同じ行の既存カードと表示領域が重ならないことを確認する。
        for existing in self.cards.values():
            if existing["row"] == row and column < existing["column"] + existing["span"] and existing["column"] < column + span:
                raise ValueError(f"Overlapping UI cards: {card_id}")

        frame = tk.Frame(self.content, bg="white", padx=16, pady=14, highlightbackground="#CBD5E1", highlightthickness=1)
        frame.grid(row=row, column=column, columnspan=span, sticky="nsew", padx=6, pady=6)
        frame.columnconfigure(0, weight=1)
        title = tk.Label(frame, text=message.label or card_id, anchor="w", bg="white", fg="#64748B", font=("Segoe UI", 10))
        title.grid(row=0, column=0, sticky="ew")

        # typeに応じて適切なコントロールを生成する。
        lamp = lamp_text = progress = None
        if isinstance(message, UILamp):
            # ランプと状態文字列を同じ行に配置する。
            lamp_frame = tk.Frame(frame, bg="white")
            lamp_frame.grid(row=1, column=0, sticky="w", pady=(2, 0))
            lamp = tk.Label(lamp_frame, text="●", bg="white", fg="#94A3B8", font=("Segoe UI", 32))
            lamp.grid(row=0, column=0, sticky="w")
            # ランプ横の状態文字列はUIログのtextで更新する。
            lamp_text = tk.Label(lamp_frame, text="", bg="white", fg="#0F172A", font=("Segoe UI", 17, "bold"))
            lamp_text.grid(row=0, column=1, sticky="w", padx=(10, 0))
            value = None
        else:
            value = tk.Label(frame, anchor="w", bg="white", fg="#0F172A", font=("Segoe UI", 17, "bold"))
            value.grid(row=1, column=0, sticky="ew", pady=(6, 0))
            if isinstance(message, UIProgress):
                # 初回は通常Progressbarを生成する。
                progress = ttk.Progressbar(frame, maximum=100)
                progress.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.cards[card_id] = {
            "frame": frame, "title": title, "value": value,
            "lamp": lamp, "lamp_text": lamp_text, "progress": progress,
            "gauge": None, "ticks": None, "type": message.type,
            "row": row, "column": column, "span": span,
            "unit": getattr(message, "unit", None) or "",
            "color": getattr(message, "color", None) or "#94A3B8",
            "reference": None, "range": None, "current": None,
        }

    def enable_reference(self, card):
        """更新ログで基準値が設定された時点で中央基準型へ切り替える。"""
        card["progress"].destroy()
        card["progress"] = None
        frame = card["frame"]
        gauge = tk.Canvas(frame, height=22, bg="white", highlightthickness=0)
        gauge.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ticks = tk.Label(frame, bg="white", fg="#64748B", font=("Segoe UI", 9))
        ticks.grid(row=3, column=0, sticky="ew")
        card["gauge"], card["ticks"] = gauge, ticks
        gauge.bind("<Configure>", lambda event, c=card: self.draw_reference(c))

    def draw_reference(self, card):
        """基準からの偏差を左右に描画し、目盛りを更新する。"""
        canvas = card["gauge"]
        width = canvas.winfo_width()
        center = width / 2
        reference = card["reference"]
        lower, upper = card["range"]

        canvas.delete("all")
        canvas.create_rectangle(0, 5, width, 17, fill="#E2E8F0", outline="")

        current = card["current"]
        if current is not None:
            delta = current - reference

            # マイナス側とプラス側で異なる表示範囲を使用する。
            if delta < 0:
                endpoint = center + max(delta, lower) / abs(lower) * center
                color = "#16A34A"
            else:
                endpoint = center + min(delta, upper) / upper * center
                color = "#EF4444"

            if delta != 0:
                canvas.create_rectangle(min(center, endpoint), 5, max(center, endpoint), 17, fill=color, outline="")

        canvas.create_line(center, 0, center, 22, fill="#0F172A", width=2)

        unit = card["unit"]
        card["ticks"].configure(
            text=f"{reference+lower:g} {unit}     |     {reference:g} {unit} (基準)     |     {reference+upper:g} {unit}"
        )
        
    def update_progress(self, card, message):
        """通常Progressbarまたは中央基準型の値を更新する。"""
        reference, spread = message.reference, message.range
        # 通常型から中央基準型へ切り替える場合は両方の設定が必須。
        if card["gauge"] is None and (reference is not None or spread is not None):
            if reference is None or spread is None:
                raise ValueError("reference and range must be specified together")
            card["reference"], card["range"] = reference, spread
            self.enable_reference(card)
        if card["gauge"] is None:
            if message.value is not None:
                if message.value > 100:
                    raise ValueError("Normal progress exceeds 100")
                card["current"] = message.value
                card["progress"]["value"] = message.value
                card["value"].configure(text=f"{message.value:g}%")
            return

        # 中央基準型では基準値・範囲・単位も更新できる。
        if reference is not None:
            card["reference"] = reference
        if spread is not None:
            card["range"] = spread
        if message.unit is not None:
            card["unit"] = message.unit
        if message.value is not None:
            card["current"] = message.value
        current = card["current"]
        if current is not None:
            difference = current - card["reference"]
            card["value"].configure(text=f"{current:g} {card['unit']}   ({difference:+g} {card['unit']})")
        self.draw_reference(card)

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
        for key, original in (("row", card["row"]), ("column", card["column"]), ("columnspan", card["span"])):
            new_value = getattr(message, key)
            if new_value is not None and new_value != original:
                raise ValueError(f"Cannot change {key}: {message.id}")
        if message.label is not None:
            card["title"].configure(text=message.label)

        if isinstance(message, UILamp):
            # 指定されたカラーコードでランプを更新する。
            if message.color is not None:
                card["color"] = message.color
            card["lamp"].configure(fg=card["color"])
            # textが指定された場合、ランプ横の文字を更新する。
            if message.text is not None:
                card["lamp_text"].configure(text=message.text)
        elif isinstance(message, UIValue):
            # 通常値と状態値を同じLabelで表示する。
            if message.unit is not None:
                card["unit"] = message.unit
            if message.color is not None:
                card["value"].configure(fg=message.color)
            if message.value is not None:
                card["value"].configure(text=f"{message.value} {card['unit']}".strip())
        elif isinstance(message, UIProgress):
            self.update_progress(card, message)

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
