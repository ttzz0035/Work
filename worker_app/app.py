# SPDX-License-Identifier: MIT
import flet as ft
import logging, queue, time, configparser
from datetime import datetime, date, timedelta
from dataclasses import dataclass
from pathlib import Path
from logging.handlers import RotatingFileHandler

from db_ops import get_items
from worker import run_worker
from consts import *  # UI文字列・寸法

LOG_DIR = Path("logs"); LOG_DIR.mkdir(exist_ok=True)
CONFIG_FILE = Path("config.ini")
CONFIG_SECTION = "APP"

# ------------------------------------------------------------
# ロガー設定
# ------------------------------------------------------------
class UILogHandler(logging.Handler):
    def __init__(self, q: "queue.Queue[str]"):
        super().__init__()
        self.q = q
        self.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
    def emit(self, record):
        try:
            self.q.put(self.format(record))
        except Exception:
            pass


def setup_logger(ui_queue: "queue.Queue[str]"):
    lg = logging.getLogger()
    lg.setLevel(logging.INFO)
    for h in list(lg.handlers): lg.removeHandler(h)
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh = RotatingFileHandler(LOG_DIR / "app.log", maxBytes=1_000_000, backupCount=5, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(); ch.setFormatter(fmt)
    lg.addHandler(fh); lg.addHandler(ch); lg.addHandler(UILogHandler(ui_queue))
    return lg


# ------------------------------------------------------------
# 設定クラス
# ------------------------------------------------------------
@dataclass
class AppConfig:
    selected_item_id: int | None = None
    exec_mode: str = MODE_REGISTER


# ------------------------------------------------------------
# メインコントローラ
# ------------------------------------------------------------
class AppController:
    def __init__(self, page: ft.Page):
        self.page = page
        self.ui_log_q: queue.Queue[str] = queue.Queue()
        setup_logger(self.ui_log_q)

        # ---------- 設定読込 ----------
        self.config = configparser.ConfigParser()
        self._load_config()

        self.cfg = AppConfig()
        self.cfg.selected_item_id = self.config.getint(CONFIG_SECTION, "last_job_id", fallback=None)
        self.cfg.exec_mode = self.config.get(CONFIG_SECTION, "exec_mode", fallback=MODE_REGISTER)

        self.runtime = dict(
            running=False, ticks=0, started_at=None, last_tick_at=None,
            item_id=None, start=None, end=None, mode=self.cfg.exec_mode
        )

        # ---------- UI参照 ----------
        self.tf_logs = None
        self.lbl_started = None
        self.lbl_elapsed = None
        self.lbl_ticks = None
        self.status_badge = None
        self.mode_group = None

        # ---------- Window設定 ----------
        self._fix_window_width(CARD_WIDTH)

        # ---------- ページ ----------
        self.page.title = APP_TITLE
        self.page.window_resizable = True
        self.page.vertical_alignment = ft.MainAxisAlignment.START
        self.page.padding = 0
        self.page.spacing = 0
        self.page.scroll = "always"

        # ---------- ボタンスタイル ----------
        self.button_style_primary = ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE,
            shape=ft.RoundedRectangleBorder(radius=6),
        )
        self.button_style_secondary = ft.ButtonStyle(
            color=ft.Colors.BLUE,
            bgcolor=ft.Colors.BLUE_50,
            shape=ft.RoundedRectangleBorder(radius=6),
        )

        self.page.on_route_change = self.route_change
        self.page.go("/")

    # --------------------------------------------------------
    # 設定ファイル操作
    # --------------------------------------------------------
    def _load_config(self):
        if CONFIG_FILE.exists():
            self.config.read(CONFIG_FILE, encoding="utf-8")
        if CONFIG_SECTION not in self.config:
            self.config[CONFIG_SECTION] = {}
            self._save_config()

    def _save_config(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            self.config.write(f)

    def _save_job_id(self, job_id: int | None):
        if job_id is None: return
        self.config[CONFIG_SECTION]["last_job_id"] = str(job_id)
        self._save_config()
        logging.info(f"[CFG] ジョブID保存: {job_id}")

    def _save_exec_mode(self, mode: str):
        self.config[CONFIG_SECTION]["exec_mode"] = mode
        self._save_config()
        logging.info(f"[CFG] 実行モード保存: {mode}")

    # --------------------------------------------------------
    def _fix_window_width(self, w: int):
        win = self.page.window
        try:
            win.maximized = False
            win.resizable = True
            win.min_width = w
            win.max_width = w
            win.width = w
            time.sleep(0.05)
            win.fit_content = True
            self.page.update()
            logging.info(f"[APP] Window幅固定 {w}px")
        except Exception as e:
            logging.warning(f"[APP] Window固定に失敗: {e}")
        self.card_width = w
        self.field_width = w - IN_PADDING * 2

    # --------------------------------------------------------
    def append_logs_from_queue(self):
        if not hasattr(self, "tf_logs") or self.tf_logs is None:
            return
        while not self.ui_log_q.empty():
            self.tf_logs.value += self.ui_log_q.get_nowait() + "\n"
        self.page.update()

    def update_status(self):
        if not self.status_badge:
            return
        rt = self.runtime
        if rt["running"]:
            self.status_badge.bgcolor = ft.Colors.GREEN
            self.status_badge.content = ft.Text("RUNNING", color=ft.Colors.WHITE, weight="bold")
            if rt["started_at"]:
                self.lbl_started.value = rt["started_at"].strftime("%Y-%m-%d %H:%M:%S")
                sec = int((datetime.now() - rt["started_at"]).total_seconds())
                h, m, s = sec // 3600, (sec % 3600)//60, sec % 60
                self.lbl_elapsed.value = f"{h:02d}:{m:02d}:{s:02d}"
            self.lbl_ticks.value = str(rt["ticks"])
        else:
            self.status_badge.bgcolor = ft.Colors.GREY
            self.status_badge.content = ft.Text("STOPPED", color=ft.Colors.WHITE, weight="bold")
        self.page.update()

    # --------------------------------------------------------
    # 設定画面
    # --------------------------------------------------------
    def build_settings(self) -> ft.View:
        cw, fw = self.card_width, self.field_width
        items = get_items()

        dd_job = ft.Dropdown(
            label=LBL_JOB_SETTING,
            options=[ft.dropdown.Option(str(i), nm) for i, nm in items],
            width=fw,
            value=str(self.cfg.selected_item_id) if self.cfg.selected_item_id else None,
            on_change=lambda e: self._on_job_change(dd_job),
        )

        tf_start = ft.TextField(label=LBL_START_DATE, value=self.today_str(), width=(fw // 2 - 4))
        tf_end   = ft.TextField(label=LBL_END_DATE, value=self.today_str(), width=(fw // 2 - 4))

        # 実行モード
        self.mode_group = ft.RadioGroup(
            value=self.cfg.exec_mode,
            content=ft.Row([
                ft.Radio(value=MODE_REGISTER, label=LBL_MODE_REGISTER),
                ft.Radio(value=MODE_VERIFY, label=LBL_MODE_VERIFY),
            ], alignment=ft.MainAxisAlignment.START, spacing=20),
            on_change=lambda e: self._on_mode_change(),
        )

        def set_this_month(_):
            d = date.today()
            first = d.replace(day=1)
            next_month = (first.replace(day=28) + timedelta(days=4)).replace(day=1)
            last = next_month - timedelta(days=1)
            tf_start.value, tf_end.value = first.strftime("%Y-%m-%d"), last.strftime("%Y-%m-%d")
            self.page.update()

        def set_today(_):
            s = self.today_str()
            tf_start.value = tf_end.value = s
            self.page.update()

        def start_run(_):
            self.runtime.update(dict(
                running=True, ticks=0,
                started_at=datetime.now(), last_tick_at=None,
                item_id=self.cfg.selected_item_id,
                start=tf_start.value, end=tf_end.value,
                mode=self.mode_group.value,
            ))
            logging.info(f"[RUN] 実行開始: {self.runtime}")
            self.page.go("/run")
            run_worker(self.runtime, self.append_logs_from_queue, self.update_status)
            self.page.go("/")

        job_card = ft.Card(
            content=ft.Container(
                width=cw, padding=IN_PADDING,
                content=ft.Column([ft.Text(LBL_JOB_SETTING, size=16, weight="bold"), dd_job], spacing=10)
            )
        )

        date_card = ft.Card(
            content=ft.Container(
                width=cw, padding=IN_PADDING,
                content=ft.Column([
                    ft.Text(LBL_PERIOD_SETTING, size=16, weight="bold"),
                    ft.Row([tf_start, tf_end], alignment=ft.MainAxisAlignment.SPACE_BETWEEN, width=fw),
                    ft.Row([
                        ft.ElevatedButton(BTN_SET_THIS_MONTH, on_click=set_this_month, style=self.button_style_secondary),
                        ft.ElevatedButton(BTN_SET_TODAY, on_click=set_today, style=self.button_style_secondary),
                    ], alignment=ft.MainAxisAlignment.START, width=fw),
                ], spacing=10)
            )
        )

        control_card = ft.Card(
            content=ft.Container(
                width=cw, padding=IN_PADDING,
                content=ft.Row([
                    self.mode_group,
                    ft.ElevatedButton(BTN_RUN, on_click=start_run, width=fw // 2, style=self.button_style_primary),
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            )
        )

        layout = ft.ListView(expand=True, spacing=10, padding=0, controls=[job_card, date_card, control_card])
        return ft.View("/", [ft.Container(width=cw, content=layout)])

    # --------------------------------------------------------
    # 実行画面
    # --------------------------------------------------------
    def build_run(self) -> ft.View:
        cw, fw = self.card_width, self.field_width

        self.status_badge = ft.Container(
            content=ft.Text("STOPPED", color=ft.Colors.WHITE, weight="bold"),
            bgcolor=ft.Colors.GREY,
            padding=ft.padding.symmetric(8, 4),
            border_radius=999,
        )
        self.lbl_started = ft.Text("—")
        self.lbl_elapsed = ft.Text("—")
        self.lbl_ticks = ft.Text("0")
        self.tf_logs = ft.TextField(
            label="", multiline=True, read_only=True,
            min_lines=12, max_lines=12, width=fw, expand=False
        )

        def stop_run(_):
            self.stop()
            self.page.go("/")

        status_card = ft.Card(
            content=ft.Container(
                width=cw, padding=IN_PADDING,
                content=ft.Column([
                    ft.Row([ft.Text(LBL_STATUS, size=16, weight="bold"), self.status_badge],
                           alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    ft.Row([ft.Text(LBL_STARTED), self.lbl_started]),
                    ft.Row([ft.Text(LBL_ELAPSED), self.lbl_elapsed]),
                    ft.Row([ft.Text(LBL_TICKS), self.lbl_ticks]),
                    ft.Row([ft.ElevatedButton(BTN_STOP_AND_RETURN, on_click=stop_run, style=self.button_style_secondary)]),
                ], spacing=6)
            )
        )

        log_card = ft.Card(
            content=ft.Container(
                width=cw, padding=IN_PADDING,
                content=ft.Column([
                    ft.Text(LBL_LOG_TITLE, size=16, weight="bold"),
                    self.tf_logs,
                    ft.Row([
                        ft.TextButton(BTN_CLEAR_LOG,
                                      on_click=lambda e: (setattr(self.tf_logs, "value", ""), self.page.update()))
                    ], alignment=ft.MainAxisAlignment.END)
                ], spacing=10)
            )
        )

        layout = ft.ListView(expand=True, spacing=10, padding=0,
                             controls=[status_card, log_card])
        return ft.View("/run", [ft.Container(width=cw, content=layout)])

    # --------------------------------------------------------
    # ハンドラ・ルート
    # --------------------------------------------------------
    def _on_job_change(self, dd):
        self.cfg.selected_item_id = int(dd.value) if dd.value else None
        self._save_job_id(self.cfg.selected_item_id)
        logging.info(f"[APP] ジョブ選択: {self.cfg.selected_item_id}")
        self.append_logs_from_queue()

    def _on_mode_change(self):
        self.cfg.exec_mode = self.mode_group.value
        self._save_exec_mode(self.cfg.exec_mode)
        logging.info(f"[APP] 実行モード変更: {self.cfg.exec_mode}")

    def stop(self):
        self.runtime["running"] = False
        logging.info("[RUN] 停止要求")
        self.update_status()

    def route_change(self, e: ft.RouteChangeEvent):
        self._fix_window_width(CARD_WIDTH)
        self.page.views.clear()
        if self.page.route == "/run":
            self.page.views.append(self.build_run())
        else:
            self.page.views.append(self.build_settings())
        self.page.update()

    @staticmethod
    def today_str() -> str:
        return date.today().strftime("%Y-%m-%d")


# ------------------------------------------------------------
def main(page: ft.Page):
    AppController(page)


ft.app(target=main)
