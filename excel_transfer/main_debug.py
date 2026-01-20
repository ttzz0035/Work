# excel_transfer/main.py
import os
import sys
import traceback

from utils.log import init_logger
from utils.configs import load_context
from ui.app import ExcelApp


class _EmptyContext:
    """
    ライセンス・INI・設定が一切なくても
    アプリを起動させるための最小 Context
    """
    def __init__(self):
        self.tabs_enabled = None
        self.license_manager = None


def main():
    # ----------------------------------------
    # base dir（exe / script 両対応）
    # ----------------------------------------
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # ----------------------------------------
    # logger
    # ----------------------------------------
    logger = init_logger(base_dir)
    logger.info(f"[MAIN] base_dir={base_dir}")

    # ----------------------------------------
    # context（license / ini が無くても起動）
    # ----------------------------------------
    try:
        ctx = load_context(base_dir, logger)
        logger.info("[MAIN] context loaded")
    except Exception as ex:
        logger.warning(
            "[MAIN] load_context failed -> start without license/config\n"
            + "".join(traceback.format_exception_only(type(ex), ex)).strip()
        )
        ctx = _EmptyContext()

    # ----------------------------------------
    # App 起動（全タブ生成）
    # ----------------------------------------
    app = ExcelApp(ctx, logger)

    # ----------------------------------------
    # INI によるタブ削除（存在する場合のみ）
    # ----------------------------------------
    if getattr(ctx, "tabs_enabled", None):
        tab_count = app.nb.index("end")
        logger.info(f"[MAIN] tab count={tab_count}")

        for i in reversed(range(tab_count)):
            try:
                tab_key = f"tab{i + 1}"
                enabled = ctx.tabs_enabled.get(tab_key, True)
                if not enabled:
                    app.nb.forget(i)
                    logger.info(f"[MAIN] removed tab idx={i} key={tab_key}")
            except Exception as ex:
                logger.warning(f"[MAIN] tab remove failed idx={i} err={ex}")

        if app.nb.index("end") == 1:
            app.nb.select(0)
    else:
        logger.info("[MAIN] tab filter not defined (skip)")

    # ----------------------------------------
    # Run
    # ----------------------------------------
    logger.info("[MAIN] app start")
    app.run()


if __name__ == "__main__":
    main()
