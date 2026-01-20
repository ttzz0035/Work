# excel_transfer/main.py
import os
import sys

from utils.log import init_logger
from utils.configs import load_context
from ui.app import ExcelApp


def main():
    # ----------------------------------------
    # base dir（exe / script 両対応）
    # ----------------------------------------
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # ----------------------------------------
    # logger / context
    # ----------------------------------------
    logger = init_logger(base_dir)
    ctx = load_context(base_dir, logger)

    logger.info(f"[MAIN] base_dir={base_dir}")

    # ----------------------------------------
    # App 起動（全タブ生成）
    # ----------------------------------------
    app = ExcelApp(ctx, logger)

    # ----------------------------------------
    # INI による後削除（先頭からの index のみ）
    # ----------------------------------------
    if ctx.tabs_enabled:
        tab_count = app.nb.index("end")
        logger.info(f"[MAIN] tab count={tab_count}")

        for i in reversed(range(tab_count)):
            try:
                tab_key = f"tab{i + 1}"  # idx=0 -> tab1

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
    app.run()


if __name__ == "__main__":
    main()
