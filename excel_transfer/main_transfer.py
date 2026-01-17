# excel_transfer/transfer_main.py
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

    logger.info(f"[TRANSFER_MAIN] base_dir={base_dir}")

    # ----------------------------------------
    # App 起動
    # ----------------------------------------
    app = ExcelApp(ctx, logger)

    # Transfer タブ以外を削除
    transfer_label = ctx.labels.get("section_transfer", "Transfer")
    for i in reversed(range(app.nb.index("end"))):
        try:
            if app.nb.tab(i, "text") != transfer_label:
                app.nb.forget(i)
                logger.info(f"[TRANSFER_MAIN] removed tab idx={i}")
        except Exception as ex:
            logger.warning(f"[TRANSFER_MAIN] tab remove failed idx={i} err={ex}")

    if app.nb.index("end") == 1:
        app.nb.select(0)

    app.run()


if __name__ == "__main__":
    main()
