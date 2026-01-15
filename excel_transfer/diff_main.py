# excel_transfer/diff_main.py
import os
import sys

from utils.log import init_logger
from utils.configs import load_context
from ui.app import ExcelApp


def main():
    # ----------------------------------------
    # base dir 解決
    # ----------------------------------------
    base_dir = (
        os.path.dirname(sys.executable)
        if getattr(sys, "frozen", False)
        else os.path.dirname(os.path.abspath(__file__))
    )

    # ----------------------------------------
    # logger / context
    # ----------------------------------------
    logger = init_logger(base_dir)
    ctx = load_context(base_dir, logger)

    # ----------------------------------------
    # 通常 ExcelApp を生成
    # ----------------------------------------
    app = ExcelApp(ctx, logger)

    # ----------------------------------------
    # Diff 以外のタブを削除
    # ----------------------------------------
    diff_label = ctx.labels.get("section_diff", "Diff")

    # index は動くので後ろから消す
    for i in reversed(range(app.nb.index("end"))):
        try:
            tab_text = app.nb.tab(i, "text")
            if tab_text != diff_label:
                app.nb.forget(i)
                logger.info(f"[DIFF_MAIN] removed tab idx={i} text={tab_text}")
        except Exception as ex:
            logger.warning(f"[DIFF_MAIN] tab remove failed idx={i} err={ex}")

    # 念のため Diff タブを選択
    try:
        if app.nb.index("end") == 1:
            app.nb.select(0)
    except Exception:
        pass

    # ----------------------------------------
    # 起動
    # ----------------------------------------
    app.run()


if __name__ == "__main__":
    main()
