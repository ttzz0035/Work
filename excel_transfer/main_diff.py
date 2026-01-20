# excel_transfer/diff_main.py
import os
import sys
import configparser

from utils.log import init_logger
from utils.configs import load_context
from ui.app import ExcelApp


def _resolve_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _load_app_version(base_dir: str) -> str:
    ini_path = os.path.join(base_dir, "config.ini")
    if not os.path.isfile(ini_path):
        raise FileNotFoundError(ini_path)

    cp = configparser.ConfigParser()
    cp.read(ini_path, encoding="utf-8")

    if not cp.has_option("app", "version"):
        raise KeyError("config.ini [app] version is required")

    return cp.get("app", "version")


def _load_third_party_licenses(base_dir: str) -> str:
    path = os.path.join(
        base_dir, "licensing", "THIRD_PARTY_LICENSES.txt"
    )
    if not os.path.isfile(path):
        raise FileNotFoundError(path)

    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def main():
    # ----------------------------------------
    # base dir
    # ----------------------------------------
    base_dir = _resolve_base_dir()

    # ----------------------------------------
    # logger / context
    # ----------------------------------------
    logger = init_logger(base_dir)
    ctx = load_context(base_dir, logger)

    logger.info(f"[DIFF_MAIN] base_dir={base_dir}")

    # ----------------------------------------
    # app version
    # ----------------------------------------
    ctx.app_version = _load_app_version(base_dir)
    logger.info(f"[DIFF_MAIN] version={ctx.app_version}")

    # ----------------------------------------
    # third party licenses
    # ----------------------------------------
    ctx.third_party_licenses_text = _load_third_party_licenses(base_dir)
    logger.info(
        "[DIFF_MAIN] third party licenses loaded "
        f"size={len(ctx.third_party_licenses_text)}"
    )

    # ----------------------------------------
    # App 起動
    # ----------------------------------------
    app = ExcelApp(ctx, logger)

    # Diff タブ以外を削除
    diff_label = ctx.labels.get("section_diff", "Diff")
    for i in reversed(range(app.nb.index("end"))):
        try:
            if app.nb.tab(i, "text") != diff_label:
                app.nb.forget(i)
                logger.info(f"[DIFF_MAIN] removed tab idx={i}")
        except Exception as ex:
            logger.warning(
                f"[DIFF_MAIN] tab remove failed idx={i} err={ex}"
            )

    if app.nb.index("end") == 1:
        app.nb.select(0)

    app.run()


if __name__ == "__main__":
    main()
