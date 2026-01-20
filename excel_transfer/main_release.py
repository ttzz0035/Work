# excel_transfer/main_release.py
import os
import sys
import configparser

from utils.log import init_logger
from utils.configs import load_context
from ui.app import ExcelApp

# >>> LICENSE
from licensing.license_state import LicenseManager, LicenseStatus
# <<< LICENSE


def _load_app_version(base_dir: str, logger) -> str:
    """
    config.ini から app.version を読む
    フォールバック禁止
    """
    ini_path = os.path.join(base_dir, "config.ini")
    logger.info(f"[APP] load version from {ini_path}")

    if not os.path.isfile(ini_path):
        raise FileNotFoundError(ini_path)

    cp = configparser.ConfigParser()
    cp.read(ini_path, encoding="utf-8")

    if not cp.has_option("app", "version"):
        raise KeyError("config.ini [app] version is required")

    return cp.get("app", "version")


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

    # ====================================================
    # >>> APP VERSION（ここで確定）
    # ====================================================
    try:
        ctx.app_version = _load_app_version(base_dir, logger)
        logger.info(f"[APP] version={ctx.app_version}")
    except Exception as ex:
        logger.error(f"[APP] version load failed: {ex}", exc_info=True)
        raise
    # ====================================================
    # <<< APP VERSION
    # ====================================================

    # ====================================================
    # >>> THIRD PARTY LICENSES
    # ====================================================
    try:
        lic_path = os.path.join(
            base_dir, "licensing", "THIRD_PARTY_LICENSES.txt"
        )
        logger.info(f"[LICENSE] load third party licenses: {lic_path}")

        if not os.path.isfile(lic_path):
            raise FileNotFoundError(lic_path)

        with open(lic_path, "r", encoding="utf-8") as f:
            ctx.third_party_licenses_text = f.read()

        logger.info(
            "[LICENSE] third party licenses loaded "
            f"size={len(ctx.third_party_licenses_text)}"
        )

    except Exception as ex:
        logger.error(
            f"[LICENSE] failed to load THIRD_PARTY_LICENSES.txt: {ex}",
            exc_info=True,
        )
        raise
    # ====================================================
    # <<< THIRD PARTY LICENSES
    # ====================================================

    # ====================================================
    # >>> LICENSE 判定
    # ====================================================
    try:
        lm = LicenseManager(
            subscription_sku_store_id="your_subscription_sku_store_id"
        )
        lic_state = lm.get_state()

        logger.info(
            f"[LICENSE] status={lic_state.status} "
            f"remain={lic_state.remaining_days}"
        )

        ctx.license_manager = lm
        ctx.license_status = lic_state.status
        ctx.license_remaining_days = lic_state.remaining_days
        ctx.automation_enabled = (
            lic_state.status == LicenseStatus.SUBSCRIBED
        )

    except Exception as ex:
        logger.error(f"[LICENSE] check failed: {ex}", exc_info=True)
        ctx.license_status = LicenseStatus.EXPIRED
        ctx.license_remaining_days = 0
        ctx.automation_enabled = False
    # ====================================================
    # <<< LICENSE
    # ====================================================

    # ----------------------------------------
    # App 起動
    # ----------------------------------------
    app = ExcelApp(ctx, logger)
    app.run()


if __name__ == "__main__":
    main()
