# excel_transfer/utils/log.py
import os, logging

def init_logger(base_dir: str) -> logging.Logger:
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    logger = logging.getLogger("excel_transfer")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    if not logger.handlers:
        fh = logging.FileHandler(os.path.join(log_dir, "app.log"), encoding="utf-8")
        fh.setFormatter(fmt); logger.addHandler(fh)
        sh = logging.StreamHandler(); sh.setFormatter(fmt); logger.addHandler(sh)
    return logger
