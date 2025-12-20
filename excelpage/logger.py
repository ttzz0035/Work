import logging
import os

def get_logger(name: str):
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    os.makedirs("logs", exist_ok=True)

    fh = logging.FileHandler("logs/app.log", encoding="utf-8")
    fmt = logging.Formatter(
        "%(asctime)s - %(levelname)s - [%(name)s] %(message)s"
    )
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger
