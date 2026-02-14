import logging
import sys
import os  # ★ 追加
from typing import Optional

class Logger:
    level_values = {
        "CRITICAL": logging.CRITICAL,
        "FATAL": logging.CRITICAL,
        "ERROR": logging.ERROR,
        "WARNING": logging.WARNING,
        "WARN": logging.WARNING,
        "INFO": logging.INFO,
        "DEBUG": logging.DEBUG,
        "TRACE": 5,  # 追加: 独自TRACEレベル
        "NOTSET": logging.NOTSET,
    }

    def __init__(self, name: str, log_file_path: str = "", level: str = "INFO"):
        # TRACEレベルを正式登録（INFOより下）
        logging.addLevelName(5, "TRACE")

        self.logger = logging.getLogger(name)

        # 同じLogger名でハンドラが重複しないよう初期化
        if self.logger.handlers:
            for h in self.logger.handlers[:]:
                self.logger.removeHandler(h)

        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - [%(name)s] %(message)s',
            '%Y-%m-%d %H:%M:%S'
        )

        # --- コンソール出力 ---
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)

        # --- ファイル出力（任意） ---
        if log_file_path:
            # ★ ここでディレクトリを自動作成する
            log_dir = os.path.dirname(log_file_path)
            if log_dir:
                os.makedirs(log_dir, exist_ok=True)

            file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)

        # --- ログレベル設定 ---
        self.setLogLevel(level)

    # ------------------------------------
    #  共通API
    # ------------------------------------
    def setLogLevel(self, level: str = "INFO") -> bool:
        upper = level.upper()
        if upper not in self.level_values:
            self.logger.warning(f"Unknown log level: {level}")
            return False
        self.logger.setLevel(self.level_values[upper])
        return True

    def trace(self, message: str):
        self.logger.log(5, message)

    def debug(self, message: str):
        self.logger.debug(message)

    def info(self, message: str):
        self.logger.info(message)

    def warning(self, message: str):
        self.logger.warning(message)

    def error(self, message: str):
        self.logger.error(message)

    def critical(self, message: str):
        self.logger.critical(message)
