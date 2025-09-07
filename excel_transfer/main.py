# excel_transfer/main.py
import os, sys
from utils.log import init_logger
from utils.configs import load_context
from ui.app import ExcelApp

def main():
    base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    logger = init_logger(base_dir)
    ctx = load_context(base_dir, logger)
    ExcelApp(ctx, logger).run()

if __name__ == "__main__":
    main()
