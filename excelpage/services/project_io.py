import json
from typing import Any, Dict
from logger import get_logger

logger = get_logger("ProjectIO")


def save_project(path: str, tree_view) -> None:
    logger.info("[ProjectIO] save start path=%s", path)

    data: Dict[str, Any] = tree_view.export_project()

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    logger.info("[ProjectIO] save completed")


def load_project(path: str, tree_view) -> None:
    logger.info("[ProjectIO] load start path=%s", path)

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    tree_view.import_project(data)

    logger.info("[ProjectIO] load completed")
