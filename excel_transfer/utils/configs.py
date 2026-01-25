import os
import yaml
from dataclasses import dataclass
from typing import Dict, Any


# ==================================================
# YAML helpers
# ==================================================
def _load_yaml(path: str, default=None):
    if not os.path.exists(path):
        return default if default is not None else {}
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _save_yaml(path: str, data: dict):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True)


# ==================================================
# デフォルト定義（公開 EXE の正）
# ==================================================
DEFAULT_APP_SETTINGS = {
    "app": {
        "window_size": "900x650",
        "default_dir": "",
    },
    "components": {
        "transfer_config": {"row": 0, "width": 70},
        "grep_root": {"row": 0, "width": 70},
        "grep_keyword": {"row": 1, "width": 40},
        "diff_file_a": {"row": 0, "width": 70},
        "diff_file_b": {"row": 1, "width": 70},
        "diff_key_cols": {"row": 2, "width": 40},
    },
}


# ==================================================
# 起動保証用ラベル（JA 完全集合）
# ==================================================
DEFAULT_LABELS_JA = {
    "app_title": "Offismith",

    "section_excel": "Excel",
    "section_transfer": "転記",
    "section_grep": "Grep",
    "section_diff": "Diff",
    "section_count": "Count",

    "label_log": "ログ",
    "button_run": "実行",
    "button_ellipsis": "...",
    "button_html_report": "HTMLレポート",

    "menu_license": "License",
    "menu_about": "About",
    "menu_third_party_licenses": "Third Party Licenses",
}


# ==================================================
# タブ表示デフォルト（INI 廃止）
# ==================================================
DEFAULT_TABS_ENABLED = {
    "tab1": True,  # ExcelViewer
    "tab2": True,  # Transfer
    "tab3": True,  # Grep
    "tab4": True,  # Diff
    "tab5": True,  # Count
}


# ==================================================
# Context
# ==================================================
@dataclass
class AppContext:
    base_dir: str
    output_dir: str
    config_dir: str
    labels: Dict[str, Any]
    app_settings: Dict[str, Any]
    user_paths: Dict[str, Any]
    user_paths_file: str
    tabs_enabled: Dict[str, bool]

    def save_user_path(self, key: str, value: str):
        self.user_paths[key] = value
        _save_yaml(self.user_paths_file, self.user_paths)

    def default_dir_for(self, current: str = "") -> str:
        if current and os.path.exists(os.path.dirname(current)):
            return os.path.dirname(current)
        return self.app_settings.get("app", {}).get("default_dir", "") or self.base_dir


# ==================================================
# labels loader（言語別・フォールバック無し）
# ==================================================
def _load_labels(config_dir: str, lang: str) -> Dict[str, Any]:
    path = os.path.join(config_dir, "labels", f"{lang}.yaml")
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    return _load_yaml(path, default=None)


# ==================================================
# Context loader
# ==================================================
def load_context(base_dir: str, logger) -> AppContext:
    # ------------------------------------------
    # directories
    # ------------------------------------------
    config_dir = os.path.join(base_dir, "data", "config")
    os.makedirs(config_dir, exist_ok=True)

    output_dir = os.path.join(base_dir, "outputs")
    os.makedirs(output_dir, exist_ok=True)

    # ------------------------------------------
    # app settings
    # ------------------------------------------
    app_settings = DEFAULT_APP_SETTINGS.copy()
    app_settings.update(
        _load_yaml(
            os.path.join(config_dir, "app_settings.yaml"),
            default={},
        )
    )
    logger.debug(f"[CONFIG] app_settings={app_settings}")

    # ------------------------------------------
    # user paths（先に読む）
    # ------------------------------------------
    user_paths_file = os.path.join(base_dir, "user_paths.yaml")
    user_paths = _load_yaml(user_paths_file, default={})

    # ------------------------------------------
    # language（次回起動用）
    # ------------------------------------------
    lang = user_paths.get("app_lang")
    if not lang:
        lang = "ja"
    logger.info(f"[CONFIG] language={lang}")

    # ------------------------------------------
    # labels
    # ------------------------------------------
    labels = DEFAULT_LABELS_JA.copy()
    labels.update(_load_labels(config_dir, lang))
    logger.debug(
        f"[CONFIG] labels(lang={lang}) keys={sorted(labels.keys())}"
    )

    # ------------------------------------------
    # tabs
    # ------------------------------------------
    tabs_enabled = DEFAULT_TABS_ENABLED.copy()
    logger.info("[CONFIG] tabs_enabled=default (INI disabled)")

    ctx = AppContext(
        base_dir=base_dir,
        output_dir=output_dir,
        config_dir=config_dir,
        labels=labels,
        app_settings=app_settings,
        user_paths=user_paths,
        user_paths_file=user_paths_file,
        tabs_enabled=tabs_enabled,
    )

    # lang を Context に保持
    ctx.lang = lang

    return ctx
