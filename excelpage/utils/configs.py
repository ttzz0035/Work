# excel_transfer/utils/configs.py
import os, yaml, json
from dataclasses import dataclass
from typing import Dict, Any

def _load_yaml(path: str, default=None):
    if not os.path.exists(path):
        return default if default is not None else {}
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}

def _save_yaml(path: str, data: dict):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True)

@dataclass
class AppContext:
    base_dir: str
    output_dir: str
    config_dir: str
    labels: Dict[str, Any]
    app_settings: Dict[str, Any]
    user_paths: Dict[str, Any]
    user_paths_file: str

    def save_user_path(self, key: str, value: str):
        self.user_paths[key] = value
        _save_yaml(self.user_paths_file, self.user_paths)

    def default_dir_for(self, current: str = "") -> str:
        if current and os.path.exists(os.path.dirname(current)): 
            return os.path.dirname(current)
        return self.app_settings.get("app", {}).get("default_dir", "") or self.base_dir

def load_context(base_dir: str, logger) -> AppContext:
    config_dir = os.path.join(base_dir, "data", "config")
    os.makedirs(config_dir, exist_ok=True)
    output_dir = os.path.join(base_dir, "outputs")
    os.makedirs(output_dir, exist_ok=True)

    app_settings = _load_yaml(os.path.join(config_dir, "app_settings.yaml"), default={
        "app": {"window_size": "900x650", "default_dir": ""},
        "components": {
            "transfer_config": {"row":0,"width":70},
            "grep_root": {"row":0,"width":70},
            "grep_keyword": {"row":1,"width":40},
            "diff_file_a": {"row":0,"width":70},
            "diff_file_b": {"row":1,"width":70},
            "diff_key_cols": {"row":2,"width":40}
        }
    })

    # ★ labels.yaml → label.yml に変更
    labels = _load_yaml(os.path.join(config_dir, "label.yml"), default={
        "ja": {
            "app_title": "Excel ユーティリティ（転記 / Grep / Diff / Count）",
            "section_transfer": "転記",
            "section_grep": "Grep（横断検索）",
            "section_diff": "Diff（差分）",
            "label_transfer_config": "転記定義CSV",
            "button_transfer": "実行",
            "label_grep_root": "検索ルートフォルダ",
            "label_grep_keyword": "検索キーワード",
            "button_grep": "実行",
            "label_diff_file_a": "File A",
            "label_diff_file_b": "File B",
            "label_diff_key_cols": "主キー列（カンマ区切り）",
            "button_diff": "実行",
            "check_ignore_case": "大文字小文字を無視",
            "check_compare_formula": "数式で比較（表示値ではなく）"
        }
    })["ja"]

    user_paths_file = os.path.join(base_dir, "user_paths.yaml")
    user_paths = _load_yaml(user_paths_file, default={})

    return AppContext(
        base_dir=base_dir,
        output_dir=output_dir,
        config_dir=config_dir,
        labels=labels,
        app_settings=app_settings,
        user_paths=user_paths,
        user_paths_file=user_paths_file
    )
