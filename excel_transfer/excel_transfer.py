import csv
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import yaml
import logging
import sys
import datetime
import xlwings as xw
import re
import gc
from openpyxl.utils import column_index_from_string, get_column_letter

# --- ベースディレクトリ設定 ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_DIR = os.path.join(BASE_DIR, "config")
LOG_DIR = os.path.join(BASE_DIR, "logs")
USER_PATHS_FILE = os.path.join(BASE_DIR, "user_paths.yaml")
os.makedirs(LOG_DIR, exist_ok=True)

# --- ログ設定 ---
logger = logging.getLogger("excel_transfer")
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

file_handler = logging.FileHandler(os.path.join(LOG_DIR, "app.log"), encoding="utf-8")
file_handler.setFormatter(formatter)

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(stream_handler)

# --- ファイル操作ユーティリティ ---
def load_yaml(file_path, default_data=None):
    if not os.path.exists(file_path):
        if default_data is not None:
            save_yaml(file_path, default_data)
            return default_data
        else:
            return {}
    with open(file_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def save_yaml(file_path, data):
    with open(file_path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True)

# --- バックアップ作成 ---
def backup_file(file_path):
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    backup_path = f"{os.path.splitext(file_path)[0]}_backup_{timestamp}.xlsx"
    shutil.copy2(file_path, backup_path)
    logger.info(f"バックアップ作成: {backup_path}")
    return backup_path

# --- UI用ログ追記 ---
def append_log(message):
    log_textbox.insert(tk.END, message + "\n")
    log_textbox.see(tk.END)

def copy_log_to_clipboard():
    root.clipboard_clear()
    root.clipboard_append(log_textbox.get("1.0", tk.END))
    messagebox.showinfo("クリップボード", "ログをクリップボードにコピーしました。")

def apply_offset(cell, row_offset=0, col_offset=0):
    match = re.match(r"([A-Z]+)(\d+)", cell.strip())
    if not match:
        raise ValueError(f"無効なセル形式: {cell}")
    col_letter, row = match.groups()
    logger.debug(f"{col_letter},{int(col_offset)}")
    col_index = column_index_from_string(col_letter)
    if col_index + int(col_offset) < 0: raise ValueError(f"無効なセル形式: {cell}")
    new_col_letter = get_column_letter(col_index + int(col_offset))
    new_row = int(row) + int(row_offset)
    return f"{new_col_letter}{new_row}"

# --- Excel転記処理 ---
def transfer_excel_values_csv(config_csv_path, labels):
    log_textbox.delete("1.0", tk.END)
    append_log("=== 転記開始 ===")

    for csv_path in config_csv_path:
        if not os.path.exists(csv_path):
            logger.error(labels["error_file_not_found"])
            append_log(labels["error_file_not_found"])
            messagebox.showerror(labels["message_title_error"], f"{labels['error_file_not_found']}\n{csv_path}")
            return

    for csv_path in config_csv_path:
        workbooks = {}
        backups = {}
        try:
            try:
                with open(csv_path, newline="", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    jobs = list(reader)
            except UnicodeDecodeError:
                with open(csv_path, newline="", encoding="shift-jis") as f:
                    reader = csv.DictReader(f)
                    jobs = list(reader)

            if not jobs:
                logger.error(labels["log_transfer_no_jobs"])
                messagebox.showerror(labels["message_title_error"], labels["error_transfer_no_jobs"])
                return

            # バックアップ作成
            involved_files = set(job["destination_file"] for job in jobs)
            for file in involved_files:
                full_path = os.path.join(BASE_DIR, file)
                if os.path.exists(full_path):
                    backup_path = backup_file(full_path)
                    backups[file] = backup_path

            for job in jobs:
                src_path = os.path.join(BASE_DIR, job["source_file"])
                dst_path = os.path.join(BASE_DIR, job["destination_file"])

                if src_path not in workbooks:
                    if not os.path.exists(src_path):
                        raise FileNotFoundError(f"{src_path} が見つかりません。")
                    app = xw.App(visible=False, add_book=False)
                    workbooks[src_path] = app.books.open(src_path, read_only=False)

                if dst_path not in workbooks:
                    app = xw.App(visible=False, add_book=False)
                    if os.path.exists(dst_path):
                        workbooks[dst_path] = app.books.open(dst_path, read_only=False)
                    else:
                        workbooks[dst_path] = app.books.add()

                if job["source_sheet"] not in workbooks[src_path].sheet_names:
                    raise ValueError(f"転記元シートが存在しません: {job['source_sheet']} in {job['source_file']}")

                if job["destination_sheet"] not in workbooks[dst_path].sheet_names:
                    workbooks[dst_path].sheets.add(job["destination_sheet"])
                    append_log(f"シート作成: {job['destination_sheet']} in {job['destination_file']}")

                # オフセット取得
                src_row_offset = int(job.get("source_row_offset", 0) or 0)
                src_col_offset = int(job.get("source_col_offset", 0) or 0)
                dst_row_offset = int(job.get("destination_row_offset", 0) or 0)
                dst_col_offset = int(job.get("destination_col_offset", 0) or 0)

                src_cell = apply_offset(job["source_cell"], src_row_offset, src_col_offset)
                dst_cell = apply_offset(job["destination_cell"], dst_row_offset, dst_col_offset)

                value = workbooks[src_path].sheets[job["source_sheet"]].range(src_cell).value or ""
                workbooks[dst_path].sheets[job["destination_sheet"]].range(dst_cell).value = value

                msg = f"転記: {job['source_file']}[{job['source_sheet']}!{src_cell}] → {job['destination_file']}[{job['destination_sheet']}!{dst_cell}]"
                logger.info(msg)
                append_log(msg)

            # 保存
            for index, (path, wb) in enumerate(workbooks.items()):
                if index % 2 != 0:
                    wb.save(path)
                    append_log(f"保存完了: {path}")

        except Exception as e:
            logger.error(f"転記エラー: {e}")
            append_log(f"エラー発生: {e}")

        finally:
            for path, wb in workbooks.items():
                name = wb.name
                wb.close()
                append_log(f"{name} Close 完了")
            del workbooks
            gc.collect()

# --- ファイル選択 ---
def get_initial_dir(entry_widget, app_settings):
    current_path = entry_widget.get().strip()
    if current_path and os.path.exists(os.path.dirname(current_path)):
        return os.path.dirname(current_path)
    else:
        return app_settings["app"]["default_dir"]

def browse_file(entry_widget, filetypes, key, user_paths, app_settings, labels, save=False):
    try:
        initial_dir = get_initial_dir(entry_widget, app_settings)
        if save:
            path = filedialog.asksaveasfilename(initialdir=initial_dir, defaultextension=filetypes[0][1], filetypes=filetypes)
        else:
            path = filedialog.askopenfilenames(initialdir=initial_dir, filetypes=filetypes)
        if path:
            entry_widget.delete(0, tk.END)
            path = "?".join(path)
            entry_widget.insert(0,path)
            user_paths[key] = path
            save_yaml(USER_PATHS_FILE, user_paths)
            logger.info(f"{labels['log_file_selected']}: {path}")
    except Exception as e:
        logger.error(f"{labels['log_file_select_error']}: {e}")

# --- UI部品作成 ---
def create_ui_entry(frame, label_text, comps_key, user_path_key, filetypes, comps, user_paths, app_settings, labels, save=False):
    row = comps[comps_key]["row"]
    width = comps[comps_key]["width"]
    tk.Label(frame, text=labels[label_text]).grid(row=row, column=0, sticky="w")
    entry = tk.Entry(frame, width=width)
    entry.grid(row=row, column=1)
    entry.insert(0, user_paths.get(user_path_key, ""))
    tk.Button(frame, text="...", command=lambda: browse_file(entry, filetypes, user_path_key, user_paths, app_settings, labels, save)).grid(row=row, column=2)
    return entry

# --- メイン関数 ---
def main():
    global root, log_textbox
    app_settings = load_yaml(os.path.join(CONFIG_DIR, "app_settings.yaml"))
    labels_all = load_yaml(os.path.join(CONFIG_DIR, "labels.yaml"))
    lang = "ja"
    labels = labels_all[lang]
    user_paths = load_yaml(USER_PATHS_FILE, default_data={"transfer_config": ""})

    logger.info(labels["log_startup"])

    root = tk.Tk()
    root.title(labels["app_title"])
    root.geometry(app_settings["app"]["window_size"])

    comps = app_settings["components"]

    frame_transfer = tk.LabelFrame(root, text=labels["section_transfer"], padx=10, pady=10)
    frame_transfer.pack(fill="both", padx=10, pady=10)

    transfer_config_entry = create_ui_entry(
        frame_transfer, "label_transfer_config", "transfer_config", "transfer_config", [("CSV Files", "*.csv")],
        comps, user_paths, app_settings, labels
    )

    tk.Button(frame_transfer, text=labels["button_transfer"], height=2,
              command=lambda: transfer_excel_values_csv(transfer_config_entry.get().split("?"), labels)
              ).grid(row=0, column=3, pady=10)

    # ログ表示テキストボックス
    log_textbox = tk.Text(root, height=10)
    log_textbox.pack(fill="both", padx=10, pady=10)

    tk.Button(root, text="ログをクリップボードにコピー", command=copy_log_to_clipboard).pack(pady=5)

    root.mainloop()
    logger.info(labels["log_shutdown"])

if __name__ == "__main__":
    main()
