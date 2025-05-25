import os
import sys
import configparser

def app_dir_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def resource_path(filename):
    return os.path.join(app_dir_path(), filename)

def load_ini(filename):
    path = resource_path(filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f"INIファイルが見つかりません: {path}")
    config = configparser.ConfigParser()
    with open(path, encoding="utf-8") as f:
        config.read_file(f)
    return config

def save_ini(config, filename):
    path = resource_path(filename)
    with open(path, "w", encoding="utf-8") as f:
        config.write(f)
