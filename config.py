import os
import json
from pathlib import Path
from dotenv import load_dotenv, dotenv_values

# --- 定数 ---
CONFIG_FILE_PATH = "config.json"

# --- .envからの読み込み ---
load_dotenv()

# Notion関連の秘匿情報
NOTION_API_TOKEN = os.getenv("NOTION_API_TOKEN")
PAGE_ID_CONTAINING_DB = os.getenv("NOTION_DATABASE_ID")

# SMTPサーバー情報
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))

# --- config.jsonからの読み込み ---
def load_json_config():
    """config.jsonから設定を読み込む"""
    if not os.path.exists(CONFIG_FILE_PATH):
        # デフォルト設定でファイルを作成
        default_config = {
            "EXCEL_TEMPLATE_PATH": "C:\\Users\\SEIZOU-20\\Desktop\\注文書.xlsx",
            "PDF_SAVE_DIR": os.path.join(Path.home(), "Desktop", "注文書")
        }
        save_json_config(default_config)
        return default_config
    
    try:
        with open(CONFIG_FILE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        # ファイルが空または壊れている場合、デフォルト設定を返す
        return {
            "EXCEL_TEMPLATE_PATH": "C:\\Users\\SEIZOU-20\\Desktop\\注文書.xlsx",
            "PDF_SAVE_DIR": os.path.join(Path.home(), "Desktop", "注文書")
        }

def save_json_config(data):
    """設定をconfig.jsonに保存する"""
    with open(CONFIG_FILE_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

# アプリケーション全体で使う設定変数を初期化
config_data = load_json_config()
EXCEL_TEMPLATE_PATH = config_data.get("EXCEL_TEMPLATE_PATH", "")
PDF_SAVE_DIR = config_data.get("PDF_SAVE_DIR", os.path.join(Path.home(), "Desktop", "注文書"))

# --- .envからのメールアカウント読み込み ---
def load_email_accounts():
    """
    .envファイルからEMAIL_SENDER_xxとEMAIL_PASSWORD_xxのペアを読み込み、
    アカウント情報を辞書として返します。
    """
    config = dotenv_values()
    accounts = {}
    sender_prefix = "EMAIL_SENDER_"
    sender_keys = [key for key in config if key.startswith(sender_prefix)]
    for key in sender_keys:
        account_name = key[len(sender_prefix):]
        password_key = f"EMAIL_PASSWORD_{account_name}"
        sender_email = config.get(key)
        password = config.get(password_key)
        if sender_email and password:
            accounts[account_name] = {"sender": sender_email, "password": password}
    return accounts

# --- 設定の動的更新 ---
def reload_config():
    """config.jsonを再読み込みして、グローバル変数を更新する"""
    global EXCEL_TEMPLATE_PATH, PDF_SAVE_DIR
    config_data = load_json_config()
    EXCEL_TEMPLATE_PATH = config_data.get("EXCEL_TEMPLATE_PATH", "")
    PDF_SAVE_DIR = config_data.get("PDF_SAVE_DIR", os.path.join(Path.home(), "Desktop", "注文書"))