import os
import json
from dotenv import load_dotenv

# .envファイルからNotionのトークンなどを読み込む
load_dotenv()

# --- JSONファイルから設定を一括で読み込む ---
def _load_settings_from_json(file_path="email_accounts.json"):
    """設定ファイル(JSON)を読み込む内部関数"""
    if not os.path.exists(file_path):
        print(f"警告: 設定ファイルが見つかりません: {file_path}")
        return {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"設定ファイルの読み込み中にエラーが発生しました: {e}")
        return {}

# 起動時に一度だけ設定を読み込む
_settings = _load_settings_from_json()

# --- 各種設定を変数としてエクスポート ---

# Notion関連 (引き続き.envから)
NOTION_API_TOKEN = os.getenv("NOTION_API_TOKEN")
PAGE_ID_CONTAINING_DB = os.getenv("NOTION_DATABASE_ID")
NOTION_SUPPLIER_DATABASE_ID = os.getenv("NOTION_SUPPLIER_DATABASE_ID")

# SMTPサーバー情報 (JSONから)
SMTP_SERVER = _settings.get("smtp_server", "smtp.office365.com")
SMTP_PORT = int(_settings.get("smtp_port", 587))

# パス設定 (.envから)
EXCEL_TEMPLATE_PATH = os.getenv("EXCEL_TEMPLATE_PATH")
PDF_SAVE_DIR = os.getenv("PDF_SAVE_DIR")

# メールアカウント情報 (JSONから)
def load_email_accounts():
    """
    読み込まれた設定からメールアカウント情報を返します。
    """
    return _settings.get("accounts", {})