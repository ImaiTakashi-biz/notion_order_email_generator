import os
from dotenv import load_dotenv, dotenv_values

# --- .envからの読み込み ---
load_dotenv()

# Notion関連の秘匿情報
NOTION_API_TOKEN = os.getenv("NOTION_API_TOKEN")
PAGE_ID_CONTAINING_DB = os.getenv("NOTION_DATABASE_ID")

# SMTPサーバー情報
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))

# パス設定 (.envファイルにフルパスで指定してください)
EXCEL_TEMPLATE_PATH = os.getenv("EXCEL_TEMPLATE_PATH")
PDF_SAVE_DIR = os.getenv("PDF_SAVE_DIR")

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
