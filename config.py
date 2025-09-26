import os
import json
from dotenv import load_dotenv
from cryptography.fernet import Fernet

# .envファイルからNotionのトークンなどを読み込む
load_dotenv()

# --- JSONファイルから設定を一括で読み込む ---
def _load_settings_from_json(file_path="email_accounts.json"):
    """設定ファイル(JSON)を読み込む内部関数"""
    if not os.path.exists(file_path):
        return {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        return {}

# 起動時に一度だけ設定を読み込む
_settings = _load_settings_from_json()

# --- 暗号化関連 ---
def _get_encryption_key():
    """環境変数から暗号化キーを取得または生成する"""
    key = os.getenv("EMAIL_ENCRYPTION_KEY")
    if key is None:
        key = Fernet.generate_key().decode()
        # .envファイルにキーを追記
        with open('.env', 'a') as f:
            f.write(f'\nEMAIL_ENCRYPTION_KEY="{key}"')
        load_dotenv() # .envを再読み込み
        print("新しい暗号化キーを生成し、.envファイルに保存しました。")
    return key.encode()

_fernet = Fernet(_get_encryption_key())

def encrypt_password(password):
    """パスワードを暗号化する"""
    if not password:
        return ""
    return _fernet.encrypt(password.encode()).decode()

def decrypt_password(encrypted_password):
    """暗号化されたパスワードを復号化する"""
    if not encrypted_password:
        return ""
    try:
        return _fernet.decrypt(encrypted_password.encode()).decode()
    except Exception:
        # 復号化に失敗した場合（キーが異なる、データが破損しているなど）
        return ""

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

# アプリケーション定数
class AppConstants:
    # UI関連
    SPINNER_ANIMATION_DELAY = 80  # ミリ秒
    QUEUE_CHECK_INTERVAL = 100    # ミリ秒
    
    # Notion API関連
    NOTION_API_DELAY = 0.35       # 秒
    
    # Excel関連
    EXCEL_CELLS = {
        'SUPPLIER_NAME': 'A5',
        'SALES_CONTACT': 'A7', 
        'SENDER_NAME': 'D8',
        'SENDER_EMAIL': 'D14',
        'ITEM_START_ROW': 16
    }
    
    # 会社情報
    COMPANY_INFO = {
        'name': '株式会社 新井精密',
        'postal_code': '〒368-0061',
        'address': '埼玉県秩父市小柱670番地',
        'tel': 'TEL: 0494-26-7786',
        'fax': 'FAX: 0494-26-7787'
    }
    
    # メールテンプレート
    EMAIL_TEMPLATE = {
        'subject': '注文書送付の件',
        'greeting': 'いつも大変お世話になります。',
        'body': '添付の通り注文宜しくお願い致します。'
    }

# メールアカウント情報 (JSONから)
def load_email_accounts():
    """
    読み込まれた設定からメールアカウント情報を返します。
    パスワードは復号化して返します。
    """
    accounts = _settings.get("accounts", {})
    # 辞書のコピーを作成して変更を反映
    decrypted_accounts = {}
    for key, details in accounts.items():
        decrypted_details = details.copy()
        if "password" in decrypted_details:
            decrypted_details["password"] = decrypt_password(decrypted_details["password"])
        decrypted_accounts[key] = decrypted_details
    return decrypted_accounts

def load_department_defaults():
    """
    読み込まれた設定から部署ごとのデフォルトアカウント情報を返します。
    """
    return _settings.get("department_defaults", {})

def load_departments():
    """
    読み込まれた設定から部署のリストを返します。
    """
    return _settings.get("departments", [])

def load_department_guidance_numbers():
    """
    読み込まれた設定から部署ごとのガイダンス番号情報を返します。
    """
    return _settings.get("department_guidance_numbers", {})


def validate_config():
    """設定値の妥当性を検証する"""
    errors = []
    
    # 必須環境変数のチェック
    required_env_vars = {
        "NOTION_API_TOKEN": NOTION_API_TOKEN,
        "NOTION_DATABASE_ID": PAGE_ID_CONTAINING_DB,
        "NOTION_SUPPLIER_DATABASE_ID": NOTION_SUPPLIER_DATABASE_ID,
        "EXCEL_TEMPLATE_PATH": EXCEL_TEMPLATE_PATH,
        "PDF_SAVE_DIR": PDF_SAVE_DIR
    }
    
    for var_name, var_value in required_env_vars.items():
        if not var_value:
            errors.append(f"環境変数 {var_name} が設定されていません")
    
    # ファイル存在チェック
    if EXCEL_TEMPLATE_PATH and not os.path.exists(EXCEL_TEMPLATE_PATH):
        errors.append(f"Excelテンプレートファイルが見つかりません: {EXCEL_TEMPLATE_PATH}")
    
    if PDF_SAVE_DIR and not os.path.exists(PDF_SAVE_DIR):
        try:
            os.makedirs(PDF_SAVE_DIR, exist_ok=True)
        except Exception as e:
            errors.append(f"PDF保存ディレクトリを作成できません: {PDF_SAVE_DIR} - {e}")
    
    # アカウント設定のチェック
    accounts = load_email_accounts()
    if not accounts:
        errors.append("メールアカウントが設定されていません")
    
    return len(errors) == 0, errors

def save_settings(json_data):
    """
    GUIから受け取った設定をJSONファイルに保存する
    パスワードは暗号化して保存します。
    """
    try:
        # パスワードを暗号化
        accounts_to_save = {}
        for key, details in json_data.get("accounts", {}).items():
            details_to_save = details.copy()
            if "password" in details_to_save and details_to_save["password"]:
                details_to_save["password"] = encrypt_password(details_to_save["password"])
            accounts_to_save[key] = details_to_save
        json_data["accounts"] = accounts_to_save

        # JSON ファイルの保存
        with open("email_accounts.json", 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        
        return True, "設定を保存しました。"
    except Exception as e:
        return False, f"設定の保存中にエラーが発生しました: {e}"