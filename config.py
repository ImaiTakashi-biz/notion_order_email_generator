import os
import json
from dotenv import load_dotenv
from typing import Dict, Any, List, Tuple

# .envファイルからNotionのトークンなどを読み込む
load_dotenv()

# --- JSONファイルから設定を一括で読み込む ---
def _load_settings_from_json(file_path: str = "email_accounts.json") -> Dict[str, Any]:
    """設定ファイル(JSON)を読み込む内部関数"""
    if not os.path.exists(file_path):
        return {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            # json.loadがNoneを返す可能性も考慮
            settings = json.load(f)
            return settings if isinstance(settings, dict) else {}
    except (json.JSONDecodeError, IOError):
        return {}

# 起動時に一度だけ設定を読み込む
_settings: Dict[str, Any] = _load_settings_from_json()

# --- 各種設定を変数としてエクスポート ---

# Notion関連 (引き続き.envから)
NOTION_API_TOKEN: str = os.getenv("NOTION_API_TOKEN", "")
PAGE_ID_CONTAINING_DB: str = os.getenv("NOTION_DATABASE_ID", "")
NOTION_SUPPLIER_DATABASE_ID: str = os.getenv("NOTION_SUPPLIER_DATABASE_ID", "")

# SMTPサーバー情報 (JSONから)
SMTP_SERVER: str = _settings.get("smtp_server", "smtp.office365.com")
SMTP_PORT: int = int(_settings.get("smtp_port", 587))

# パス設定 (.envから)
PDF_SAVE_DIR: str = os.getenv("PDF_SAVE_DIR", "")

# アプリケーション定数
class AppConstants:
    # UI関連
    SPINNER_ANIMATION_DELAY: int = 80  # ミリ秒
    QUEUE_CHECK_INTERVAL: int = 100    # ミリ秒
    
    # Notion API関連
    NOTION_API_DELAY: float = 0.35       # 秒
    
    # 会社情報
    COMPANY_INFO: Dict[str, str] = {
        'name': '株式会社 新井精密',
        'postal_code': '〒368-0061',
        'address': '埼玉県秩父市小柱670',
        'tel_base': '0494-26-7786',
        'url': 'http://araiseimitsu.com/'
    }
    
    # メールテンプレート
    EMAIL_TEMPLATE: Dict[str, str] = {
        'subject': '注文書送付の件',
        'greeting': 'いつも大変お世話になります。',
        'body': '添付の通り注文宜しくお願い致します。'
    }

# メールアカウント情報 (JSONから)
def load_email_accounts() -> Dict[str, Any]:
    """
    読み込まれた設定からメールアカウント情報を返します。
    """
    return _settings.get("accounts", {})

def load_department_defaults() -> Dict[str, str]:
    """
    読み込まれた設定から部署ごとのデフォルトアカウント情報を返します。
    """
    return _settings.get("department_defaults", {})

def load_departments() -> List[str]:
    """
    読み込まれた設定から部署のリストを返します。
    """
    return _settings.get("departments", [])

def load_department_guidance_numbers() -> Dict[str, str]:
    """
    読み込まれた設定から部署ごとのガイダンス番号情報を返します。
    """
    return _settings.get("department_guidance_numbers", {})


def validate_config() -> Tuple[bool, List[str]]:
    """設定値の妥当性を検証する"""
    errors: List[str] = []
    
    # 必須環境変数のチェック
    required_env_vars: Dict[str, str] = {
        "NOTION_API_TOKEN": NOTION_API_TOKEN,
        "NOTION_DATABASE_ID": PAGE_ID_CONTAINING_DB,
        "NOTION_SUPPLIER_DATABASE_ID": NOTION_SUPPLIER_DATABASE_ID,
        "PDF_SAVE_DIR": PDF_SAVE_DIR
    }
    
    for var_name, var_value in required_env_vars.items():
        if not var_value:
            errors.append(f"環境変数 {var_name} が設定されていません")
    
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

def save_settings(json_data: Dict[str, Any]) -> Tuple[bool, str]:
    """
    GUIから受け取った設定をJSONファイルに保存する
    """
    try:
        # JSON ファイルの保存
        with open("email_accounts.json", 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        
        return True, "設定を保存しました。"
    except Exception as e:
        return False, f"設定の保存中にエラーが発生しました: {e}"