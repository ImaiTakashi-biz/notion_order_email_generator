import os
import sys
import json
from dotenv import load_dotenv
from typing import Dict, Any, List, Tuple

def _get_resource_path(relative_path: str) -> str:
    """
    PyInstallerでビルドされた場合と通常実行の場合の両方に対応してリソースパスを取得する
    
    Args:
        relative_path: リソースファイルの相対パス
        
    Returns:
        リソースファイルの絶対パス
    """
    # PyInstallerでビルドされた場合、一時ディレクトリのパスがsys._MEIPASSに設定される
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        # 通常実行時はスクリプトのディレクトリを基準にする
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)

# .envファイルのパスを取得して読み込む
env_path = _get_resource_path(".env")
if os.path.exists(env_path):
    load_dotenv(env_path)
else:
    # .envファイルが存在しない場合でも、環境変数から読み込む
    load_dotenv()

# --- JSONファイルから設定を一括で読み込む ---
def _load_settings_from_json(file_path: str = "email_accounts.json") -> Dict[str, Any]:
    """設定ファイル(JSON)を読み込む内部関数"""
    # PyInstaller環境に対応したパスを取得
    resource_path = _get_resource_path(file_path)
    
    if not os.path.exists(resource_path):
        return {}
    try:
        with open(resource_path, 'r', encoding='utf-8') as f:
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

def load_department_name_mapping() -> Dict[str, str]:
    """
    読み込まれた設定から部署名のマッピング（表示名 → Notion名）を返します。
    マッピングが設定されていない場合は、表示名をそのまま返すための空の辞書を返します。
    """
    return _settings.get("department_name_mapping", {})

def convert_display_name_to_notion_name(display_name: str) -> str:
    """
    表示名（アプリで使用する名前）をNotion名に変換します。
    マッピングが存在しない場合は、表示名をそのまま返します。
    
    Args:
        display_name: アプリで表示する部署名
        
    Returns:
        Notionで使用する部署名
    """
    mapping = load_department_name_mapping()
    return mapping.get(display_name, display_name)

def convert_display_names_to_notion_names(display_names: List[str]) -> List[str]:
    """
    表示名のリストをNotion名のリストに変換します。
    
    Args:
        display_names: アプリで表示する部署名のリスト
        
    Returns:
        Notionで使用する部署名のリスト
    """
    return [convert_display_name_to_notion_name(name) for name in display_names]

def convert_notion_name_to_display_name(notion_name: str) -> str:
    """
    Notion名を表示名（アプリで使用する名前）に変換します。
    マッピングが存在しない場合は、Notion名をそのまま返します。
    
    Args:
        notion_name: Notionで使用する部署名
        
    Returns:
        アプリで表示する部署名
    """
    mapping = load_department_name_mapping()
    # マッピングの逆引き（値からキーを探す）
    for display_name, mapped_notion_name in mapping.items():
        if mapped_notion_name == notion_name:
            return display_name
    return notion_name


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
    PyInstaller環境では実行ファイルと同じディレクトリに保存する
    """
    try:
        # PyInstaller環境では実行ファイルのディレクトリに保存
        # 通常実行時はスクリプトのディレクトリに保存
        if hasattr(sys, '_MEIPASS'):
            # 実行ファイルのディレクトリを取得
            if getattr(sys, 'frozen', False):
                # PyInstallerでビルドされた場合
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        json_path = os.path.join(base_path, "email_accounts.json")
        
        # JSON ファイルの保存
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        
        return True, "設定を保存しました。"
    except Exception as e:
        return False, f"設定の保存中にエラーが発生しました: {e}"