import smtplib
import os
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import parseaddr
from typing import Dict, Any, Optional, List, Tuple

import config
import keyring
from keyring.errors import KeyringError

SERVICE_NAME = "NotionOrderApp"
LOG_DIR = os.path.join(os.getcwd(), "logs")
EMAIL_LOG_PATH = os.path.join(LOG_DIR, "email_send.log")


def _append_email_log(status: str, supplier: str, detail: str = "") -> None:
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        safe_detail = detail.replace("\n", " ").strip()
        safe_supplier = supplier.replace("\n", " ").strip()
        with open(EMAIL_LOG_PATH, "a", encoding="utf-8") as fp:
            fp.write(f"{timestamp}\t{status}\t{safe_supplier}\t{safe_detail}\n")
    except Exception:
        # ログ記録が失敗してもユーザー処理は継続する
        pass


def _sanitize_header(value: str) -> str:
    return value.splitlines()[0].strip()


def _extract_addresses(value: str) -> List[str]:
    normalized = value.replace(";", ",")
    addresses: List[str] = []
    for chunk in normalized.split(","):
        addr = chunk.strip()
        if not addr:
            continue
        _, parsed = parseaddr(addr)
        if parsed:
            addresses.append(parsed)
    return addresses


def send_smtp_mail(
    info: Dict[str, Any],
    pdf_path: str,
    sender_creds: Dict[str, Any],
    account_name: str,
    selected_department: Optional[str] = None
) -> Tuple[bool, Optional[str]]:
    """SMTPサーバー経由でPDF添付メールを送信する"""
    try:
        sender_email = sender_creds.get("sender")
        password = sender_creds.get("password")

        if not sender_email or not password:
            message = "送信元メールアドレスまたはパスワードが不明です。"
            print(f"✗ メール送信エラー: {message}")
            return False, message

        msg = MIMEMultipart()
        msg["From"] = sender_email

        raw_to = (info.get("email") or "").strip()
        raw_cc = (info.get("email_cc") or "").strip()

        to_header = _sanitize_header(raw_to) if raw_to else ""
        cc_header = _sanitize_header(raw_cc) if raw_cc else ""

        to_addresses = _extract_addresses(to_header) if to_header else []
        cc_addresses = _extract_addresses(cc_header) if cc_header else []

        if not to_addresses:
            message = "宛先メールアドレスが設定されていません。"
            print(f"✗ メール送信エラー: {message}")
            return False, message

        msg["To"] = ", ".join(to_addresses)
        if cc_addresses:
            msg["Cc"] = ", ".join(cc_addresses)

        template = config.AppConstants.EMAIL_TEMPLATE
        company = config.AppConstants.COMPANY_INFO
        msg["Subject"] = template['subject']

        if selected_department:
            order_contact = f"{selected_department} {account_name}"
        else:
            order_contact = account_name

        guidance_numbers = config.load_department_guidance_numbers()
        raw_guidance = guidance_numbers.get(selected_department, "")
        guidance_number = "".join(filter(str.isdigit, raw_guidance))
        tel_line = f"TEL: {company['tel_base']}" + (f"（ガイダンス{guidance_number}番）" if guidance_number else "")

        body = (
            f"{info['supplier_name']}\n"
            f"{info.get('sales_contact', 'ご担当者')} 様\n\n"
            f"{template['greeting']}\n"
            f"{template['body']}\n\n"
            "∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝\n"
            f"{company['name']}\n"
            f"発注担当： {order_contact}\n"
            f"{company['postal_code']} {company['address']}\n"
            f"Email: {sender_creds['sender']}\n"
            f"{tel_line}\n"
            f"URL: {company['url']}\n"
            "∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝"
        )
        msg.attach(MIMEText(body, 'plain'))

        with open(pdf_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(pdf_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_path)}"'
        msg.attach(part)

        with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as server:
            server.starttls()
            server.login(sender_email, password)
            recipients = to_addresses + cc_addresses
            server.sendmail(sender_email, recipients, msg.as_string())

        return True, None
    except smtplib.SMTPAuthenticationError:
        message = "SMTP認証に失敗しました。ログイン情報を確認してください。"
        print(f"✗ メール送信エラー: {message}")
        return False, message
    except (smtplib.SMTPConnectError, ConnectionRefusedError, OSError) as exc:
        message = f"SMTPサーバー({config.SMTP_SERVER}:{config.SMTP_PORT})に接続できません。詳細: {exc}"
        print(f"✗ メール送信エラー: {message}")
        return False, message
    except Exception as exc:
        message = f"予期せぬエラーが発生しました: {exc}"
        print(f"✗ メール送信エラー: {message}")
        return False, message


def prepare_and_send_order_email(
    account_key: str,
    sender_creds: Dict[str, Any],
    items: List[Dict[str, Any]],
    pdf_path: str,
    selected_department: Optional[str] = None
) -> Tuple[bool, Optional[str]]:
    """UIからの情報をもとにメール送信の準備と実行を行う"""
    sender_email = sender_creds.get("sender")
    display_name = sender_creds.get("display_name", account_key)

    try:
        password = keyring.get_password(SERVICE_NAME, sender_email or "")
    except KeyringError as err:
        message = f"OSの資格情報ストアにアクセスできませんでした ({err})."
        print(f"✗ パスワード取得エラー: {message}")
        return False, message
    if not password:
        message = f"{sender_email} のパスワードがOSに保存されていません。"
        print(f"✗ パスワード取得エラー: {message}")
        return False, message

    creds_with_pass = sender_creds.copy()
    creds_with_pass["password"] = password

    if not items:
        message = "対象アイテムがありません。"
        print(f"✗ メール送信エラー: {message}")
        return False, message
    if not pdf_path or not os.path.exists(pdf_path):
        message = "添付するPDFファイルが見つかりません。"
        print(f"✗ メール送信エラー: {message}")
        return False, message

    supplier_name = items[0].get("supplier_name", "")
    success, error_message = send_smtp_mail(
        info=items[0],
        pdf_path=pdf_path,
        sender_creds=creds_with_pass,
        account_name=display_name,
        selected_department=selected_department
    )

    if success:
        _append_email_log("success", supplier_name, "メール送信に成功")
        return True, None

    detail = error_message or "送信に失敗しました"
    _append_email_log("failed", supplier_name, detail)
    return False, detail
