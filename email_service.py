import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import Dict, Any, Optional, List
import config
import keyring

# アプリケーション識別のためのサービス名 (settings_gui.pyと合わせる)
SERVICE_NAME = "NotionOrderApp"

def send_smtp_mail(info: Dict[str, Any], pdf_path: str, sender_creds: Dict[str, Any], account_name: str, selected_department: Optional[str] = None) -> bool:
    """
    SMTPサーバー経由でPDF添付メールを送信する
    """
    try:
        # --- パスワードをsender_credsから取得 ---
        sender_email = sender_creds.get("sender")
        password = sender_creds.get("password")

        if not sender_email or not password:
            print("❌ メール送信エラー: 送信元メールアドレスまたはパスワードが不明です。")
            return False
        # --- ここまでが変更点 ---

        msg = MIMEMultipart()
        msg["From"] = sender_creds["sender"]
        
        # CRLFインジェクション対策: メールヘッダーに設定する値から改行コードを除去
        to_email = info.get("email")
        cc_email = info.get("email_cc")
        
        if to_email:
            msg["To"] = to_email.splitlines()[0]
        
        if cc_email:
            msg["Cc"] = cc_email.splitlines()[0]
        template = config.AppConstants.EMAIL_TEMPLATE
        company = config.AppConstants.COMPANY_INFO
        
        msg["Subject"] = template['subject']

        # 部署名を発注担当名前に追加
        if selected_department:
            order_contact = f"{selected_department} {account_name}"
        else:
            order_contact = account_name

        # ガイダンス番号を取得
        guidance_numbers = config.load_department_guidance_numbers()
        raw_guidance = guidance_numbers.get(selected_department, "")
        guidance_number = "".join(filter(str.isdigit, raw_guidance))
        tel_line = f"TEL: {company['tel_base']}" + (f"（ガイダンス{guidance_number}番）" if guidance_number else "")

        # メール本文 (署名を動的に生成)
        body = f"""{info['supplier_name']}
{info.get('sales_contact', 'ご担当者')} 様

{template['greeting']}
{template['body']}

∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝
{company['name']}
発注担当： {order_contact}
{company['postal_code']} {company['address']}
Email: {sender_creds["sender"]}
{tel_line}
URL: {company['url']}
∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝"""
        msg.attach(MIMEText(body, 'plain'))

        # PDF添付
        with open(pdf_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(pdf_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_path)}"'
        msg.attach(part)

        # SMTPサーバーに接続して送信
        with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as server:
            server.starttls()
            server.login(sender_email, password)
            recipients = [info["email"]] + ([info["email_cc"]] if info["email_cc"] else [])
            server.sendmail(sender_creds["sender"], recipients, msg.as_string())
            
        return True
    except smtplib.SMTPAuthenticationError:
        print("❌ メール送信エラー: SMTP認証に失敗しました。")
        print("   -> メールアドレスまたはパスワードが間違っている可能性があります。")
        print("   -> 設定画面からアカウント情報をご確認ください。")
        return False
    except (smtplib.SMTPConnectError, ConnectionRefusedError, OSError) as e:
        print(f"❌ メール送信エラー: SMTPサーバー({config.SMTP_SERVER}:{config.SMTP_PORT})に接続できません。")
        print("   -> サーバーアドレス、ポート番号、ネットワーク接続を確認してください。")
        print(f"   -> 詳細: {e}")
        return False
    except Exception as e:
        print(f"❌ メール送信中に予期せぬエラーが発生しました: {e}")
        return False

def prepare_and_send_order_email(account_key: str, sender_creds: Dict[str, Any], items: List[Dict[str, Any]], pdf_path: str, selected_department: Optional[str] = None) -> bool:
    """
    UIからの情報をもとにメール送信の準備と実行を行う
    """
    sender_email = sender_creds.get("sender")
    display_name = sender_creds.get("display_name", account_key)

    # keyringからパスワードを取得
    password = keyring.get_password(SERVICE_NAME, sender_email)
    if not password:
        print(f"❌ パスワード取得エラー: {sender_email} のパスワードがOSに保存されていません。")
        print("   -> [設定]画面からアカウントを一度開き、パスワードを再保存してください。")
        return False
    
    # send_smtp_mailに渡すために、sender_credsにパスワードを一時的に追加
    creds_with_pass = sender_creds.copy()
    creds_with_pass["password"] = password

    if not items:
        print("❌ メール送信エラー: 対象アイテムがありません。")
        return False

    # メール送信を実行
    success = send_smtp_mail(
        info=items[0], 
        pdf_path=pdf_path, 
        sender_creds=creds_with_pass, 
        account_name=display_name, 
        selected_department=selected_department
    )
    
    return success