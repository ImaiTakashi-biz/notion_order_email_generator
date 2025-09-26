import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import config

def send_smtp_mail(info, pdf_path, sender_creds, account_name, selected_department=None):
    """
    SMTPサーバー経由でPDF添付メールを送信する
    """
    try:
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
        guidance_number = guidance_numbers.get(selected_department, "")
        tel_with_guidance = company['tel']
        if guidance_number:
            tel_with_guidance = f"{company['tel']} ({guidance_number})"

        # メール本文 (署名を動的に生成)
        body = f"""{info['supplier_name']}
{info.get('sales_contact', 'ご担当者')} 様

{template['greeting']}
{template['body']}

∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝
{company['name']}
発注担当： {order_contact}
{company['postal_code']}
{company['address']}
Email: {sender_creds["sender"]}
{tel_with_guidance}
{company['fax']}
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
            server.login(sender_creds["sender"], sender_creds["password"])
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
