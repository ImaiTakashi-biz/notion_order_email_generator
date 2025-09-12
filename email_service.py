import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import config

def send_smtp_mail(info, pdf_path, sender_creds, account_name):
    """
    SMTPサーバー経由でPDF添付メールを送信する
    """
    try:
        msg = MIMEMultipart()
        msg["From"] = sender_creds["sender"]
        msg["To"] = info["email"]
        if info["email_cc"]:
            msg["Cc"] = info["email_cc"]
        msg["Subject"] = "注文書送付の件"

        # メール本文 (署名を動的に生成)
        body = f"""{info['supplier_name']}
{info.get('sales_contact', 'ご担当者')} 様

いつも大変お世話になります。
添付の通り注文宜しくお願い致します。

∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝
株式会社　新井精密
発注担当： {account_name}
〒368-0061
埼玉県秩父市小柱670番地
Email: {sender_creds["sender"]}
TEL: 0494-26-7786
FAX: 0494-26-7787
∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝"""
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
            
        print(f"-> メール送信完了 (To: {info['email']})")
        return True
    except Exception as e:
        print(f"メール送信エラー: {e}")
        return False
