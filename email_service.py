import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import config

def send_smtp_mail(info, pdf_path, sender_creds):
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

        # メール本文
        body = f'''{info["supplier_name"]}\n{info.get('sales_contact', 'ご担当者')} 様\n\nいつも大変お世話になります。\n添付の通り注文宜しくお願い致します。\n\n∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝\n株式会社　新井精密\n製造課　発注担当\n〒368-0061\n埼玉県秩父市小柱670番地\nTEL: 0494-26-7786\nFAX: 0494-26-7787\n∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝'''
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
