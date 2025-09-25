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
        template = config.AppConstants.EMAIL_TEMPLATE
        company = config.AppConstants.COMPANY_INFO
        
        msg["Subject"] = template['subject']

        # メール本文 (署名を動的に生成)
        body = f"""{info['supplier_name']}
{info.get('sales_contact', 'ご担当者')} 様

{template['greeting']}
{template['body']}

∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝∝
{company['name']}
発注担当： {account_name}
{company['postal_code']}
{company['address']}
Email: {sender_creds["sender"]}
{company['tel']}
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
    except Exception as e:
        return False
