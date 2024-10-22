import imaplib
import smtplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
import time
import threading
import requests
import json
import re
import yaml

# 从配置文件导入参数
def load_config(config_file):
    with open(config_file, 'r') as file:
        return yaml.safe_load(file)

config = load_config('config.yaml')

# QQ邮箱IMAP和SMTP服务器配置
IMAP_SERVER = config['imap_server']
SMTP_SERVER = config['smtp_server']
IMAP_PORT = config['imap_port']
SMTP_PORT = config['smtp_port']

# QQ邮箱账号和授权码（类似于密码，用于登录）
EMAIL_ACCOUNT = config['email_account']
EMAIL_PASSWORD = config['email_password']

# 模型API密钥和URL
API_URL = config['api_url']
API_KEY = config['api_key']

# 其他邮箱设置
CC_EMAILS = config['cc_emails']

def connect_to_imap():
    # 连接到IMAP服务器
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    return mail

def fetch_unread_emails(mail):
    mail.select("inbox")
    status, messages = mail.search(None, 'UNSEEN')
    email_ids = messages[0].split()
    return email_ids

def generate_reply(content):
    # 使用模型接口生成回复内容
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "THUDM/glm-4-9b-chat",
        "messages": [
            {"role": "system", "content": "Your task is to understand the content of an email and categorize the issue. Determine which department should address this issue based on the following options:\n0.  产品部门\n1. 销售部门\n2. 开发部门\n3. 市场部门\n4. 其他部门\nOutput format should be a single number representing the option, without any additional output."},
            {"role": "user", "content": content}
        ],
        "temperature": 0.7,
        "top_p": 0.7,
        "max_tokens": 200
    }
    response = requests.post(API_URL, headers=headers, data=json.dumps(payload))
    if response.status_code == 200:
        reply = response.json()["choices"][0]["message"]["content"]
        match = re.search(r'\d+', reply)
        if match:
            print(f"Attempting to forward to department {int(match.group())}")
            return int(match.group())
    else:
        print(f"Failed to generate reply: {response.text}")
    return 3  # 默认值为3

def send_reply(to_address, subject, body, cc_key=None):
    # 创建回复邮件
    msg = MIMEText(body)
    msg["From"] = EMAIL_ACCOUNT
    msg["To"] = to_address
    msg["Subject"] = "Re: " + subject
    if cc_key is not None:
        msg["Cc"] = CC_EMAILS.get(str(cc_key), "")

    # 发送邮件
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
            cc_address = [CC_EMAILS.get(str(cc_key), "")] if cc_key is not None else []
            server.sendmail(EMAIL_ACCOUNT, [to_address] + cc_address, msg.as_string())
            print(f"Sent reply to {to_address} (CC: {', '.join(cc_address)})")
    except smtplib.SMTPException as e:
        print(f"Failed to send email to {to_address}: {e}")

def process_emails(mail):
    email_ids = fetch_unread_emails(mail)
    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding or 'utf-8')
                
                from_address = email.utils.parseaddr(msg["From"])[1]
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            body = part.get_payload(decode=True).decode()
                            break
                else:
                    body = msg.get_payload(decode=True).decode()

                print(f"Processing email from {from_address} with subject: {subject}")
                # 固定回复内容
                reply_body = "I have received your email and will address your issue promptly. Thank you. This is an automated response."
                send_reply(from_address, subject, reply_body)

                # 将邮件内容抄送给指定的部门
                cc_key = generate_reply(body)
                forward_subject = "Fwd: " + subject
                forward_body = f"Forwarded message:\n\n{body}"
                if str(cc_key) in CC_EMAILS:
                    send_reply(CC_EMAILS[str(cc_key)], forward_subject, forward_body, cc_key)
                    print(f"Successfully forwarded to department {cc_key}")
                # 标记为已读
                mail.store(email_id, '+FLAGS', '\Seen')

def auto_reply():
    try:
        mail = connect_to_imap()
        while True:
            process_emails(mail)
            print("Checking for new emails...")
            time.sleep(30)  # 每30秒检查一次新邮件
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        mail.logout()

def run_in_background():
    threading.Thread(target=auto_reply, daemon=True).start()
    while True:
        time.sleep(30)  # 主线程保持运行，每30秒打印一次检查新邮件的消息
        print("Checking for new emails...")

if __name__ == '__main__':
    run_in_background()
