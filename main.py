import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import math
import configparser
from pathlib import Path
import ssl

# 讀取配置文件
def load_config():
    config = configparser.ConfigParser()
    config_path = Path(__file__).parent / 'config.ini' # 設定檔案路徑

    if not config_path.exists():
        raise FileNotFoundError(f"配置文件不存在: {config_path}")

    config.read(config_path)
    return {
        'server': config['SMTP']['server'],
        'port': config['SMTP'].getint('port'),
        'username': config['SMTP']['username'],
        'password': config['SMTP']['password'],
        'sender_email': config['SMTP']['sender_email'],
        'use_tls': config['SMTP'].getboolean('use_tls')
    }

# Load SMTP settings
try:
    smtp_config = load_config()
except Exception as e:
    print(f"讀取配置文件失敗: {str(e)}")
    exit(1)

# Function to send email via SMTP
def send_email(subject, body, to_address):
    try:
        # 建立郵件
        msg = MIMEMultipart()
        msg['From'] = smtp_config['sender_email']
        msg['To'] = to_address
        msg['Subject'] = subject

        # 加入內文
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        # 建立自定義 SSL context
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        context.set_ciphers('DEFAULT@SECLEVEL=1')
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE

        # 嘗試建立連接
        try:
            # 首先嘗試使用 SMTP_SSL
            with smtplib.SMTP_SSL(smtp_config['server'], smtp_config['port'], context=context) as server:
                server.login(smtp_config['username'], smtp_config['password'])
                server.send_message(msg)
        except:
            # 如果 SMTP_SSL 失敗，嘗試普通 SMTP 並手動啟動 TLS
            with smtplib.SMTP(smtp_config['server'], smtp_config['port']) as server:
                server.starttls(context=context)
                server.login(smtp_config['username'], smtp_config['password'])
                server.send_message(msg)

        print(f"郵件寄送成功: {to_address}")
        return True
    except Exception as e:
        print(f"郵件寄送失敗: {to_address}, 原因: {str(e)}")
        return False

# Load the data
try:
    file_path = 'data/your_file.xlsx'
    data = pd.read_excel(file_path)
except Exception as e:
    print(f"讀取 Excel 檔案失敗: {str(e)}")
    exit(1)

# 記錄發送結果
success_count = 0
fail_count = 0
failed_emails = []

# Send emails to all students
for index, row in data.iterrows():
    student_id = '000000'
    if pd.isna(student_id):
        print("檢測到 NaN，停止發送信件")
        break

    try:
        student_name = '學生A'
        email = 'test@example.com'

        subject = f'{student_name}同學 測試課程成績'
        body = (
            f'親愛的{student_name}同學,\n'
            f'以下是您的測試課程成績\n'
            'nan 表示缺少成績\n'
            '\n'
            f'作業一: {row.get("作業一", "N/A")}\n'
            f'作業二: {row.get("作業二", "N/A")}\n'
            f'作業三: {row.get("作業三", "N/A")}\n'
            f'期中考: {row.get("期中考", "N/A")}\n'
            '\n'
            '如有任何問題，請於期限前回覆本信件\n'
            '信件由自動發送系統寄出\n'
            '\n'
            '\n'
            'Best regards,\n'
            '自動發送系統'
        )

        if send_email(subject, body, email):
            success_count += 1
        else:
            fail_count += 1
            failed_emails.append(email)

    except Exception as e:
        print(f"處理學生資料時發生錯誤: {str(e)}")
        fail_count += 1
        failed_emails.append(f"學號: {student_id}")

# 輸出發送統計
print("\n發送統計:")
print(f"成功發送: {success_count}")
print(f"失敗發送: {fail_count}")
if failed_emails:
    print("\n失敗清單:")
    for email in failed_emails:
        print(email)
