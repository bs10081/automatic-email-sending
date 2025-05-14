import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import configparser
from pathlib import Path
import ssl

# 讀取配置文件
def load_config():
    config = configparser.ConfigParser()
    config_path = Path(__file__).parent / 'config.ini'

    if not config_path.exists():
        config['SMTP'] = {
            'server': 'smtp.example.com',
            'port': '465',
            'username': 'your_username',
            'password': 'your_password',
            'sender_email': 'sender@example.com',
            'use_tls': 'True'
        }
        
        config['TEST'] = {
            'recipient_name': '測試學員',
            'recipient_email': 'test@example.com',
            'enable_test_mode': 'True'
        }

        with open(config_path, 'w') as configfile:
            config.write(configfile)

    config.read(config_path)
    
    smtp_settings = {
        'server': config['SMTP']['server'],
        'port': config['SMTP'].getint('port'),
        'username': config['SMTP']['username'],
        'password': config['SMTP']['password'],
        'sender_email': config['SMTP']['sender_email'],
        'use_tls': config['SMTP'].getboolean('use_tls')
    }
    
    test_settings = {}
    if 'TEST' in config:
        test_settings = {
            'recipient_name_config': config['TEST'].get('recipient_name', '測試學員'),
            'recipient_email_config': config['TEST'].get('recipient_email', 'test@example.com'),
            'enable_test_mode': config['TEST'].getboolean('enable_test_mode', True)
        }
    
    return smtp_settings, test_settings

# Load SMTP and Test settings
try:
    smtp_config, test_config = load_config()
except Exception as e:
    print(f"讀取配置文件失敗: {str(e)}")
    exit(1)

# Function to send email via SMTP with attachment
def send_email_with_attachment(subject, body, to_address, attachment_path=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_config['sender_email']
        msg['To'] = to_address
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as file:
                attach = MIMEApplication(file.read(), _subtype="pdf")
                attach.add_header('Content-Disposition', 'attachment', 
                                  filename=os.path.basename(attachment_path))
                msg.attach(attach)

        context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        context.set_ciphers('DEFAULT@SECLEVEL=1')
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE

        try:
            with smtplib.SMTP_SSL(smtp_config['server'], smtp_config['port'], context=context) as server:
                server.login(smtp_config['username'], smtp_config['password'])
                server.send_message(msg)
        except:
            with smtplib.SMTP(smtp_config['server'], smtp_config['port']) as server:
                server.starttls(context=context)
                server.login(smtp_config['username'], smtp_config['password'])
                server.send_message(msg)

        print(f"郵件成功寄送至: {to_address}")
    except Exception as e:
        print(f"郵件寄送失敗: {to_address}, 原因: {str(e)}")

# --- 測試腳本主要邏輯 ---
COURSE_NAME = "2025 職場新紀元：AI 融合翻譯與設計的實戰秘笈"
CONTACT_DATA_FILE = 'data/0412 聯絡資料.xlsx'
CERTIFICATE_DIR = 'data/0412 證書'

print("--- 開始測試郵件發送腳本 ---")

# 1. 讀取聯絡資料 (用於信件內容和證書查找)
try:
    contacts_df = pd.read_excel(CONTACT_DATA_FILE)
    if contacts_df.empty:
        print(f"錯誤: 聯絡資料檔案 '{CONTACT_DATA_FILE}' 為空。無法進行測試。")
        exit(1)
    print(f"成功讀取聯絡資料，共 {len(contacts_df)} 筆記錄。")
    content_data_source_student = contacts_df.iloc[0]
    student_name_for_body_and_cert = content_data_source_student.get('姓名')
    original_student_email_from_data = content_data_source_student.get('電子郵件')

    # 緊急停止條件1: Excel 原始資料無效
    if pd.isna(student_name_for_body_and_cert) or not str(student_name_for_body_and_cert).strip():
        print("\n" + "="*70)
        print("********** CLI 錯誤: Excel 聯絡資料第一筆記錄缺少 '姓名' **********")
        print(f"  > 檔案: '{CONTACT_DATA_FILE}'")
        print(f"  > 偵測到的姓名: '{student_name_for_body_and_cert}'")
        print("  > 測試郵件的內容生成依賴此姓名，無法繼續。")
        print("="*70 + "\n")
        print("--- 測試郵件發送腳本結束 (因 Excel 原始姓名資料無效) ---")
        exit(1)
        
    if pd.isna(original_student_email_from_data) or not str(original_student_email_from_data).strip():
        print("\n" + "="*70)
        print("********** CLI 錯誤: Excel 聯絡資料第一筆記錄缺少 '電子郵件' **********")
        print(f"  > 檔案: '{CONTACT_DATA_FILE}'")
        print(f"  > 學員姓名: '{student_name_for_body_and_cert}'")
        print(f"  > 偵測到的原始 Email: '{original_student_email_from_data}'")
        print("  > 測試郵件的邏輯依賴此原始 Email (即使可能被 config.ini 覆蓋)，無法繼續。")
        print("="*70 + "\n")
        print("--- 測試郵件發送腳本結束 (因 Excel 原始 Email 資料無效) ---")
        exit(1)

except Exception as e:
    print(f"讀取聯絡資料檔案 '{CONTACT_DATA_FILE}' 失敗: {str(e)}")
    exit(1)

# 2. 決定實際的收件人 Email 地址
actual_recipient_email = original_student_email_from_data
email_recipient_source_info = f"聯絡資料中的第一位學員 ({student_name_for_body_and_cert} <{original_student_email_from_data}>)"

if test_config.get('enable_test_mode', True) and test_config.get('recipient_email_config'):
    actual_recipient_email = test_config['recipient_email_config']
    email_recipient_source_info = f"config.ini 中指定的測試信箱 <{actual_recipient_email}>"
    print(f"測試模式啟用：郵件將發送至 config.ini 指定的信箱: {actual_recipient_email}")
    print(f"郵件內容 (學員姓名、證書) 將基於聯絡資料中的: {student_name_for_body_and_cert}")
else:
    if not test_config.get('recipient_email_config') and test_config.get('enable_test_mode', True):
        print("警告: 測試模式已啟用，但 config.ini 中未指定 'recipient_email' 以覆蓋。")
    print(f"測試模式未啟用或未指定覆蓋信箱：郵件將發送至聯絡資料中的第一位學員: {actual_recipient_email}")

# 緊急停止條件2: 最終收件人 Email 地址無效
if pd.isna(actual_recipient_email) or not str(actual_recipient_email).strip():
    log_message = (
        f"警告日誌: 偵測到無效或缺失的【最終收件人】Email 地址。\n"
        f"  > 嘗試使用的 Email: '{actual_recipient_email}'\n"
        f"  > Email 來源: {email_recipient_source_info}\n"
        f"  > 因此，將跳過此次測試郵件的準備與發送。\n"
        f"  > 請檢查您的 '{CONTACT_DATA_FILE}' 檔案中的第一筆資料，或 config.ini 中的 [TEST] recipient_email 設定。"
    )
    print("\n" + "="*70)
    print("********** CLI 警告: 無效的【最終收件人】EMAIL 地址 **********")
    print(log_message)
    print("="*70 + "\n")
    print("--- 測試郵件發送腳本結束 (因最終收件人Email無效而跳過發送) ---")
    exit(0) # 以 0 退出，表示因資料問題正常跳過

# 3. 尋找對應的證書 (使用從聯絡資料中讀取的姓名)
certificate_file_path = None
try:
    for file_name in os.listdir(CERTIFICATE_DIR):
        if file_name.endswith('.pdf') and student_name_for_body_and_cert in file_name:
            certificate_file_path = os.path.join(CERTIFICATE_DIR, file_name)
            print(f"找到證書檔案: {certificate_file_path} 給 {student_name_for_body_and_cert}")
            break
    if not certificate_file_path:
        print(f"警告: 找不到 {student_name_for_body_and_cert} 的證書檔案。郵件將不含附件。")
except Exception as e:
    print(f"讀取證書目錄 '{CERTIFICATE_DIR}' 時發生錯誤: {str(e)}。郵件將不含附件。")

# 4. 準備信件內容 (使用從聯絡資料中讀取的姓名)
subject = f"「{COURSE_NAME}」課程證書寄發通知｜感謝您的參與！ (測試郵件)"
body = (
    f"{student_name_for_body_and_cert} 同學，您好：\n\n"
    f"感謝您參加「{COURSE_NAME}」課程，我們很高興與您一同探索 AI 的應用，見證您的學習成長與成果！\n\n"
    "您已順利完成本次課程，並依規定完成所有作品繳交，寄發電子課程證書，以茲證明。\n\n"
    "如您發現證書內容有誤或無法順利下載，請於 7 日內回信通知，我們將協助您更正或補發。\n\n"
    "再次感謝您的投入與參與，我們期待未來與您在更多課程中再次相見，共同開啟更多 AI 學習與實作的可能！\n\n"
    "敬祝 學習順利！\n\n"
    "自主學習與資訊專業成長教學團隊\n\n"
    "📧 聯絡信箱：ncnu.webcamping@gmail.com"
)

# 5. 發送郵件
print(f"\n準備發送測試郵件..." )
print(f"  收件人 (實際發送): {actual_recipient_email} (來源: {email_recipient_source_info})")
print(f"  原始學員信箱（來自 Excel）: {original_student_email_from_data}")
print(f"  學員姓名 (信件內容): {student_name_for_body_and_cert}")
print(f"  附件: {certificate_file_path if certificate_file_path else '無'}")

send_email_with_attachment(subject, body, actual_recipient_email, certificate_file_path)

print("--- 測試郵件發送腳本結束 ---")
