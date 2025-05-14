import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import math
import configparser
from pathlib import Path
import ssl
import time

# 讀取配置文件
def load_config():
    config = configparser.ConfigParser()
    config_path = Path(__file__).parent / 'config.ini'

    if not config_path.exists():
        # 如果 config.ini 不存在，創建一個包含預設值的
        config['SMTP'] = {
            'server': 'smtp.example.com',
            'port': '465',
            'username': 'your_username',
            'password': 'your_password',
            'sender_email': 'sender@example.com',
            'use_tls': 'True'
        }
        config['TEST'] = {
            'recipient_name': '測試姓名', # 用於測試模式下查找證書和郵件稱呼
            'recipient_email': 'test_recipient@example.com', # 測試郵件的接收地址
            'enable_test_mode': 'False' # 設為 True 則只發送測試郵件
        }
        with open(config_path, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
        print(f"配置文件 '{config_path}' 不存在，已創建預設配置。請修改後再運行。")
        exit(1) # 首次創建後退出，讓用戶修改

    config.read(config_path, encoding='utf-8')
    
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
            'recipient_name_config': config['TEST'].get('recipient_name', ''),
            'recipient_email_config': config['TEST'].get('recipient_email', ''),
            'enable_test_mode': config['TEST'].getboolean('enable_test_mode', False) # 預設為 False
        }
    else: # 如果 TEST 區段不存在，也提供預設值
        test_settings = {
            'recipient_name_config': '測試學員',
            'recipient_email_config': 'test@example.com',
            'enable_test_mode': False
        }
            
    return smtp_settings, test_settings

# Load SMTP and Test settings
try:
    smtp_config, test_config = load_config()
except FileNotFoundError as e:
    print(e) # load_config 內部已處理 FileNotFoundError，這裡理論上不會觸發
    exit(1)
except Exception as e:
    print(f"讀取配置文件時發生未預期錯誤: {str(e)}")
    exit(1)

# Function to send email with attachment
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

        # 嘗試 SMTP_SSL 或 SMTP with STARTTLS
        server = None
        try:
            server = smtplib.SMTP_SSL(smtp_config['server'], smtp_config['port'], context=context)
            server.login(smtp_config['username'], smtp_config['password'])
            server.send_message(msg)
        except Exception as e_ssl:
            print(f"SMTP_SSL 連接失敗: {e_ssl}. 嘗試使用 STARTTLS...")
            if server: server.quit()
            try:
                server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
                server.starttls(context=context)
                server.login(smtp_config['username'], smtp_config['password'])
                server.send_message(msg)
            except Exception as e_starttls:
                print(f"郵件寄送失敗 (SMTP_SSL 和 STARTTLS 皆失敗): {to_address}, 原因: {e_starttls}")
                if server: server.quit()
                return False
        finally:
            if server: server.quit()

        print(f"郵件成功寄送至: {to_address}")
        return True
    except Exception as e_outer:
        print(f"準備郵件或整體發送過程中發生錯誤: {to_address}, 原因: {str(e_outer)}")
        return False

# 設定常數
COURSE_NAME = "2025 未來造浪 AI Studio"
CERTIFICATE_DIR = Path("data/0419 證書")
CONTACT_FILE = Path("data/0419 聯絡資料.xlsx")

# 載入聯絡資料 (僅在非測試模式或測試模式設定無效時才需要完整載入)
try:
    if not CONTACT_FILE.exists():
        print(f"錯誤: 聯絡資料檔案 '{CONTACT_FILE}' 不存在。無法執行。")
        exit(1)
    contacts_df = pd.read_excel(CONTACT_FILE)
    print(f"成功讀取聯絡資料，共 {len(contacts_df)} 筆記錄。")
except Exception as e:
    print(f"讀取聯絡資料檔案 '{CONTACT_FILE}' 失敗: {str(e)}")
    exit(1)

# 獲取所有證書文件並建立名稱到路徑的映射
name_to_certificate = {}
try:
    if not CERTIFICATE_DIR.exists() or not CERTIFICATE_DIR.is_dir():
        print(f"錯誤: 證書目錄 '{CERTIFICATE_DIR}' 不存在或不是一個目錄。")
        exit(1)
    for cert_file_path in CERTIFICATE_DIR.glob('*.pdf'):
        # 從檔名取得姓名部分，假設格式為 "課程名稱證書-姓名.pdf" 或 "證書-姓名.pdf"
        file_stem = cert_file_path.stem # 檔名不含副檔名
        name_part = file_stem.split('-')[-1].strip() #取最後一部分並去除空白
        if name_part:
            name_to_certificate[name_part] = cert_file_path
    print(f"找到 {len(name_to_certificate)} 個證書檔案並已建立映射。")
    if not name_to_certificate:
        print(f"警告: 在證書目錄 '{CERTIFICATE_DIR}' 中未找到任何 PDF 證書檔案。")
except Exception as e:
    print(f"讀取證書目錄 '{CERTIFICATE_DIR}' 或處理證書檔案時發生錯誤: {str(e)}")
    exit(1)

# 記錄發送結果
success_count = 0
fail_count = 0
skipped_count = 0
failed_recipients_info = [] # 改為儲存更詳細的失敗資訊

print("\n--- 開始郵件發送處理 ---")

# --- 測試模式優先邏輯 ---
if test_config.get('enable_test_mode', False):
    print("*** 測試模式已啟用 (來自 config.ini) ***")
    test_recipient_email = test_config.get('recipient_email_config', '').strip()

    # 驗證測試模式的收件人信箱是否有效
    if not test_recipient_email:
        print("="*70)
        print("********** CLI 錯誤: 測試模式設定不完整 **********")
        print("  測試模式已啟用，但 config.ini 中的 [TEST] recipient_email 未提供或為空。")
        print(f"  > 設定的測試 Email: '{test_recipient_email}'")
        print("  請在 config.ini 中提供完整的測試收件人 Email 地址。")
        print("="*70 + "\n")
        print("--- 郵件發送處理結束 (因測試模式配置無效) ---")
        exit(1)

    print(f"將遍歷所有聯絡資料，並將所有郵件內容寄送到測試信箱: {test_recipient_email}")
    if contacts_df.empty:
        print("聯絡資料表格為空，沒有可發送的郵件。")
    
    for index, row in contacts_df.iterrows():
        current_row_num = index + 2
        try:
            recipient_name = str(row.get('姓名', '')).strip()
            recipient_email = test_recipient_email  # 強制所有信件都寄到測試信箱
            original_email = str(row.get('電子郵件', '')).strip()
            
            if not recipient_name or not original_email:
                print(f"警告 (Excel 第 {current_row_num} 行): 姓名 ('{recipient_name}') 或原始 Email ('{original_email}') 為空，跳過此記錄。")
                skipped_count += 1
                continue
            
            certificate_path = name_to_certificate.get(recipient_name)
            if not certificate_path:
                print(f"警告 (Excel 第 {current_row_num} 行，學員: {recipient_name}): 找不到對應的證書檔案，跳過此記錄。")
                skipped_count += 1
                failed_recipients_info.append(f"Excel 行 {current_row_num}: {recipient_name} <{original_email}> - 原因: 無證書")
                continue
            
            subject = f"「{COURSE_NAME}」課程證書寄發通知｜感謝您的參與！ (測試模式)"
            body = (
                f"{recipient_name} 同學，您好： (此為測試模式郵件，實際寄送至 {test_recipient_email})\n\n"
                f"感謝您參加「{COURSE_NAME}」課程，我們很高興與您一同探索 AI 的應用，見證您的學習成長與成果！\n\n"
                "您已順利完成本次課程，並依規定完成所有作品繳交，寄發電子課程證書，以茲證明。\n\n"
                "如您發現證書內容有誤或無法順利下載，請於 7 日內回信通知，我們將協助您更正或補發。\n\n"
                "再次感謝您的投入與參與，我們期待未來與您在更多課程中再次相見，共同開啟更多 AI 學習與實作的可能！\n\n"
                "敬祝 學習順利！\n\n"
                "自主學習與資訊專業成長教學團隊\n\n"
                "📧 聯絡信箱：ncnu.webcamping@gmail.com"
            )
            
            print(f"\n[測試模式] 準備發送郵件給 (Excel 第 {current_row_num} 行): {recipient_name} (實際寄送至 {test_recipient_email}) ...")
            
            if send_email_with_attachment(subject, body, recipient_email, certificate_path):
                success_count += 1
            else:
                fail_count += 1
                failed_recipients_info.append(f"Excel 行 {current_row_num}: {recipient_name} <{original_email}> - 原因: 發送失敗")
            
            if success_count + fail_count < len(contacts_df) - skipped_count:
                print(f"等待 {2} 秒後發送下一封...")
                time.sleep(2)
        except Exception as e_loop:
            print(f"處理 Excel 第 {current_row_num} 行 (學員: '{row.get('姓名', '未知')}') 時發生未預期錯誤: {str(e_loop)}")
            fail_count += 1
            failed_recipients_info.append(f"Excel 行 {current_row_num}: {row.get('姓名', '未知')} <{row.get('電子郵件', '未知')}> - 原因: 迴圈中發生錯誤")
    print("--- 測試模式郵件發送完成 --- ")

else:
    # --- 正常批量發送模式 ---
    print("*** 正常批量發送模式已啟用 (測試模式未啟用或配置無效) ***")
    if contacts_df.empty:
        print("聯絡資料表格為空，沒有可發送的郵件。")
    
    for index, row in contacts_df.iterrows():
        current_row_num = index + 2 # Excel 行號通常從 1 開始，標頭佔 1 行
        try:
            recipient_name = str(row.get('姓名', '')).strip()
            recipient_email = str(row.get('電子郵件', '')).strip()
            
            if not recipient_name or not recipient_email:
                print(f"警告 (Excel 第 {current_row_num} 行): 姓名 ('{recipient_name}') 或 Email ('{recipient_email}') 為空，跳過此記錄。")
                skipped_count += 1
                continue
            
            certificate_path = name_to_certificate.get(recipient_name)
            if not certificate_path:
                print(f"警告 (Excel 第 {current_row_num} 行，學員: {recipient_name}): 找不到對應的證書檔案，跳過此記錄。")
                skipped_count += 1
                failed_recipients_info.append(f"Excel 行 {current_row_num}: {recipient_name} <{recipient_email}> - 原因: 無證書")
                continue
            
            subject = f"「{COURSE_NAME}」課程證書寄發通知｜感謝您的參與！"
            body = (
                f"{recipient_name} 同學，您好：\n\n"
                f"感謝您參加「{COURSE_NAME}」課程，我們很高興與您一同探索 AI 的應用，見證您的學習成長與成果！\n\n"
                "您已順利完成本次課程，並依規定完成所有作品繳交，寄發電子課程證書，以茲證明。\n\n"
                "如您發現證書內容有誤或無法順利下載，請於 7 日內回信通知，我們將協助您更正或補發。\n\n"
                "再次感謝您的投入與參與，我們期待未來與您在更多課程中再次相見，共同開啟更多 AI 學習與實作的可能！\n\n"
                "敬祝 學習順利！\n\n"
                "自主學習與資訊專業成長教學團隊\n\n"
                "📧 聯絡信箱：ncnu.webcamping@gmail.com"
            )
            
            print(f"\n準備發送郵件給 (Excel 第 {current_row_num} 行): {recipient_name} <{recipient_email}> ...")
            
            if send_email_with_attachment(subject, body, recipient_email, certificate_path):
                success_count += 1
            else:
                fail_count += 1
                failed_recipients_info.append(f"Excel 行 {current_row_num}: {recipient_name} <{recipient_email}> - 原因: 發送失敗")
            
            if success_count + fail_count < len(contacts_df) - skipped_count:
                 print(f"等待 {2} 秒後發送下一封...")
                 time.sleep(2)
        except Exception as e_loop:
            print(f"處理 Excel 第 {current_row_num} 行 (學員: '{row.get('姓名', '未知')}') 時發生未預期錯誤: {str(e_loop)}")
            fail_count += 1
            failed_recipients_info.append(f"Excel 行 {current_row_num}: {row.get('姓名', '未知')} <{row.get('電子郵件', '未知')}> - 原因: 迴圈中發生錯誤")

# --- 輸出發送統計 ---
print("\n" + "="*30 + " 發送統計 " + "="*30)
print(f"成功發送: {success_count}")
print(f"失敗發送: {fail_count}")
print(f"略過記錄 (資料不完整或無證書): {skipped_count}")

if failed_recipients_info:
    print("\n--- 失敗或部分成功記錄詳情 ---")
    for info in failed_recipients_info:
        print(f"  - {info}")
print("="*70)
print("--- 郵件發送處理結束 ---")
