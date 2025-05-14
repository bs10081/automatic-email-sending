# 自動化課程證書寄送系統

一個用於自動化寄送課程證書（或成績單）給學員的 Python 腳本，支援批量寄送、測試模式、證書自動附加、資料驗證與詳細發送統計。

## 主要功能
- 從 Excel 讀取學員聯絡資料
- 根據學員姓名自動附加對應 PDF 證書
- 支援測試模式（所有信件內容都寄到 config.ini 的測試信箱）
- 支援 SSL/TLS 郵件傳送
- 詳細錯誤處理與發送統計
- SMTP 設定可於 config.ini 靈活調整
- 支援多課程、多資料夾切換（只需修改 main.py 常數）

## 需求套件
- Python 3.6+
- pandas >= 2.1.0
- openpyxl >= 3.1.2
- configparser >= 6.0.0

## 安裝方式
1. 下載本專案
```bash
git clone https://github.com/bs10081/autoSentMail.git
cd autoSentMail
```
2. 安裝依賴
```bash
pip install -r requirements.txt
```

## 設定說明
請在腳本目錄下建立 `config.ini`：
```ini
[SMTP]
server = smtp.example.com
port = 465
username = your_username
password = your_password
sender_email = sender@example.com
use_tls = True

[TEST]
recipient_email = test@example.com
# recipient_name 不再需要，測試模式下會自動遍歷所有學員
enable_test_mode = True
```

## 檔案結構範例
```
autoSentMail/
├── main.py
├── config.ini
├── requirements.txt
├── README.md
└── data/
    ├── 0419 聯絡資料.xlsx
    ├── 0419 證書/
    │   ├── 課程證書-王小明.pdf
    │   └── ...
    └── ...
```

## 資料檔案格式
聯絡資料 Excel 檔案需包含下列欄位：
- 姓名
- 電子郵件

證書 PDF 檔案需以「課程證書-學員姓名.pdf」命名，並放在對應資料夾下。

## 使用方式
1. 準備好聯絡資料 Excel 及證書 PDF 檔案，放入 `data/` 目錄下。
2. 修改 `main.py` 中的課程名稱、聯絡資料與證書目錄常數，例如：
```python
COURSE_NAME = "2025 AI 實戰課程"
CERTIFICATE_DIR = Path("data/0419 證書")
CONTACT_FILE = Path("data/0419 聯絡資料.xlsx")
```
3. 執行腳本：
```bash
python main.py
```

## 測試模式
- 將 `config.ini` 的 `[TEST] enable_test_mode` 設為 `True`，並指定 `recipient_email`。
- 腳本會遍歷所有聯絡資料，將每一筆內容（學員姓名、證書）都寄到同一個測試信箱，方便驗證所有功能。
- 設為 `False` 則會依照聯絡資料逐一寄送給每位學員。

## 執行結果範例
```
[測試模式] 準備發送郵件給 (Excel 第 2 行): 測試學員 (實際寄送至 test@example.com) ...
郵件成功寄送至: test@example.com
...

============================== 發送統計 ==============================
成功發送: 10
失敗發送: 0
略過記錄 (資料不完整或無證書): 1

--- 郵件發送處理結束 ---
```

## 常見問題
- SMTP 連線錯誤：請檢查 SMTP 設定與 SSL 憑證
- Excel 檔案找不到：請確認檔案路徑與檔名
- 學員資料或證書缺失：腳本會自動跳過並記錄於統計

## 最佳實踐
- 先以測試模式驗證所有流程
- config.ini 請勿上傳至公開倉庫
- 定期備份學員資料與證書
- 執行完畢請檢查統計結果

## 授權
MIT License