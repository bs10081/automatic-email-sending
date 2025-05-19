# 自動化課程證書寄送系統

一個用於自動化寄送課程證書（或成績單）給學員的系統，提供 Python 與 Java 腳本/應用程式版本。支援批量寄送、測試模式、證書自動附加、資料驗證與詳細發送統計。

## 主要功能
- 從 Excel 讀取學員聯絡資料
- 根據學員姓名自動附加對應 PDF 證書
- 支援測試模式（所有信件內容都寄到 config.ini 的測試信箱）
- 支援 SSL/TLS 郵件傳送
- 詳細錯誤處理與發送統計
- SMTP 設定可於設定檔中靈活調整
- 支援多課程、多資料夾切換（需修改程式碼中的常數或 config.ini）

## 需求套件

### Python 版本
- Python 3.6+
- pandas >= 2.1.0
- openpyxl >= 3.1.2
- configparser >= 6.0.0

### Java 版本
- Java Development Kit (JDK) 11+ (或您的專案適用版本)
- Maven 或 Gradle (用於專案建置與依賴管理)
- 主要函式庫 (請根據您的 Java 專案填寫，例如):
    - Apache POI (用於讀取 Excel 檔案)
    - JavaMail API (用於郵件寄送)
    - (其他相關函式庫，例如日誌框架、設定檔讀取函式庫等)

## 安裝方式

### Python 版本
1. git clone 本專案
```bash
git clone https://github.com/bs10081/automatic-email-sending/tree/main
cd automatic-email-sending
```
2. 安裝依賴
```bash
pip install -r requirements.txt
```

### Java 版本
1. git clone 本專案
```bash
git clone https://github.com/bs10081/automatic-email-sending/tree/main
cd automatic-email-sending
```
2. 使用 Maven 或 Gradle 建置專案：
   ```bash
   # 如果使用 Maven
   mvn clean compile
   # 如果使用 Gradle （未經測試）
   gradle build
   ```

## 設定檔(`config.ini`)
請在 Python 及 Java 腳本目錄下建立 `config.ini`：
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
autoSentMail/                 # 專案根目錄
├── main.py                 # Python 主要執行腳本
├── test.py                 # Python 測試相關腳本（現已棄用，功能合併至 main.py）
├── pom.xml                 # Java Maven 專案設定檔
├── src/                    # Java 原始碼目錄
│   └── main/
│       └── java/           # Java 主要程式碼
│           └── ...         # (您的 Java 套件與類別)
├── target/                 # Java Maven 建置輸出目錄
│   └── classes/            # 建置產物
│       └── ...             # 建置產物
├── config.ini              # 通用設定檔 (Python 與 Java 共用，需自行建立)
├── requirements.txt        # Python 依賴套件列表
├── data/                   # 共用資料目錄
│   ├── 0419 聯絡資料.xlsx
│   ├── 0419 證書/
│   │   ├── 課程證書-王小明.pdf
│   │   └── ...
│   └── ...
├── .gitignore
├── README.md
└── config.ini.example      # 設定檔範例
```
**注意:** Java 專案的確切結構可能因建置工具 (Maven/Gradle) 而異。上述結構建議將 Python 和 Java 專案放在同一工作區下的不同目錄中，共用 `data` 目錄。

## 資料檔案格式 (共通)
聯絡資料 Excel 檔案需包含下列欄位：
- 姓名
- 電子郵件

證書 PDF 檔案需以「${課程證書}-${學員姓名}.pdf」(或其他約定格式，需與程式邏輯一致) 命名，並放在對應資料夾下。

## 使用方式

### Python 版本
1. 準備好聯絡資料 Excel 及證書 PDF 檔案，放入 `data/` 目錄下。
2. 修改 `Python/main.py` 中的課程名稱、聯絡資料與證書目錄常數，例如：
```python
COURSE_NAME = "2025 AI 實戰課程"
CERTIFICATE_DIR = Path("../data/0419 證書") # 相對於 main.py 的路徑
CONTACT_FILE = Path("../data/0419 聯絡資料.xlsx") # 相對於 main.py 的路徑
```
3. 執行 Python 腳本 (於 `Python/` 目錄下)：
```bash
python main.py
```

### Java 版本
1. 準備好聯絡資料 Excel 及證書 PDF 檔案，放入 `data/` 目錄下。
3. 修改 Java 原始碼中的常數 (例如課程名稱、資料夾路徑)。
4. 執行 Java 應用程式 (通常在 Java 專案的根目錄或 target 目錄執行)：
    ```java
    mvn exec:java
    ```

## 測試模式
- **Python 版本:** 將 `Python/config.ini` 的 `[TEST] enable_test_mode` 設為 `True`，並指定 `recipient_email`。腳本會遍歷所有聯絡資料，將每一筆內容（學員姓名、證書）都寄到同一個測試信箱，方便驗證所有功能。
- **Java 版本:** 請參考 Java 專案的設定檔 (例如 `config.properties` 中的 `test.enable.mode`) 或其說明文件來啟用測試模式，並指定測試收件信箱。
- 設為 `False` (或對應設定) 則會依照聯絡資料逐一寄送給每位學員。

## 執行結果範例

### Python 版本
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

### Java 版本
(請根據 Java 應用程式的實際輸出提供範例，可能類似於 Python 版本的統計資訊)
```
[INFO] Test mode enabled. Sending all emails to: test@example.com
[INFO] Processing record for: 測試學員 (Row 2 from Excel)
[INFO] Email successfully sent to test@example.com (Original recipient: real_student@example.com)
...
[INFO] ============================== Sending Statistics ==============================
[INFO] Successfully sent: 10
[INFO] Failed to send: 0
[INFO] Skipped records (incomplete data or missing certificate): 1
[INFO] --- Email sending process finished ---
```
(上述 Java 輸出為示意，請依實際情況調整)

## 常見問題 (共通)
- SMTP 連線錯誤：請檢查 SMTP 設定 (伺服器、埠號、帳號密碼) 與 SSL/TLS 設定。
- Excel 檔案找不到：請確認檔案路徑與檔名是否正確，以及程式是否有讀取權限。
- 學員資料或證書缺失：腳本/程式應能自動跳過並記錄於統計中。
- 郵件被視為垃圾郵件：檢查寄件者信譽、郵件內容、SPF/DKIM/DMARC 設定。

## 最佳實踐 (共通)
- 先以測試模式驗證所有流程與設定。
- `config.ini` 等包含敏感資訊的設定檔請勿上傳至公開程式碼倉庫 (應加入 `.gitignore`)。
- 定期備份學員資料與證書。
- 執行完畢請檢查統計結果與錯誤日誌。

## 授權
MIT License