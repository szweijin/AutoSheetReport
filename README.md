# AutoSheetReport
自動化 Google Sheets 報表產生與上傳系統

## 專案結構建議
```
AutoSheetReport/
├── app_gui.py                 # 主程式
├── report_generator.py        # 報表產生邏輯
├── upload_to_drive.py         # Google Drive 上傳模組
├── sheets_config.json         # Google Sheets 設定
├── config.json                # 其他設定檔（可排除敏感資訊）
├── credentials.json           # **這個不上傳，要用 .gitignore 忽略**
├── README.md                  # 專案說明文件
├── requirements.txt           # 依賴套件清單
└── .gitignore                 # 忽略清單（例如 credentials.json、.venv/）
```

## 功能
- 從多個 Google Sheets 抓取資料，合併後產生 Excel 報表
- 依版本號自動命名報表檔案
- 自動上傳報表至指定 Google Drive 資料夾
- 支援打包成單檔執行檔 (EXE / macOS app)

## 使用說明
1. 準備 Google 服務帳戶憑證 `credentials.json` （下方有寫準備步驟）
2. 修改 `sheets_config.json` 設定要抓取的試算表ID與部門名稱
3. 修改 `config.json` 設定 Google Drive 資料夾 ID
4. 安裝相依套件：`pip install -r requirements.txt`
5. 執行程式：`python app_gui.py`

## 注意事項
- `credentials.json` 請勿上傳至 GitHub，避免憑證外洩

## 未來優化方向
- 新增排程自動執行功能
- 報表自動寄送郵件

# GitHub 專案下載與使用指南
### 1. 下載專案: 用 Git 指令（推薦）
  ```
  git clone https://github.com/你的帳號/你的專案名稱.git
  cd 你的專案名稱
  ```
  或者直接點 GitHub 頁面右上「Code」按鈕，選「Download ZIP」，下載後解壓縮


### 2. 安裝 Python 環境
  確保電腦已安裝 Python 3（建議 3.7+）
  在終端機或命令提示字元輸入：
  ```python --version```
  如果沒有安裝，可以到 python.org 下載安裝。

### 3. 建立虛擬環境（建議）
*   避免與其他 Python 專案套件衝突 ```python -m venv venv```

#### 啟動虛擬環境： 
  * Windows: `venv\Scripts\activate`
  * macOS/Linux: `source venv/bin/activate`

### 4. 安裝相依套件
  專案裡會有 requirements.txt ，執行：
  `pip install -r requirements.txt`

### 5. 準備 Google API 憑證
* 申請並下載 Google 服務帳戶憑證 credentials.json，放在專案資料夾根目錄（或 README 說明路徑）
* 修改 sheets_config.json，填入你的 Google Sheets 試算表ID及部門名稱
* 修改 config.json，設定你的 Google Drive 資料夾 ID

### 6. 執行專案
`python app_gui.py`

### 7. 打包後使用（如果有）
如果你有用 PyInstaller 打包成執行檔，直接執行打包出來的檔案即可。

# 其他說明
1. credentials.json 請妥善保管，不要公開放在 GitHub。
2. 如果程式找不到設定檔，請確認檔案路徑和名稱是否正確。
3. 報表會產生在指定的輸出資料夾（預設 output 或你自己設定的路徑）。


# credentials.json準備與啟用 Google Sheets API
### 1.：建立 Google Sheet 串接（OAuth 服務帳戶）
#### 開通 Google Sheets API（只做一次）
* 到 Google Cloud Console 建立新專案（或選一個現有專案）
* 在左側選單點「API 和服務」→「啟用 API 與服務」
* 搜尋 `Google Sheets API`、`Google Drive API`，點進去後按「啟用」

### 2.建立服務帳戶憑證（credentials.json）
* 左側選單點「API 和服務」→「憑證」
* 點選「建立憑證」→ 選「服務帳戶」
* 取名後一路下一步（不用選角色）
* 建立完成後，點該服務帳戶 → 頁面下方橫排選單「金鑰」→ 新增鍵 → 選擇 JSON 就會自動下載
`這個檔案會是你程式用來登入 Google API 的 json檔案，請改名為 credentials.json，並放到資料夾根目錄`

### 3. 將你的 Google Sheet 分享給服務帳戶 email
你在 credentials.json 裡會看到一個 email 類似：
`xxxx@your-project.iam.gserviceaccount.com`
請打開你要讀取的 Google Sheet，右上角「共用」→ 加入這個信箱 → 權限「編輯者」

