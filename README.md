# Epaper-Automation-GAS
情資收集_cert電子報  
CERT 電子報自動收集與轉發系統 (Google Apps Script)

---

## 📌 專案簡介
此專案是一個使用 **Google Apps Script (GAS)** 開發的自動化腳本，專為 **Google Workspace 用戶** 設計，用於每月自動執行以下任務：

- **抓取 (Fetch)**：自動計算並抓取「台灣電腦網路危機處理暨協調中心 (TWCERT/CC)」發佈的上個月電子報內容（https://epaper.twcert.org.tw）。
- **歸檔 (Archive)**：將電子報 HTML 內容自動儲存至指定的 Google Drive 資料夾。
- **轉發 (Forward)**：以 HTML 郵件形式發送給指定收件人或群組信箱。
- **錯誤通知 (Error Handling)**：若抓取失敗，會自動發送錯誤通知郵件。

此解決方案利用 **Google Apps Script 的排程功能**，完全運行在 Google 環境中，無需額外伺服器。

---

## 🚀 核心功能
- **全自動排程**：透過 `setupTrigger()` 函式，設定每月固定日期自動執行。
- **相對路徑修正**：將內容中的圖片或連結改為絕對路徑，確保顯示正常。
- **雙重備份**：郵件即時通知 + Drive 歸檔。
- **跨年處理**：正確處理 1 月執行時需抓取前一年 12 月份的情境。

---

## ⚙️ 部署步驟 (Google Workspace 用戶)

### 步驟 1: 建立 Apps Script 專案
1. 開啟 Google Drive，或任一 Google Sheet / Docs。  
2. 選擇 **擴充功能 (Extensions) → Apps Script**。  
3. 將本專案的 `Code.gs` 貼到編輯器，並刪除預設內容。  
4. 將專案名稱改為 `Epaper Auto Fetcher` (或任意名稱)。  

---

### 步驟 2: 修改【設定區域】變數

| 變數名稱       | 說明 | 範例值 |
|----------------|------|--------|
| `baseDomain`   | 電子報的根網址，通常無需修改 | `"https://epaper.twcert.org.tw"` |
| `recipientEmail` | 接收電子報與錯誤通知的信箱，支援多個 (逗號分隔) | `"mail@yourorg.com, archiver@yourorg.com"` |
| `driveFolderId` | 【重要】存放 HTML 歸檔檔案的 Google Drive 資料夾 ID | `"drive.google.com/drive/folders/{your-folder-id}"` |

---

### 步驟 3: 執行並授權 (僅需一次)
1. 在函式下拉選單中選擇 `setupTrigger`。  
2. 點擊 **執行 (Run)**。  
3. 完成 Google 帳號授權，允許存取 Gmail 與 Google Drive。  
4. 成功後會建立 **每月 1 號中午 12 點自動執行**的觸發器。  

---

### 步驟 4: 測試執行 (可選)
1. 在函式下拉選單中選擇 `testMain`。  
2. 點擊 **執行 (Run)**。  
3. 檢查收件匣與指定的 Google Drive 資料夾，確認測試郵件與歸檔檔案是否正確。  

---

## 📜 程式碼詳解

| 函式名稱             | 說明 | 觸發方式 |
|----------------------|------|----------|
| `main(testDate)`     | 主執行函式，負責抓取、歸檔與發送流程 | 自動 (每月 1 號) 或手動 |
| `setupTrigger()`     | 建立觸發器的函式，僅需執行一次 | 手動執行 |
| `getPreviousMonthPath()` | 計算上個月的 URL 路徑 (例：`2025_10/`) | `main()` 內部呼叫 |
| `fetchAndExtractContent()` | 抓取網頁 `<body>` 內容，並修正相對路徑 | `main()` 內部呼叫 |
| `saveToDrive()`      | 將內容儲存為 HTML 至指定的 Drive 資料夾 | `main()` 內部呼叫 |
| `testMain()`         | 測試與模擬不同日期的邏輯，並發送測試郵件 | 手動執行 |

---

## 🔑 授權注意事項
此腳本需存取以下 Google 服務，請確認 Workspace 帳戶允許：

- **UrlFetchApp**：訪問 https://epaper.twcert.org.tw  
- **MailApp**：發送電子報與錯誤通知  
- **DriveApp**：建立與儲存 HTML 檔案  

---
