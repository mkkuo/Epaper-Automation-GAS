# Epaper-Automation-GAS
情資收集_cert電子報
CERT 電子報自動收集與轉發系統 (Google Apps Script)

專案簡介
此專案是一個使用 Google Apps Script (GAS) 開發的自動化腳本。它專為 Google Workspace 用戶設計，用於每月自動執行以下任務：

抓取 (Fetch): 自動計算並抓取「台灣電腦網路危機處理暨協調中心 (TWCERT/CC)」發佈的上個月電子報內容（epaper.twcert.org.tw）。

歸檔 (Archive): 將抓取到的電子報 HTML 內容自動儲存至使用者指定的 Google Drive 資料夾，便於組織內部歸檔和搜尋。

轉發 (Forward): 將電子報內容以 HTML 郵件形式發送給指定的收件人或群組信箱，確保資訊即時傳遞。

錯誤通知 (Error Handling): 如果抓取失敗 (例如網址不存在或伺服器錯誤)，會自動發送錯誤通知郵件給收件人。

此解決方案利用 Google Apps Script 的排程功能，完全運行在 Google 環境中，無需額外的伺服器。

核心功能
全自動排程： 透過 setupTrigger() 函式，設定每月固定日期自動執行。

相對路徑修正： 自動將電子報內容中的圖片或連結的相對路徑修正為絕對路徑，確保郵件和 Drive 檔案中的顯示正常。

雙重備份： 郵件即時通知 + Drive 歸檔，確保情資不遺漏。

跨年處理： 正確處理每年的 1 月執行時，需要抓取前一年 12 月份的邏輯。

部署步驟 (Google Workspace 用戶)
請將 Code.gs 檔案的內容複製到您的 Google Apps Script 編輯器中。

步驟 1: 建立 Apps Script 專案
開啟您的 Google Drive，或任一 Google Sheet / Docs。

選擇 擴充功能 (Extensions) -> Apps Script。

將預設的 Code.gs 內容全部替換為本專案提供的 Code.gs 內容。

將專案名稱改為 Epaper Auto Fetcher (或您喜歡的名稱)。

步驟 2: 修改【設定區域】變數
請在程式碼的頂部，依據您的需求修改以下三個變數：

變數名稱

說明

範例值

baseDomain

電子報的根網址。通常無需修改。

"https://epaper.twcert.org.tw"

recipientEmail

接收電子報和錯誤通知的信箱。多個信箱請用逗號 , 分隔。

"mail@yourorg.com, archiver@yourorg.com"

driveFolderId

【重要】 存放歸檔 HTML 檔案的 Google Drive 資料夾 ID。

 //drive.google.com/drive/folders/"{your folder id}}"
  
步驟 3: 執行並授權 (僅需一次)
在 Apps Script 編輯器中，從頂部函式下拉選單中選擇 setupTrigger 函式。

點擊 執行 (Run) 按鈕。

Apps Script 將要求您授權此腳本訪問您的 Gmail 和 Google Drive 服務。請依指示完成授權。

授權完成後，setupTrigger 函式將會設定一個每月 1 號中午 12 點執行的自動觸發器。

步驟 4: 測試執行 (可選)
為了確認設定和授權無誤，您可以執行以下函式：

在下拉選單中選擇 testMain 函式。

點擊 執行 (Run)。

testMain 會模擬不同日期執行 main 函式，並發送測試郵件到您的收件信箱。

檢查您的收件箱和指定的 Google Drive 資料夾，確認郵件內容和歸檔檔案是否正確。

程式碼詳解
函式名稱

說明

觸發方式

main(testDate)

主執行函式。 負責整個抓取、歸檔和發送流程。

由觸發器自動執行（每月 1 號）或手動執行。

setupTrigger()

觸發器設定函式。 只需執行一次，用來建立每月的自動排程。

手動執行一次。

getPreviousMonthPath()

根據當前日期，計算上個月的 URL 路徑 (e.g., 2025_10/)。

main() 內部呼叫。

fetchAndExtractContent()

抓取網頁內容，並提取 <body> 內容，同時修正相對路徑。

main() 內部呼叫。

saveToDrive()

將內容儲存為 HTML 檔案到指定 Drive 資料夾。

main() 內部呼叫。

testMain()

包含多個模擬情境，用於測試程式碼邏輯是否正確。

手動執行，用於測試。

授權注意事項
由於此腳本需要執行以下操作，您必須授予相應的權限：

外部連線 (UrlFetchApp): 訪問 epaper.twcert.org.tw 網址。

發送郵件 (MailApp): 發送電子報和錯誤通知郵件。

Google Drive (DriveApp): 在指定的資料夾中建立和儲存 HTML 檔案。

請確保您的 Google Workspace 帳戶允許 Apps Script 使用這些服務。