/**
 * 這是 Google Apps Script 程式碼。
 * 請將此程式碼貼到您的 Apps Script 編輯器中（在 Google Sheet 或其他 Google 服務中選擇「擴充功能」->「Apps Script」）。
 *
 * 🚨 【重要】請務必在貼上後執行一次 setupTrigger 函式來設定每月自動觸發。
 */

// ----------------------------------------------------
// 【設定區域】請修改以下變數
// ----------------------------------------------------

/**
 * @type {string} baseDomain - 電子報網址的固定前綴部分 (不含結尾斜線)。
 */
const baseDomain = "https://epaper.twcert.org.tw"; 

/**
 * @type {string} recipientEmail - 指定轉發電子報的收件人信箱。可使用逗號分隔多個信箱。
 */
const recipientEmail = "mail@yourdomain"; // <-- 請修改為您的收件信箱，可使用逗號分隔

/**
 * @type {string} driveFolderId - 指定要儲存電子報的 Google Drive 資料夾 ID。
 * 💡 如何取得 ID: 打開您的 Google Drive 資料夾，網址列中 /folders/ 後面的那一串亂碼就是 ID。
 */
const driveFolderId = "{your folder id}}"; // <-- 【重要】請修改為您的資料夾 ID

// ----------------------------------------------------
// 主要函式：每月自動執行
// ----------------------------------------------------

/**
 * 主函式：自動抓取上個月電子報並發送到指定信箱。
 * 【注意】觸發器執行時不會傳入參數，會使用真實日期。
 * @param {Date | undefined} testDate - 測試時可傳入模擬日期，否則使用真實日期。
 */
function main(testDate) {
  // 1. 取得上個月的年/月路徑
  const previousMonthPath = getPreviousMonthPath(testDate);
  const fullUrl = baseDomain + "/" + previousMonthPath;
  const monthName = previousMonthPath.replace("/", "").replace("_", "/"); // 例如 2025/10
  const subject = `【情資蒐集】[TWCA] ${monthName} 電子報`;

  Logger.log(`準備收集網址: ${fullUrl}`);

  // 2. 抓取電子報內容
  const htmlContent = fetchAndExtractContent(fullUrl);

  // 3. 檢查內容並發送郵件
  if (htmlContent.includes("存取網址時發生錯誤") || htmlContent.includes("無法從網址取得電子報內容")) {
    // 如果發生錯誤，只發送錯誤通知
    MailApp.sendEmail({
      to: recipientEmail,
      subject: `【錯誤通知】TWCA 電子報自動收集失敗 - ${monthName}`,
      body: `嘗試收集網址 ${fullUrl} 時發生錯誤。\n錯誤訊息:\n${htmlContent}`,
    });
    Logger.log("執行失敗，已發送錯誤通知郵件。");
  } else {
    // 💡 新增：儲存至 Google Drive
    const filename = `${monthName.replace("/", "_")}_TWCA電子報.html`;
    saveToDrive(filename, htmlContent);

    // 成功抓取，發送電子報
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlContent, // 以 HTML 格式發送內容
    });
    Logger.log(`成功收集並發送電子報給 ${recipientEmail}。`);
  }
}

// ----------------------------------------------------
// 輔助函式
// ----------------------------------------------------

/**
 * 計算上一個月的年份和月份，並格式化為 'YYYY_MM/' 格式。
 * @param {Date | undefined} dateToUse - 要用來計算「上一個月」的基準日期。
 * 範例: 如果傳入 2025/09/02，則回傳 '2025_08/'。
 * @returns {string} 上個月的 URL 路徑部分。
 */
function getPreviousMonthPath(dateToUse) {
  // 複製日期物件，確保不會修改傳入的原始物件
  const date = new Date(dateToUse || new Date()); 
  
  // 將日期設定為上個月
  date.setMonth(date.getMonth() - 1);

  const year = date.getFullYear();
  // 月份加 1 (因為 getMonth() 回傳 0-11)，然後用 padStart 補零
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  
  // 格式化為 'YYYY_MM/'
  const path = `${year}_${month}/`;
  Logger.log(`計算路徑結果 (基準日: ${(dateToUse || new Date()).toLocaleDateString()}): ${path}`);
  return path;
}

/**
 * 從指定網址抓取網頁內容，並嘗試提取主要內容。
 * @param {string} url - 要抓取的網址。
 * @returns {string} 提取到的 HTML 內容或錯誤訊息。
 */
function fetchAndExtractContent(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true // 避免 404/500 錯誤直接中斷腳本
    });
    
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      throw new Error(`HTTP 錯誤碼: ${responseCode}`);
    }

    const html = response.getContentText("UTF-8");
    
    // 擷取 <body> 標籤內的所有內容
    const bodyMatch = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);

    if (bodyMatch && bodyMatch[1]) {
      let bodyContent = bodyMatch[1];
      
      const baseUrl = baseDomain; 
      
      // 替換以 '/xxx' 開頭的相對路徑為絕對路徑
      bodyContent = bodyContent.replace(/(src|href)="(\/.*?)"/gi, `$1="${baseUrl}$2"`);
      
      return bodyContent;
    } else {
      return '無法從網址取得電子報內容。請檢查網頁結構是否改變。';
    }
  } catch (e) {
    return '存取網址時發生錯誤：' + e.message;
  }
}

/**
 * 【新增函式】將內容儲存為 HTML 檔案到指定的 Google Drive 資料夾。
 * @param {string} filename - 要儲存的檔案名稱 (例如: '2025_10_TWCA電子報.html')。
 * @param {string} content - 檔案內容 (HTML 格式)。
 */
function saveToDrive(filename, content) {
  try {
    if (!driveFolderId || driveFolderId === "YOUR_DRIVE_FOLDER_ID") {
      Logger.log("Drive 資料夾 ID 未設定，跳過儲存到 Drive 的步驟。");
      return;
    }
    
    const folder = DriveApp.getFolderById(driveFolderId);
    
    // 檢查檔案是否已存在，避免重複儲存 (這裡使用 getFilesByName，如果檔案名相同，則跳過)
    const files = folder.getFilesByName(filename);
    if (files.hasNext()) {
      Logger.log(`檔案 ${filename} 已存在於 Drive，跳過重複儲存。`);
      return;
    }

    // 儲存檔案，設定 MimeType 為 HTML
    folder.createFile(filename, content, MimeType.HTML);
    Logger.log(`✅ 成功將電子報儲存為檔案: ${filename}`);

  } catch (e) {
    Logger.log(`❌ 儲存檔案到 Google Drive 失敗：請確認資料夾 ID 是否正確，以及 Apps Script 是否有 Drive 權限。錯誤: ${e.message}`);
    
    // 儲存 Drive 失敗時，發送錯誤通知，但不中斷主流程
    MailApp.sendEmail({
      to: recipientEmail,
      subject: `【部分失敗通知】TWCA 電子報 Drive 儲存失敗`,
      body: `嘗試將檔案 ${filename} 儲存到 Drive 時發生錯誤。\n請檢查 driveFolderId 設定和 Apps Script 的 Drive 權限。\n錯誤訊息:\n${e.message}`,
    });
  }
}


// ----------------------------------------------------
// 測試函式
// ----------------------------------------------------

/**
 * 測試函式：模擬在特定日期執行 main()，以確保邏輯正確。
 * 您可以修改下方的日期來測試不同月份的邊界情況。
 */
function testMain() {
  Logger.log("--- 開始執行測試 ---");
  
  // 1. 模擬在 2025年1月2日 執行 (應取得 2024_12/) - 測試跨年情況
  // 注意：月份在 JavaScript 中是 0-indexed (0=一月, 11=十二月)
  const simulatedNewYearDate = new Date(2025, 0, 2); 
  Logger.log("--- 測試 1: 跨年邊界 (模擬 2025/01/02 執行) ---");
  main(simulatedNewYearDate);
  
  // 2. 模擬在 2025年09月2日 執行 (應取得 2025_08/) - 測試一般情況
  const simulatedNormalDate = new Date(2025, 8, 2); 
  Logger.log("--- 測試 2: 一般月份 (模擬 2025/09/02 執行) ---");
  main(simulatedNormalDate);
  
  Logger.log("--- 測試結束。請檢查 Log 紀錄中的網址是否正確，以及信箱是否收到測試郵件 ---");
}

// ----------------------------------------------------
// 觸發器設定函式 (只需執行一次)
// ----------------------------------------------------

/**
 * 設置時間驅動觸發器，使 main() 函式在每個月的 1 號自動執行。
 * 【重要】請手動執行此函式一次來啟用自動排程。
 */
function setupTrigger() {
  // 刪除所有現有的觸發器，以防重複設定
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "main") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 創建新的時間驅動觸發器：每月 1 號執行 main() 函式
  ScriptApp.newTrigger("main")
      .timeBased()
      .onDayOfMonth(1)
      .atHour(12) // 設定在中午 12 點左右執行 (Apps Script 會有幾分鐘的浮動)
      .create();

  Logger.log("自動觸發器已設定完成：將在每月的 1 號中午 12 點執行 main() 函式。");
}
