/**
 * 一鍵產生 PDF 工具 － Google 試算表自動寫入
 *
 * 使用方式：
 * 1. 新增/開啟一個 Google 試算表
 * 2. 擴充功能 → Apps Script
 * 3. 將此檔案內容貼上，儲存後「部署」→「新增部署」→ 類型選「網頁應用程式」
 * 4. 執行身分：我；存取權：任何人 → 部署後複製「網頁應用程式 URL」
 * 5. 在工具頁「給客人明細」區的「Google 試算表寫入」欄位貼上該 URL
 * 6. 之後每次一鍵產生 PDF，資料會自動寫入此試算表
 */

function doPost(e) {
  try {
    var json = e.postData ? e.postData.contents : null;
    if (!json) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No body' })).setMimeType(ContentService.MimeType.JSON);
    }
    var data = JSON.parse(json);
    var header = data.header || [];
    var rows = data.rows || [];
    if (rows.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ ok: true, written: 0 })).setMimeType(ContentService.MimeType.JSON);
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow === 0 && header.length > 0) {
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
      sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
      lastRow = 1;
    }
    if (rows.length > 0) {
      sheet.getRange(lastRow + 1, 1, lastRow + rows.length, rows[0].length).setValues(rows);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, written: rows.length })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}
