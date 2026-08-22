/**
 * BNI 引薦案例素材庫 — Google 試算表後端橋接腳本
 *
 * 使用方式：
 * 1. 打開你的 Google 試算表 → 擴充功能 → Apps Script
 * 2. 把這個檔案的內容整個貼進去（取代預設的 myFunction 內容）
 * 3.（建議）左側齒輪圖示「專案設定」→ 最下方「指令碼屬性」→新增屬性
 *    鍵：WRITE_SECRET   值：自己取一個不容易猜到的字串（例如一串英數字）
 *    這組值等一下也要設進 Render 的環境變數 SHEETS_WEBAPP_SECRET，兩邊要一致。
 *    （不設也可以運作，只是任何人拿到你的網頁應用程式網址就能寫入資料。）
 * 4. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    - 執行身分：我
 *    - 誰可以存取：任何人
 *    部署後會拿到一個網址（結尾是 /exec），複製起來設成 Render 的環境變數 SHEETS_WEBAPP_URL
 * 5. 之後如果修改這份程式碼，要用「部署」→「管理部署作業」→ 編輯（鉛筆圖示）→
 *    版本選「新版本」→ 部署，網址才會套用最新程式碼（單純存檔不會生效）。
 *
 * 試算表第一列（標題列）請填入以下欄位名稱，順序不拘，但文字要完全一致：
 * 誰給的引薦 | 引薦給誰 | 金額 | 引薦類型 | 內部／外部 | 案例簡述 | 熱心填寫者 | 填寫時間
 *
 * 「引薦類型」欄位的值請填「週期性固定引薦」或「單次引薦」
 * 「內部／外部」欄位的值請填「內部引薦」或「外部引薦」
 * 建議幫這兩欄設「資料驗證」下拉選單（選取欄位 → 資料 → 資料驗證），
 * 這樣手動輸入時就不會打錯字，網頁那邊才讀得到正確分類。
 * 「引薦給誰」這欄網頁表單只能選會員名單裡的名字，
 * 你手動在試算表輸入時也建議設成同一份會員名單的下拉選單，維持資料一致。
 * 「誰給的引薦」「熱心填寫者」這兩欄網頁表單可以選會員名單裡的名字，也可以選「其他」
 * 自行輸入名單外的名字（例如來賓、非會員協助填寫的情況），所以允許填名單外的名字。
 * 「填寫時間」手動新增案例時可以留空，網頁會顯示「時間未填」，不影響其他功能。
 */

const SHEET_NAME = ""; // 留空 = 使用試算表的第一個工作表分頁；如果你的分頁不是第一個，改成分頁名稱，例如 "案例"

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const headers = values[0];
  const records = values
    .slice(1)
    .filter((row) => row.some((cell) => cell !== "" && cell !== null))
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        const val = row[i];
        // Date 物件轉成 ISO 字串，方便 Node 那邊直接 new Date() 解析
        obj[h] = val instanceof Date ? val.toISOString() : val;
      });
      return obj;
    });

  return ContentService.createTextOutput(JSON.stringify(records))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  const requiredSecret = props.getProperty("WRITE_SECRET");

  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonError_("請求格式不正確");
  }

  if (requiredSecret && body.secret !== requiredSecret) {
    return jsonError_("unauthorized");
  }

  const sheet = getSheet_();
  const headers = sheet.getDataRange().getValues()[0];

  const valueByHeader = {
    "誰給的引薦": body.referrer || "",
    "引薦給誰": body.recipient || "",
    "金額": body.amount || 0,
    "引薦類型": body.referralTypeLabel || "",
    "內部／外部": body.sourceLabel || "",
    "案例簡述": body.note || "",
    "熱心填寫者": body.submitter || "",
    "填寫時間": new Date(),
  };

  // 依照試算表實際的欄位順序組一列，避免欄位順序跟這裡假設的不一樣時寫錯欄位
  const row = headers.map((h) => (h in valueByHeader ? valueByHeader[h] : ""));
  sheet.appendRow(row);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError_(message) {
  return ContentService.createTextOutput(JSON.stringify({ error: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
