/**
 * 潛意識自我破壞程序測驗 — 後端接收端點
 *
 * 部署方式：
 * 1. 在 https://script.google.com 建立新專案，貼上本檔案內容。
 * 2. 上方工具列「部署」→「新增部署作業」。
 * 3. 類型選「網頁應用程式」；「執行身分」選「我」；「誰可以存取」選「任何人」。
 * 4. 部署後複製 .../exec 網址，貼到前端 index.html 的 GAS_ENDPOINT_URL 常數。
 * 5. 若之後修改本檔案，需要「管理部署作業」→ 編輯 → 新版本，才會生效。
 *
 * 前端資料傳輸方式：改用「隱藏表單 + 隱藏 iframe」提交（非 fetch），
 * 原因是像 Google Sites 這類會把嵌入內容放進沙盒 iframe 的平台，
 * 常會限制 fetch/XHR 對外部網域的請求（CSP connect-src），但一般的
 * <form> 提交屬於瀏覽器原生的導覽行為，不受此限制影響，相容性更好。
 *
 * 注意：Apps Script 的 e.parameter 在 doPost 中「不會」自動解析
 * application/x-www-form-urlencoded 的表單本文（這點與大多數人預期不同），
 * 實際收到的內容一律要從 e.postData.contents 讀取原始字串
 * （表單提交時內容會是 "payload=<url-encoded JSON>"，需自行解碼；
 * 若直接以 JSON 字串送出 body，例如用 fetch，內容本身就是完整 JSON）。
 */

var SHEET_ID = "18UTL2sB0dkR_CpvQUhP8WYlBSfsDVomRYeO4gEQD7L8";
var FOLDER_ID = "1yOjgnWBiZLlCd6f7_wO6q7mJihW_dY5a";
var DOMAIN_TITLES = ["領域一：金錢與事業", "領域二：關係與合作", "領域三：自我價值與內在穩定", "領域四：能量與執行力"];
var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
var MAX_IMAGES = 4;
var MAX_IMAGE_BYTES = 5 * 1024 * 1024; // 每張圖片解碼後上限 5MB，避免濫用塞爆雲端硬碟
var MAX_TEXT_LENGTH = 2000; // 每個領域文字欄位上限，避免異常巨量文字寫入試算表

/**
 * 這個網址沒有帳號登入機制（前端是純靜態頁面），任何知道網址的人都能直接呼叫。
 * 以下驗證只能擋掉明顯不完整／異常的請求，降低垃圾資料誤寫入的機率，
 * 並不是真正的存取控制，無法防止有心人蓄意濫用。
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("error: no data received");
    }

    var raw = extractPayloadString(e.postData.contents);
    if (!raw) {
      return ContentService.createTextOutput("error: no data received");
    }

    var data;
    try {
      data = JSON.parse(raw);
    } catch (parseErr) {
      return ContentService.createTextOutput("error: invalid JSON - " + parseErr.message);
    }

    var name = String(data.name || "").trim();
    var email = String(data.email || "").trim();
    var domainA = String(data.domainA || "");
    var domainB = String(data.domainB || "");
    var domainC = String(data.domainC || "");
    var domainD = String(data.domainD || "");
    var images = Array.isArray(data.imageB64) ? data.imageB64 : []; // 依領域一～四順序的 4 筆 base64 圖片（data URL 或純 base64 皆可）

    if (!name) {
      return ContentService.createTextOutput("error: missing name");
    }
    if (!EMAIL_RE.test(email)) {
      return ContentService.createTextOutput("error: invalid email");
    }
    if (images.length > MAX_IMAGES) {
      return ContentService.createTextOutput("error: too many images");
    }
    var textFields = [domainA, domainB, domainC, domainD];
    for (var t = 0; t < textFields.length; t++) {
      if (textFields[t].length > MAX_TEXT_LENGTH) {
        return ContentService.createTextOutput("error: domain text too long");
      }
    }

    var uploaded = saveImagesToDrive(images, name);

    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var targetRow = sheet.getLastRow() + 1;

    sheet.getRange(targetRow, 1, 1, 6).setValues([[name, email, domainA, domainB, domainC, domainD]]);
    sheet.getRange(targetRow, 7).setRichTextValue(buildFinalReportRichText(uploaded));
    sheet.getRange(targetRow, 3, 1, 5).setWrap(true); // 領域一～四＋最終報告：多行內容自動換行顯示

    // 不自訂任何 Header，純文字回應；表單提交本來就不受 CORS 限制，這裡維持單純文字方便日後除錯
    return ContentService.createTextOutput("success");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err.message);
  }
}

/**
 * 解析請求本文：
 * - 表單提交（application/x-www-form-urlencoded）時，內容會是 "payload=<url-encoded JSON>"，需先解碼。
 * - 若內容本身就是 JSON（例如直接以 fetch 送出 JSON 字串 body），則原樣使用。
 */
function extractPayloadString(contents) {
  var match = /^payload=([\s\S]*)$/.exec(contents);
  if (!match) return contents;
  return decodeURIComponent(match[1].replace(/\+/g, " "));
}

/**
 * 將 Base64 圖片逐一解碼、存入指定資料夾，並嘗試設定「知道連結的人皆可檢視」。
 * 權限設定失敗（常見於企業／學校網域政策限制）時只記錄警告，不中斷流程。
 * 檔名格式：「<姓名> - <領域X名稱>.jpg」。
 * 回傳格式：[{ label: "領域一：金錢與事業", url: "https://drive..." }, ...]（只含成功上傳的項目）。
 */
function saveImagesToDrive(images, name) {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var results = [];

  for (var i = 0; i < images.length; i++) {
    var raw = images[i];
    if (!raw) continue;

    var label = DOMAIN_TITLES[i] || ("領域" + (i + 1));

    try {
      var base64Data = raw.indexOf(",") !== -1 ? raw.split(",")[1] : raw;

      // 粗略估算解碼後大小（base64 每 4 字元對應約 3 bytes），超過上限就跳過這張圖
      var approxBytes = Math.floor(base64Data.length * 3 / 4);
      if (approxBytes > MAX_IMAGE_BYTES) {
        console.warn("「" + label + "」圖片超過大小上限（約 " + approxBytes + " bytes），已略過");
        continue;
      }

      var decoded = Utilities.base64Decode(base64Data);
      var fileName = name + " - " + label + ".jpg";
      var blob = Utilities.newBlob(decoded, "image/jpeg", fileName);
      var file = folder.createFile(blob);

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {
        console.warn("設定檔案共用權限失敗（可能受組織網域政策限制），僅記錄警告，流程繼續：" + shareErr);
      }

      results.push({ label: label, url: file.getUrl() });
    } catch (fileErr) {
      console.warn("處理「" + label + "」圖片時發生錯誤：" + fileErr);
    }
  }

  return results;
}

/**
 * 把 [{label, url}, ...] 組成單一儲存格的富文字內容：
 * 每行只顯示「領域X名稱」這段文字本身，並把該整行文字直接設為可點擊的內嵌超連結
 * （連到該領域的圖片網址），不額外顯示原始網址，畫面更簡潔。
 */
function buildFinalReportRichText(items) {
  if (items.length === 0) {
    return SpreadsheetApp.newRichTextValue().setText("（無圖片）").build();
  }

  var lines = items.map(function (item) { return item.label; });
  var text = lines.join("\n");

  var builder = SpreadsheetApp.newRichTextValue().setText(text);
  var cursor = 0;
  items.forEach(function (item) {
    var start = cursor;
    var end = start + item.label.length;
    builder.setLinkUrl(start, end, item.url);
    cursor = end + 1; // +1 對應之後 join("\n") 補上的換行字元
  });

  return builder.build();
}
