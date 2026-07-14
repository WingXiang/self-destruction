/**
 * 潛意識自我破壞程序檢測 — 後端接收端點
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

// 結果信件用：寄件顯示名稱、寄件地址與內文中的三個內嵌連結（與前端 index.html 頁面上的文案、連結一致）
// 注意：SENDER_EMAIL 必須先在「執行這支程式的 Google 帳號」的 Gmail 設定裡，
// 於「帳戶和匯入 → 寄件地址」新增並完成驗證，否則 GmailApp.sendEmail 的 from 會失敗。
var SENDER_NAME = "Linda 洋溢人生潛意識信念專家";
var SENDER_EMAIL = "linda.hsyh@gmail.com";
var CTA_FORM_URL = "https://forms.gle/CtouUvrLKeAV8z258";
var IG_URL = "https://www.instagram.com/linda_hsyh";
var FB_URL = "https://www.facebook.com/Linda.hsyh";

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

    var blobs = buildImageBlobs(images, name); // 只解碼一次，Drive 上傳與寄信附件共用同一批 blob
    var uploaded = saveImagesToDrive(blobs);

    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var targetRow = sheet.getLastRow() + 1;

    sheet.getRange(targetRow, 1, 1, 6).setValues([[name, email, domainA, domainB, domainC, domainD]]);
    sheet.getRange(targetRow, 7).setRichTextValue(buildFinalReportRichText(uploaded));
    sheet.getRange(targetRow, 3, 1, 5).setWrap(true); // 領域一～四＋最終報告：多行內容自動換行顯示

    sendResultEmail(email, name, blobs);

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
 * 將 Base64 圖片逐一解碼成 Blob，檔名格式：「<姓名> - <領域X名稱>.jpg」。
 * 只解碼一次，讓 Drive 上傳（saveImagesToDrive）與寄信附件（sendResultEmail）共用同一批 blob，
 * 避免同一組圖片被重複解碼兩次。
 * 回傳格式：[{ label: "領域一：金錢與事業", blob: Blob }, ...]（只含成功解碼、且未超過大小上限的項目）。
 */
function buildImageBlobs(images, name) {
  var blobs = [];

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
      blobs.push({ label: label, blob: blob });
    } catch (decodeErr) {
      console.warn("解碼「" + label + "」圖片時發生錯誤：" + decodeErr);
    }
  }

  return blobs;
}

/**
 * 把已解碼的圖片 Blob 存入指定資料夾，並嘗試設定「知道連結的人皆可檢視」。
 * 權限設定失敗（常見於企業／學校網域政策限制）時只記錄警告，不中斷流程。
 * 回傳格式：[{ label: "領域一：金錢與事業", url: "https://drive..." }, ...]（只含成功上傳的項目）。
 */
function saveImagesToDrive(blobsWithLabel) {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var results = [];

  blobsWithLabel.forEach(function (item) {
    try {
      var file = folder.createFile(item.blob);

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {
        console.warn("設定檔案共用權限失敗（可能受組織網域政策限制），僅記錄警告，流程繼續：" + shareErr);
      }

      results.push({ label: item.label, url: file.getUrl() });
    } catch (fileErr) {
      console.warn("上傳「" + item.label + "」圖片時發生錯誤：" + fileErr);
    }
  });

  return results;
}

/**
 * 把四張結果圖片以附件形式寄給填答者，內文文案與前端頁面上的「最後提醒」「預約諮詢」段落一致。
 * 寄件帳號＝目前部署這支 Apps Script 的 Google 帳號（「執行身分：我」），寄件顯示名稱另外指定。
 * 寄信失敗只記錄警告、不中斷流程——試算表與雲端硬碟的寫入已經在這之前完成，不應該因為寄信失敗而報錯。
 */
function sendResultEmail(email, name, blobsWithLabel) {
  if (blobsWithLabel.length === 0) return; // 沒有任何圖片可寄送時略過

  try {
    GmailApp.sendEmail(email, buildResultEmailSubject(name), buildResultEmailPlainText(name), {
      from: SENDER_EMAIL,
      name: SENDER_NAME,
      htmlBody: buildResultEmailHtml(name),
      attachments: blobsWithLabel.map(function (item) { return item.blob; })
    });
  } catch (mailErr) {
    console.warn("寄送結果信件失敗（不影響試算表與雲端硬碟寫入）：" + mailErr);
  }
}

function buildResultEmailSubject(name) {
  return "[潛意識自我破壞程序檢測] " + name + "，你的專屬結果出爐了";
}

// 純文字版信件內容：極少數不支援 HTML 信件的信箱會退回顯示這一版，連結直接寫出網址
function buildResultEmailPlainText(name) {
  return (
    name + " 你好：\n\n" +
    "謝謝你剛剛完成了「潛意識自我破壞程序檢測」。\n\n" +
    "附件是你四大領域的完整結果圖片，記得收藏起來，有空的時候可以重新回顧一次。\n\n" +
    "很多卡住，不是因為你不夠努力，而是你一直用舊的底層設定，在面對新的人生階段。\n" +
    "當你願意看見它，你就已經開始鬆動它了。\n\n" +
    "如果你想更深入了解自己這次測出來的自我破壞模式，歡迎預約每月限量 3 場的\n" +
    "潛意識信念校準諮詢，我會陪你一起把它重新校準成支持你前進的力量。\n\n" +
    "立即填寫預約表單：" + CTA_FORM_URL + "\n\n" +
    "祝福你\n" +
    "Linda 洋溢人生潛意識信念專家\n" +
    "關注 Linda 的 Instagram：" + IG_URL + "\n" +
    "關注 Linda 的 Facebook：" + FB_URL
  );
}

function buildResultEmailHtml(name) {
  return (
    name + " 你好：<br><br>" +
    "謝謝你剛剛完成了「潛意識自我破壞程序檢測」。<br><br>" +
    "附件是你四大領域的完整結果圖片，記得收藏起來，有空的時候可以重新回顧一次。<br><br>" +
    "很多卡住，不是因為你不夠努力，而是你一直用舊的底層設定，在面對新的人生階段。<br>" +
    "當你願意看見它，你就已經開始鬆動它了。<br><br>" +
    "如果你想更深入了解自己這次測出來的自我破壞模式，歡迎預約每月限量 3 場的<br>" +
    "潛意識信念校準諮詢，我會陪你一起把它重新校準成支持你前進的力量。<br><br>" +
    "<a href=\"" + CTA_FORM_URL + "\" style=\"font-size:16px;\">立即填寫預約表單</a><br><br>" +
    "祝福你<br>" +
    "Linda 洋溢人生潛意識信念專家<br>" +
    "<a href=\"" + IG_URL + "\">關注 Linda 的 Instagram</a><br>" +
    "<a href=\"" + FB_URL + "\">關注 Linda 的 Facebook</a>"
  );
}

/**
 * 手動測試信件功能用：不需要跑過整個檢測流程，直接在 Apps Script 編輯器裡選這個函式、
 * 按「執行」即可單獨測試寄信是否正常。
 *
 * 第一次執行時，Google 會跳出授權視窗，請務必允許「以你的名義傳送電子郵件」的權限——
 * 這個授權只有在「手動執行」時才會出現，光是部署新版本並不會觸發。
 *
 * 使用方式：把下面的 email 換成你要收測試信的信箱，選這個函式，按「執行」，
 * 完成授權後檢查該信箱是否收到測試信、寄件人是否顯示 linda.hsyh@gmail.com。
 */
function testSendResultEmail() {
  var testEmail = "換成你要收測試信的信箱@gmail.com"; // ← 執行前記得先改這裡
  GmailApp.sendEmail(testEmail, "【測試】" + buildResultEmailSubject("測試"), buildResultEmailPlainText("測試"), {
    from: SENDER_EMAIL,
    name: SENDER_NAME,
    htmlBody: buildResultEmailHtml("測試")
  });
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
