/**
 * 晨光咖啡訂單 → Google Sheets
 * 貼到試算表的「擴充功能 → Apps Script」，然後部署成「網頁應用程式」。
 */
const SPREADSHEET_ID = '14SvXC2K7rUXiYRVfYtzS1F_P93k2X-_EOtx3aiHB5XM';

const HEADERS = ['訂單編號', '時間', '姓名', '電話', 'Email', '品項', '數量', '合計', '取貨方式', '備註'];
const PRICES = { '焦糖肉桂拿鐵': 99, '經典美式': 70, '燕麥奶拿鐵': 110 };

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const p = (e && e.parameter) || {};

    const name = clean(p.name, 40);
    const phone = clean(p.phone, 15);
    const item = clean(p.item, 20);
    const qty = Math.floor(Number(p.qty));
    if (!name || !phone) return json({ ok: false, error: '缺少姓名或電話' });
    if (!(item in PRICES)) return json({ ok: false, error: '品項錯誤' });
    if (!(qty >= 1 && qty <= 20)) return json({ ok: false, error: '數量錯誤' });

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }
    // 電話欄設為純文字，避免開頭的 0 被當成數字吃掉
    sheet.getRange('D:D').setNumberFormat('@');

    const now = new Date();
    const orderId = 'ML' + Utilities.formatDate(now, 'Asia/Taipei', 'yyMMddHHmmss') +
                    Math.floor(Math.random() * 90 + 10);
    sheet.appendRow([
      orderId,
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      name,
      phone,
      clean(p.email, 80),
      item,
      qty,
      PRICES[item] * qty,          // 由伺服器計算金額，不信任前端
      clean(p.pickup, 10),
      clean(p.note, 200),
    ]);
    return json({ ok: true, orderId: orderId });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, message: '晨光咖啡訂單 API 運作中' });
}

/** 去除空白、限制長度，並避免試算表公式注入 */
function clean(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
