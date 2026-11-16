/**
 * 鍋淬訂單接收器（Google Apps Script 網頁應用程式）
 * 功能：接收網站結帳表單送出的訂單 JSON，重新計算價格（不信任前端金額），
 *       並寫入試算表的「訂單」工作表。
 *
 * 部署步驟請見 gas/README.md
 */

// ===== 設定 =====
const SHEET_NAME = '訂單';
const TOKEN = '';            // 可選：設定一組共用密碼，前端 index.html 的 ORDER_TOKEN 要填相同值；留空則不驗證
const TIME_ZONE = 'Asia/Taipei';
const MAX_ITEMS = 30;        // 單筆訂單最多品項數
const MAX_QTY = 50;          // 單一品項最大數量

// 價格表：必須與 index.html 的 products 同步。價格以這裡為準。
const PRICES = {
  p1: { name: '經典鮮肉鍋貼', price: 95 },
  p2: { name: '香辣椒麻鍋貼', price: 110 },
  p3: { name: '鮮蝦韭黃鍋貼', price: 135 },
  p4: { name: '手工高麗菜水餃', price: 90 },
  p5: { name: '韭菜鮮肉水餃', price: 95 },
  p6: { name: '暖心酸辣湯', price: 55 },
  p7: { name: '清香冷泡茶', price: 45 },
  p8: { name: '黃金泡菜', price: 60 },
  p9: { name: '芝麻涼拌時蔬', price: 55 },
  gift: { name: '招牌招財鍋貼禮盒', price: 680 }
};

const HEADERS = ['建立時間', '訂單編號', '取餐方式', '姓名', '電話', 'Email', '地址', '備註', '品項', '小計(NT$)', '狀態'];

function doGet() {
  return json_({ ok: true, service: 'guocui-orders' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (TOKEN && data.token !== TOKEN) return json_({ ok: false, error: 'unauthorized' });

    const order = validateAndPrice_(data);
    const sheet = getSheet_();
    const orderId = nextOrderId_(sheet);
    const now = new Date();

    sheet.appendRow([
      Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'),
      orderId,
      order.method,
      order.name,
      order.phone,
      order.email,
      order.address,
      order.notes,
      order.itemsText,
      order.subtotal,
      '新訂單'
    ]);

    return json_({ ok: true, orderId: orderId, subtotal: order.subtotal });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function validateAndPrice_(d) {
  const methods = ['內用', '外帶', '宅配'];
  if (methods.indexOf(d.method) === -1) throw new Error('取餐方式不正確');

  const name = clean_(d.name, 50);
  const phone = clean_(d.phone, 20);
  if (!name) throw new Error('請填寫姓名');
  if (!/^[0-9+()\-\s]{8,20}$/.test(phone)) throw new Error('電話格式不正確');

  const email = clean_(d.email, 100);
  const address = clean_(d.address, 200);
  if (d.method === '宅配' && !address) throw new Error('宅配需填寫地址');
  const notes = clean_(d.notes, 300);

  if (!Array.isArray(d.items) || d.items.length === 0) throw new Error('購物車是空的');
  if (d.items.length > MAX_ITEMS) throw new Error('品項過多');

  let subtotal = 0;
  const parts = [];
  d.items.forEach(function (it) {
    const p = PRICES[it.id];
    const qty = Math.floor(Number(it.qty));
    if (!p) throw new Error('未知商品：' + it.id);
    if (!(qty >= 1 && qty <= MAX_QTY)) throw new Error('數量不正確');
    subtotal += p.price * qty;
    parts.push(p.name + ' × ' + qty + '（NT$' + p.price * qty + '）');
  });

  return {
    method: d.method, name: name, phone: phone, email: email,
    address: address, notes: notes, subtotal: subtotal,
    itemsText: parts.join('\n')
  };
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 訂單編號：GC + 日期 + 流水號（在 ScriptLock 內執行，不會重複）
function nextOrderId_(sheet) {
  const today = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyyMMdd');
  const seq = sheet.getLastRow(); // 含標題列，恰好等於下一個序號的起點
  return 'GC' + today + '-' + ('000' + seq).slice(-4);
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, max);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
