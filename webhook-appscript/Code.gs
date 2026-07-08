/**
 * WEBHOOK NHẬN ĐƠN HÀNG — LP HỒNG SÂM DONGWHA
 * Form landing page → Google Sheet + Email + Zalo Bot.
 * Cách cài: xem HUONG-DAN-WEBHOOK.md (phần Zalo Bot ở cuối file đó)
 */

var SHEET_NAME = 'Đơn hàng';
var NOTIFY_EMAIL = 'lctbinh0006@gmail.com';   // Để '' nếu không muốn nhận email

// ==== ZALO BOT — điền 2 dòng dưới theo hướng dẫn ====
var ZALO_BOT_TOKEN = '3954983369384332927:APeVViiDQFabibvfpjxkxGcBRWlYGztiDgAwErSWyLQITsbZseAzKbxVLTpThWQv';
var ZALO_CHAT_IDS  = ['3f4e22e5d1ab38f561ba'];   // Bình Lê. Thêm người nhận: ['id1','id2']

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    var headers = ['Thời gian', 'Mã đơn', 'Họ tên', 'SĐT', 'Địa chỉ', 'Sản phẩm', 'Tổng SL',
      'Bump hộp 2', 'Thanh toán', 'Tổng tiền', 'Ghi chú', 'Trạng thái gọi',
      'UTM Source', 'UTM Medium', 'UTM Campaign', 'UTM Content', 'UTM Term', 'Ref', 'fbclid', 'Event ID', 'URL nguồn'];
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d1fae5');
      sh.setFrozenRows(1);
    }

    var payLabel = (data.payment_method === 'cod') ? 'COD' : 'Chuyển khoản';
    var products = formatProducts(data.products || '') + (data.order_bump ? ' + 1 hộp Hồng Sâm (bump)' : '');
    var total = Number(data.total_amount || 0).toLocaleString('vi-VN') + 'đ';

    sh.appendRow([
      new Date(), data.code || '', data.fullname || '', "'" + (data.phone || ''), data.address || '',
      products, data.quantity || '',
      data.order_bump ? 'Có' : 'Không', payLabel, data.total_amount || '',
      data.note || '', 'Chưa gọi',
      data.utm_source || '', data.utm_medium || '', data.utm_campaign || '',
      data.utm_content || '', data.utm_term || '', data.ref || '', data.fbclid || '',
      data.event_id || '', data.event_source_url || ''
    ]);

    var msg = '🛒 ĐƠN MỚI [' + payLabel + ']' +
      '\n👤 ' + (data.fullname || '') +
      '\n📞 ' + (data.phone || '') +
      '\n📍 ' + (data.address || '') +
      '\n📦 ' + products +
      '\n💰 ' + total +
      (data.note ? '\n📝 ' + data.note : '') +
      '\n#' + (data.code || '') +
      '\n\n→ GỌI XÁC NHẬN TRONG 5 PHÚT!';

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, '🛒 ĐƠN MỚI [' + payLabel + '] ' + (data.fullname || '') + ' – ' + total, msg);
    }
    sendZalo(msg);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Gửi tin nhắn đến Zalo Bot cho tất cả chat_id đã khai báo */
function sendZalo(text) {
  if (!ZALO_BOT_TOKEN || !ZALO_CHAT_IDS.length) return;
  ZALO_CHAT_IDS.forEach(function (chatId) {
    try {
      UrlFetchApp.fetch('https://bot-api.zaloplatforms.com/bot' + ZALO_BOT_TOKEN + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chat_id: chatId, text: text }),
        muteHttpExceptions: true
      });
    } catch (e) { /* không để lỗi Zalo làm mất đơn */ }
  });
}

/**
 * CHẠY HÀM NÀY 1 LẦN để lấy chat_id:
 * 1. Điền ZALO_BOT_TOKEN ở trên → Lưu
 * 2. Mở Zalo, nhắn 1 tin bất kỳ cho bot vừa tạo (VD: "hi")
 * 3. Chọn hàm getZaloChatId → Run → xem Log (Ctrl+Enter) → copy chat_id dán vào ZALO_CHAT_IDS
 */
function getZaloChatId() {
  if (!ZALO_BOT_TOKEN) { Logger.log('⚠️ Chưa điền ZALO_BOT_TOKEN'); return; }
  Logger.log('⏳ Đang chờ tin nhắn... BÂY GIỜ hãy mở Zalo nhắn 1 tin cho bot (có 90 giây)');
  var found = [];
  var start = Date.now();
  while (Date.now() - start < 90000 && !found.length) {
    var res = UrlFetchApp.fetch('https://bot-api.zaloplatforms.com/bot' + ZALO_BOT_TOKEN + '/getUpdates', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ timeout: 25 }),
      muteHttpExceptions: true
    });
    Logger.log('📩 Zalo trả về [' + res.getResponseCode() + ']: ' + res.getContentText().slice(0, 500));
    try {
      var body = JSON.parse(res.getContentText());
      var items = Array.isArray(body.result) ? body.result : (body.result ? [body.result] : []);
      items.forEach(function (u) {
        var m = u.message || {};
        var from = m.from || {};
        var chat = m.chat || {};
        if (chat.id && found.indexOf(chat.id) < 0) {
          found.push(chat.id);
          Logger.log('✅ chat_id: ' + chat.id + (from.display_name ? '  (từ: ' + from.display_name + ')' : ''));
        }
      });
    } catch (e) {}
    if (!found.length) Utilities.sleep(1500);
  }
  if (found.length) {
    Logger.log('→ Copy chat_id ở trên, dán vào dòng: var ZALO_CHAT_IDS = [\'' + found[0] + '\'];');
  } else {
    Logger.log('❌ Hết 90 giây chưa thấy tin nào. Chạy lại hàm này rồi nhắn bot NGAY khi thấy dòng "Đang chờ".');
  }
}

/** Chạy 1 lần để cấp quyền + test: ghi Sheet OK, email OK, Zalo OK (nếu đã điền token/chat_id) */
function testSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('OK — kết nối được Sheet: ' + ss.getName());
  if (NOTIFY_EMAIL) MailApp.sendEmail(NOTIFY_EMAIL, '✅ Test webhook LP Dongwha', 'Apps Script đã sẵn sàng nhận đơn.');
  sendZalo('✅ Test: Zalo Bot đã kết nối với webhook LP Dongwha. Mỗi đơn mới sẽ báo vào đây.');
}

function formatProducts(s) {
  var names = { 'hong-sam': 'Hồng Sâm Dongwha', 'gas-whal': 'Gas Whal', 'sangsangton': 'SangSangTon Up' };
  return String(s).split(',').filter(Boolean).map(function (p) {
    var a = p.split(':');
    return (names[a[0]] || a[0]) + ' x' + (a[1] || 1);
  }).join(', ');
}
