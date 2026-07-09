/**
 * WEBHOOK NHẬN ĐƠN HÀNG — LP HỒNG SÂM DONGWHA
 * Form landing page → Google Sheet + Email + Zalo Bot.
 * Cách cài: xem HUONG-DAN-WEBHOOK.md (phần Zalo Bot ở cuối file đó)
 */

var SHEET_NAME = 'Đơn hàng';
var NOTIFY_EMAIL = '';   // Đã tắt email — chỉ báo qua Zalo. Muốn bật lại: điền email vào đây

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
    var products = formatProducts(data.products || ''); // bump Gas Whal đã nằm sẵn trong chuỗi products
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
    sendMetaCapi(data); // CAPI server-side — cùng event_id với Pixel, Meta tự khử trùng

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

// ═══════════════ BÁO CÁO TỰ ĐỘNG 22H ═══════════════
var TZ = 'Asia/Ho_Chi_Minh';
var PROJECT_NAME = 'LP Hồng Sâm DONGWHA';

/** CHẠY HÀM NÀY 1 LẦN để bật lịch báo cáo 22h mỗi tối */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'nightlyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightlyReport').timeBased().everyDays(1).atHour(22).nearMinute(0).inTimezone(TZ).create();
  Logger.log('✅ Đã hẹn giờ báo cáo 22h mỗi tối (giờ VN). Chạy thử ngay: chọn hàm nightlyReport → Run');
}

/** Chạy tự động 22h: báo ngày + tuần (CN) + tháng (ngày cuối tháng) */
function nightlyReport() {
  var now = new Date();
  var today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  var d0 = new Date(today + 'T00:00:00+07:00');

  sendZalo(buildReport(d0, now, '📊 BÁO CÁO CUỐI NGÀY (' + Utilities.formatDate(now, TZ, 'dd/MM/yyyy') + ')'));

  var dow = Number(Utilities.formatDate(now, TZ, 'u')); // 1=Thứ2 ... 7=CN
  if (dow === 7) {
    var wStart = new Date(d0.getTime() - 6 * 86400000);
    sendZalo(buildReport(wStart, now, '🗓️ BÁO CÁO TUẦN (' + Utilities.formatDate(wStart, TZ, 'dd/MM') + ' – ' + Utilities.formatDate(now, TZ, 'dd/MM/yyyy') + ')'));
  }

  var tomorrow = new Date(now.getTime() + 86400000);
  if (Utilities.formatDate(tomorrow, TZ, 'dd') === '01') {
    var mStart = new Date(Utilities.formatDate(now, TZ, 'yyyy-MM') + '-01T00:00:00+07:00');
    sendZalo(buildReport(mStart, now, '📅 BÁO CÁO THÁNG ' + Utilities.formatDate(now, TZ, 'MM/yyyy')));
  }
}

// ==== META CONVERSIONS API — bắn Lead đường server, dedup với Pixel qua event_id ====
var META_PIXEL_ID = '1353204000286435';
var META_CAPI_TOKEN = 'EAAYpr58KsFIBR5LniGLiPPhB0fblXE0h7I1uHfWJz0REZBxFisvO2fJeeHyV1xED5WCyKChTFFCexGeYu4lVZAwKMfcIZAHVDOFTUWsrEv8UNVQyitn6vZB9jZBnWlKaMyQZASCKxSsmQew84Yvd9CtahIaZBaQGkUKnMYGiPTliElv88159y1LVoTbBKOn9Pz0fgZDZD';

function sha256hex(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('');
}

function sendMetaCapi(data) {
  if (!META_CAPI_TOKEN) return null;
  try {
    var ph = String(data.phone || '').replace(/\D/g, '');
    if (ph.charAt(0) === '0') ph = '84' + ph.slice(1);
    var payload = { data: [{
      event_name: 'Lead',
      event_time: Number(data.event_time) || Math.floor(Date.now() / 1000),
      event_id: data.event_id || '',
      action_source: 'website',
      event_source_url: data.event_source_url || '',
      user_data: {
        ph: ph ? [sha256hex(ph)] : undefined,
        client_ip_address: data.client_ip_address || undefined,
        client_user_agent: data.client_user_agent || undefined,
        fbc: data.fbc || undefined,
        fbp: data.fbp || undefined
      },
      custom_data: { value: Number(data.total_amount) || 0, currency: 'VND', content_name: 'Dongwha Red Ginseng' }
    }] };
  