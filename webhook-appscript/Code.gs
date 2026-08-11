/**
 * WEBHOOK NHẬN ĐƠN HÀNG — LP HỒNG SÂM DONGWHA
 * Form landing page → Google Sheet + Email + Zalo Bot.
 * Cách cài: xem HUONG-DAN-WEBHOOK.md (phần Zalo Bot ở cuối file đó)
 */

var SHEET_NAME = 'Đơn hàng';
var NOTIFY_EMAIL = '';   // Đã tắt email — chỉ báo qua Zalo. Muốn bật lại: điền email vào đây

// ==== SECRET/KEY — KHÔNG hard-code nữa, đọc từ Script Properties (Project Settings > Script Properties) ====
// Cách set: mở Apps Script editor → biểu tượng bánh răng "Project Settings" (bên trái) → mục
// "Script Properties" → "Add script property" → nhập đúng tên key ở dưới (SP_*) + giá trị thật.
// Toàn bộ key cũ từng bị hard-code trong file này (đã public trên GitHub) coi như đã lộ —
// PHẢI tạo key mới ở Meta/Zalo/CS-Cart rồi điền key MỚI vào đây, không dùng lại key cũ.
var SP = PropertiesService.getScriptProperties();
function reqProp(name) {
  var v = SP.getProperty(name);
  if (!v) dbg('⚠️ Thiếu Script Property "' + name + '" — vào Project Settings > Script Properties để điền.');
  return v || '';
}

// Nhật ký trên đám mây (Cloud Logging) của dự án này đang không hiển thị được (chờ mãi không lên) —
// nên ghi debug trực tiếp vào Sheet để chắc chắn nhìn thấy, không phụ thuộc Executions/Cloud Logging nữa.
var DEBUG_LOG = [];
function dbg(msg) { try { Logger.log(msg); } catch (e) {} DEBUG_LOG.push(String(msg)); }

// ==== ZALO BOT — token đọc từ Script Property "SP_ZALO_BOT_TOKEN" ====
var ZALO_BOT_TOKEN = reqProp('SP_ZALO_BOT_TOKEN');
var ZALO_CHAT_IDS  = ['3f4e22e5d1ab38f561ba'];   // Bình Lê. Thêm người nhận: ['id1','id2']

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    var headers = ['Thời gian', 'Mã đơn', 'Họ tên', 'SĐT', 'Địa chỉ', 'Sản phẩm', 'Tổng SL',
      'Bump hộp 2', 'Thanh toán', 'Tổng tiền', 'Ghi chú', 'Trạng thái gọi',
      'UTM Source', 'UTM Medium', 'UTM Campaign', 'UTM Content', 'UTM Term', 'Ref', 'fbclid', 'Event ID', 'URL nguồn',
      'Mã đơn hệ thống', 'Đã nhận tiền (CK)', 'Products raw (ẩn)',
      'Tỉnh/Thành', 'Mã Tỉnh', 'Quận/Huyện', 'Mã Quận', 'Debug (CS-Cart)', 'GA Client ID'];
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d1fae5');
      sh.setFrozenRows(1);
    }
    ensureExtraColumns(sh); // phòng khi sheet cũ chưa có 3 cột mới (Mã đơn hệ thống / Đã nhận tiền / Products raw)

    var payLabel = (data.payment_method === 'cod') ? 'COD' : 'Chuyển khoản';
    var products = formatProducts(data.products || ''); // bump Gas Whal đã nằm sẵn trong chuỗi products
    var total = Number(data.total_amount || 0).toLocaleString('vi-VN') + 'đ';

    sh.appendRow([
      new Date(), data.code || '', data.fullname || '', "'" + (data.phone || ''), data.address || '',
      products, data.quantity || '',
      data.order_bump ? 'Có' : 'Không', payLabel, data.total_amount || '',
      data.note || '', (data.payment_method === 'cod' ? 'Chưa gọi' : 'Chờ chuyển khoản'),
      data.utm_source || '', data.utm_medium || '', data.utm_campaign || '',
      data.utm_content || '', data.utm_term || '', data.ref || '', data.fbclid || '',
      data.event_id || '', data.event_source_url || '', '', '', data.products || '',
      data.province || '', data.province_code || '', data.district || '', data.district_code || '',
      '', data.ga_client_id || ''
    ]);

    // COD → đẩy ngay vào hệ thống trungsoncare.com (dược sĩ xử lý luôn trong đó, khỏi gõ tay lại).
    // CK → CHƯA đẩy — chỉ đẩy khi dược sĩ tick "Đã nhận tiền (CK)" (xem pushConfirmedBankOrders, chạy mỗi 10 phút).
    var newOrderId = null; // khai báo ngoài if để dùng được ở response JSON trả về cho landing page (hiện mã đơn thật)
    if (data.payment_method === 'cod') {
      DEBUG_LOG.length = 0;
      try {
        newOrderId = tscCreateOrder(data);
      } catch (pushErr) {
        dbg('❌ tscCreateOrder EXCEPTION: ' + pushErr + (pushErr && pushErr.stack ? ('\n' + pushErr.stack) : ''));
      }
      if (newOrderId) sh.getRange(sh.getLastRow(), colIndex(sh, 'Mã đơn hệ thống')).setValue(newOrderId);
      var debugCol = colIndex(sh, 'Debug (CS-Cart)');
      if (debugCol > 0) sh.getRange(sh.getLastRow(), debugCol).setValue(DEBUG_LOG.join(' | ').slice(0, 5000));
    }

    var msg = '🛒 ĐƠN MỚI: [' + payLabel + ']' +
      '\n👤 Tên khách hàng: ' + (data.fullname || '') +
      '\n📞 SĐT: ' + (data.phone || '') +
      '\n📍 Địa chỉ: ' + (data.address || '') +
      '\n📦 Đặt hàng: ' + products +
      '\n💵 Thanh toán: ' + total +
      '\n🧾 Mã đơn: ' + (data.code || '') +
      (data.note ? '\n📝 Ghi chú: ' + data.note : '') +
      '\n\n→ GỌI XÁC NHẬN TRONG 5 PHÚT!';

    // KHÔNG báo Zalo khi có lead mới (COD lẫn CK) — chỉ ghi Sheet.
    // Thông báo sẽ do syncOrderStatus bắn khi hệ thống admin XÁC NHẬN đơn.
    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, '🛒 ĐƠN MỚI [' + payLabel + '] ' + (data.fullname || '') + ' – ' + total, msg);
    }
    sendMetaCapi(data); // CAPI server-side — cùng event_id với Pixel, Meta tự khử trùng

    // order_id: chỉ có với COD (tạo đơn ngay) — landing page dùng để hiện mã đơn CS-Cart thật thay cho mã tạm.
    return ContentService.createTextOutput(JSON.stringify({ ok: true, order_id: newOrderId || null }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('❌ doPost LỖI: ' + err + (err && err.stack ? ('\n' + err.stack) : ''));
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

// ═══════════════ ĐỒNG BỘ TRẠNG THÁI ĐƠN MỖI 10 PHÚT ═══════════════

/** CHẠY 1 LẦN để bật lịch đối soát trạng thái mỗi 10 phút */
function setupSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncOrderStatus') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncOrderStatus').timeBased().everyMinutes(10).create();
  Logger.log('✅ Đã bật đồng bộ trạng thái mỗi 10 phút. Chạy thử ngay: chọn syncOrderStatus → Run');
}

// Mã trạng thái thật của trungsoncare.com (lấy từ statuses?type=O — xem exploreTscRefData):
//   O=Đơn hàng mới  A=Đặt trước  P=Đã thanh toán  Y=Xác nhận đơn  C=Hoàn tất
//   B=Đơn hàng đặt trước  D=Hủy đơn  F=Thanh toán thất bại  I=Khách hủy đơn
//   N=Không đầy đủ (trạng thái ẩn, không nằm trong danh sách statuses?type=O, chỉ thấy qua GET đơn lẻ)
// Chỉ 2 trạng thái Y/P mới coi là "đã xác nhận" và bắn Zalo — đúng yêu cầu, KHÔNG tự phân chia/gán sale nào cả,
// chỉ theo dõi trạng thái đơn. C (Hoàn tất) cũng tính là xác nhận phòng khi đơn nhảy thẳng qua Hoàn tất.
var TSC_CONFIRM_CODES = ['Y', 'P', 'C'];
// Các trạng thái coi là "không tính" — chỉ note "Hủy", KHÔNG báo Zalo.
var TSC_CANCEL_CODES = ['D', 'I', 'N', 'F'];

/**
 * Chạy mỗi 10 phút: quét đơn hệ thống (trungsoncare.com/admin1906.php?dispatch=orders.manage) 3 ngày gần nhất, khớp SĐT với Sheet.
 * - Dòng "Chưa gọi"/"Chờ chuyển khoản" có đơn hệ thống Xác nhận đơn/Đã thanh toán (Y/P, hoặc Hoàn tất-C) → đổi thành "Xác nhận đơn ✓" + báo Zalo
 * - Đơn hệ thống là Hủy đơn/Khách hủy đơn/Không đầy đủ/Thanh toán thất bại (D/I/N/F) → đổi thành "Hủy (hệ thống)", KHÔNG báo Zalo
 * - Đơn CK vừa được xác nhận → lúc này MỚI bắn thông báo Zalo (đúng luật: CK chỉ tính khi tiền đã về)
 * - KHÔNG tự gán/phân chia sale nào xử lý — chỉ theo dõi trạng thái đơn trên hệ thống, sale nào xử lý là việc của các dược sĩ.
 * Không đè lên ô mà dược sĩ đã tự ghi chú khác.
 */
function syncOrderStatus() {
  pushConfirmedBankOrders(); // đẩy đơn CK mà dược sĩ đã tick "Đã nhận tiền" vào hệ thống trước, rồi mới đối soát bên dưới
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return;

  var now = new Date(), from = new Date(now.getTime() - 3 * 86400000);
  var byPhone = {};
  tscOrders(from, now).forEach(function (o) {
    var p = last9(o.phone || o.b_phone || o.s_phone);
    if (!p) return;
    var code = String(o.status || '');
    if (TSC_CONFIRM_CODES.indexOf(code) >= 0) byPhone[p] = { st: 'ok', total: o.total };
    else if (TSC_CANCEL_CODES.indexOf(code) >= 0 && !byPhone[p]) byPhone[p] = { st: 'huy' };
  });

  var last = sh.getLastRow();
  var start = Math.max(2, last - 300); // chỉ quét ~300 dòng gần nhất cho nhẹ
  var vals = sh.getRange(start, 1, last - start + 1, 30).getValues(); // mở rộng tới cột 30 để lấy "Mã đơn hệ thống" (22) và "GA Client ID" (30)
  var changed = 0;
  vals.forEach(function (r, i) {
    if (!(r[0] instanceof Date)) return;
    var cur = String(r[11] || '');
    if (cur !== 'Chưa gọi' && cur !== 'Chờ chuyển khoản') return; // không đè ghi chú tay
    var p = last9(r[3]);
    var hit = p ? byPhone[p] : null;

    // orders?period=C (danh sách hàng loạt) KHÔNG trả về đơn đang ở trạng thái ẩn "N" (Không đầy đủ) —
    // trạng thái này chỉ thấy được khi GET trực tiếp theo order_id (đã xác nhận qua kiểm tra thủ công).
    // Vì vậy nếu không khớp được qua danh sách (byPhone) nhưng dòng đã có "Mã đơn hệ thống" thật (cột V),
    // check trực tiếp đơn đó trước khi bỏ qua — đây là chỗ trước đây bị bỏ sót khiến đơn "Không đầy đủ"
    // không bao giờ tự chuyển "Hủy đơn" trong Sheet dù trigger vẫn chạy đều, không báo lỗi gì.
    if (!hit) {
      var orderId = r[21];
      if (orderId) {
        var single = tscFetch('orders/' + orderId);
        if (single) {
          var code2 = String(single.status || '');
          if (TSC_CONFIRM_CODES.indexOf(code2) >= 0) hit = { st: 'ok', total: single.total };
          else if (TSC_CANCEL_CODES.indexOf(code2) >= 0) hit = { st: 'huy' };
        }
      }
    }
    if (!hit) return;
    if (hit.st === 'ok') {
      sh.getRange(start + i, 12).setValue('Xác nhận đơn ✓');
      changed++;
      // Đơn chính thức được xác nhận trên hệ thống → lúc này mới báo Zalo (cả COD lẫn CK)
      sendZalo('✅ ĐƠN ĐÃ XÁC NHẬN: [' + (String(r[8]).indexOf('COD') >= 0 ? 'COD' : 'Chuyển khoản') + ']' +
        '\n👤 Tên khách hàng: ' + (r[2] || '') +
        '\n📞 SĐT: ' + (r[3] || '') +
        '\n📍 Địa chỉ: ' + (r[4] || '') +
        '\n📦 Đặt hàng: ' + (r[5] || '') +
        '\n💵 Thanh toán: ' + Number(hit.total || r[9] || 0).toLocaleString('vi-VN') + 'đ' +
        '\n🧾 Mã đơn: ' + (r[1] || ''));
      // Bắn Purchase thật (server-side, Meta CAPI + GA4 Measurement Protocol) đúng lúc đơn được XÁC NHẬN trên hệ thống —
      // điểm xác thực duy nhất cho cả COD lẫn CK, tránh báo Purchase khi khách mới chỉ tự nhận (chưa chắc đã trả tiền).
      try {
        sendMetaCapi({
          phone: r[3],
          event_time: Math.floor(Date.now() / 1000),
          event_id: 'purchase_' + (r[1] || '') + '_' + Date.now(),
          event_source_url: r[20] || '',
          total_amount: hit.total || r[9] || 0
        }, 'Purchase');
      } catch (capiErr) {
        dbg('⚠️ sendMetaCapi Purchase lỗi: ' + capiErr);
      }
      try {
        sendGA4Purchase({
          ga_client_id: r[29] || '',
          code: r[1] || '',
          total_amount: hit.total || r[9] || 0
        });
      } catch (ga4Err) {
        dbg('⚠️ sendGA4Purchase lỗi: ' + ga4Err);
      }
    } else if (hit.st === 'huy') {
      sh.getRange(start + i, 12).setValue('Hủy đơn');
      changed++;
    }
  });
  if (changed) Logger.log('Đã cập nhật ' + changed + ' dòng theo hệ thống.');
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

// ==== META CONVERSIONS API — bắn Lead/Purchase đường server, dedup với Pixel qua event_id ====
var META_PIXEL_ID = '1353204000286435';
var META_CAPI_TOKEN = reqProp('SP_META_CAPI_TOKEN'); // Script Property "SP_META_CAPI_TOKEN"

function sha256hex(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('');
}

function sendMetaCapi(data, eventName) {
  if (!META_CAPI_TOKEN) return null;
  try {
    var ph = String(data.phone || '').replace(/\D/g, '');
    if (ph.charAt(0) === '0') ph = '84' + ph.slice(1);
    var payload = { data: [{
      event_name: eventName || 'Lead',
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
    var res = UrlFetchApp.fetch('https://graph.facebook.com/v21.0/' + META_PIXEL_ID + '/events?access_token=' + META_CAPI_TOKEN, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    return res.getContentText();
  } catch (e) { return String(e); }
}

/** Chạy hàm này để test CAPI — log phải ra {"events_received":1,...} */
function testMetaCapi() {
  var r = sendMetaCapi({
    phone: '0900000001', event_time: Math.floor(Date.now() / 1000),
    event_id: 'test_' + Date.now(), event_source_url: 'https://dongwha-hong-sam.vercel.app/',
    total_amount: 356000, payment_method: 'cod'
  });
  Logger.log('Meta trả về: ' + r);
}

// ==== GA4 MEASUREMENT PROTOCOL — bắn Purchase server-side, song song với Meta CAPI ====
var GA4_MEASUREMENT_ID = reqProp('SP_GA4_MEASUREMENT_ID'); // vd: G-XXXXXXXXXX
var GA4_API_SECRET = reqProp('SP_GA4_API_SECRET');         // tạo ở GA4 > Admin > Data Streams > Measurement Protocol API secrets

/** ga_client_id lấy từ cookie _ga phía landing page (định dạng GAx.y.XXXXXXXXXX.YYYYYYYYYY).
 * Nếu không có (khách chặn cookie/không cài GA4 tag) thì tự sinh 1 client_id ngẫu nhiên —
 * GA4 vẫn nhận event nhưng sẽ tính thành session mới, không nối được vào hành trình duyệt web gốc của khách. */
function sendGA4Purchase(data) {
  if (!GA4_MEASUREMENT_ID || !GA4_API_SECRET) return null;
  try {
    var clientId = data.ga_client_id || (Math.floor(Math.random() * 2147483647) + '.' + Math.floor(Date.now() / 1000));
    var payload = {
      client_id: clientId,
      events: [{
        name: 'purchase',
        params: {
          transaction_id: data.code || data.event_id || '',
          currency: 'VND',
          value: Number(data.total_amount) || 0,
          items: [{ item_name: 'Dongwha Red Ginseng', quantity: 1, price: Number(data.total_amount) || 0 }]
        }
      }]
    };
    var url = 'https://www.google-analytics.com/mp/collect?measurement_id=' + GA4_MEASUREMENT_ID + '&api_secret=' + GA4_API_SECRET;
    var res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    return res.getResponseCode(); // GA4 mp/collect luôn trả 204 nếu nhận được, không có body — không dùng để debug nội dung được
  } catch (e) { return String(e); }
}

/** Chạy hàm này để test GA4 — không có response body để đọc (GA4 luôn trả 204 rỗng),
 * muốn xem event có vào không thì đổi endpoint sang /debug/mp/collect và đọc log validationMessages. */
function testGA4Purchase() {
  var code = sendGA4Purchase({
    ga_client_id: '', code: 'TEST-' + Date.now(), total_amount: 356000
  });
  Logger.log('GA4 response code: ' + code);
}

// ==== TRUNG SƠN CARE API — đối soát đơn thật trên hệ thống ====
var TSC_API_EMAIL = reqProp('SP_TSC_API_EMAIL'); // Script Property "SP_TSC_API_EMAIL"
var TSC_API_KEY = reqProp('SP_TSC_API_KEY');     // Script Property "SP_TSC_API_KEY"
// Trạng thái được tính là ĐƠN CHỐT (viết thường, khớp theo tên trong admin)
var TSC_CONFIRM_STATUSES = ['xác nhận đơn', 'hoàn tất', 'đã thanh toán'];

// ==== ĐẨY ĐƠN TỰ ĐỘNG VÀO admin1906.php?dispatch=orders.manage ====
// ⚠️ CHƯA DÙNG ĐƯỢC — phải điền đủ các ID dưới đây trước.
// Cách lấy: chạy hàm exploreTscRefData() 1 lần, xem Execution log, rồi điền vào đây.
var TSC_COMPANY_ID = '1';        // ✅ Bắt buộc khi POST tạo đơn — API trả lỗi 400 "Xin vui lòng chọn một cửa hàng trước" nếu để trống.
var TSC_PAYMENT_ID_COD = '6';    // ✅ "Tiền mặt (C.O.D)" — status Active
var TSC_PAYMENT_ID_BANK = '12';  // ⚠️ "Chuyển khoản" — hệ thống đang để status D (Disabled)! Cần quyết định trước khi bật (xem câu hỏi bên dưới)
var TSC_SHIPPING_ID = '9';       // ✅ "Giao hàng tận nơi" — status Active, phí cứng 25.000đ, giới hạn riêng cho nhóm
                                  // người dùng "Landing Page - Noi bo (API)" (usergroup_id=10) nên KHÔNG hiện ở checkout
                                  // thật trên toàn trang trungsoncare.com (khách vãng lai/đăng ký không thuộc nhóm này),
                                  // nhưng vẫn gán được cho đơn tạo qua API vì gán trực tiếp bằng shipping_id không bị lọc
                                  // theo nhóm người dùng. (Đã thử id=10 với trạng thái "Ẩn" trước đó — CS-Cart âm thầm từ
                                  // chối và fallback về shipping mặc định id=1 (18k) khi tạo đơn qua API — KHÔNG dùng nữa.)
                                  // id=7 đã Tắt (đơn giản là bản đặt sai tên trước đây, không dùng). id=8 lỗi cũ đã tự Disabled.
var LP_SHIPPING_FEE = 25000;     // Phí ship landing page báo khách — dùng để so sánh/log, đối chiếu total.
var TSC_PRODUCT_ID_MAP = {       // slug trong chuỗi "products" của form → product_id thật bên hệ thống
  'hong-sam': '6372',    // ✅ "NƯỚC UỐNG HỒNG SÂM DONGWHA" — giá 203,000đ khớp, status Active
  'gas-whal': '6373',    // ✅ "Nước uống thảo mộc có ga Dongwha Gas Whal 75ml" — giá 210,000đ khớp, status Active
  'sangsangton': '7851'  // ✅ "NƯỚC UỐNG TĂNG LỰC DONGWHA SANGSANGTON UP" — giá 168,000đ khớp, status Active
  // Lưu ý: mỗi sản phẩm có 2-3 product_id trùng tên (bản disabled, hoặc giá bằng 1/10 — có vẻ là giá/chai lẻ
  // thay vì giá/hộp 10 chai). Đã chọn đúng bản khớp giá hộp trên landing page + status Active.
};
var TSC_GIFT_PRODUCT_ID = '7864'; // "Áo mưa Trung Sơn - SangSangTon" (SKU 934509) — giá gốc 99.000đ, dùng làm quà tặng (giá ghi đè = 0đ)
var TSC_GIFT_THRESHOLD = 549000;  // Hóa đơn landing page (đã gồm giảm giá, CHƯA gồm ship) >= mốc này thì tự động tặng kèm áo mưa

// Ghi chú: 5 dược sĩ xử lý đơn (Huỳnh Hồng Nhung, Lê Hải Vân, Võ Thị Bích Trâm, Bùi Hồng Ni,
// Nguyễn Trần Nguyệt Quế) — chỉ để tham khảo, KHÔNG tự động gán/phân chia issuer_id qua code.
// Cột "Trạng thái gọi" trên Sheet + trạng thái đơn bên hệ thống mới là thứ cần theo dõi.

function tscFetch(path) {
  try {
    var res = UrlFetchApp.fetch('https://trungsoncare.com/api/' + path, {
      headers: { Authorization: 'Basic ' + Utilities.base64Encode(TSC_API_EMAIL + ':' + TSC_API_KEY) },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText());
  } catch (e) { return null; }
}

/** POST tới API trungsoncare — dùng để tạo/sửa đơn. Trả về JSON response hoặc null nếu lỗi. */
function tscApiPost(path, payload, method) {
  try {
    var res = UrlFetchApp.fetch('https://trungsoncare.com/api/' + path, {
      method: method || 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Basic ' + Utilities.base64Encode(TSC_API_EMAIL + ':' + TSC_API_KEY) },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var bodyText = res.getContentText();
    if (code !== 200 && code !== 201) {
      dbg('⚠️ tscApiPost [' + path + '] lỗi ' + code + ': ' + bodyText.slice(0, 800));
      return null;
    }
    dbg('ℹ️ tscApiPost [' + path + '] OK ' + code + ': ' + bodyText.slice(0, 800));
    return JSON.parse(bodyText);
  } catch (e) { dbg('⚠️ tscApiPost exception: ' + e); return null; }
}

/**
 * Tạo 1 đơn hàng thật bên hệ thống (CS-Cart REST API — xem docs.cs-cart.com/latest/developer_guide/api/entities/orders.html).
 * data = payload gốc từ form landing page (giống data trong doPost).
 * Trả về order_id (string) nếu thành công, null nếu lỗi hoặc config chưa điền đủ.
 */
function tscCreateOrder(data) {
  if (!TSC_PAYMENT_ID_COD || !TSC_PAYMENT_ID_BANK || !TSC_SHIPPING_ID) {
    dbg('⚠️ Chưa điền đủ TSC_PAYMENT_ID_COD/BANK/SHIPPING_ID — bỏ qua tạo đơn. Chạy exploreTscRefData() để lấy ID.');
    return null;
  }

  var products = {};
  var idx = 1;
  String(data.products || '').split(',').filter(Boolean).forEach(function (p) {
    var a = p.split(':');
    var pid = TSC_PRODUCT_ID_MAP[a[0]];
    if (!pid) { dbg('⚠️ Không tìm thấy product_id cho slug "' + a[0] + '" — bỏ qua đơn.'); return; }
    products[idx++] = { product_id: String(pid), amount: String(a[1] || 1) };
  });
  if (!Object.keys(products).length) { dbg('⚠️ Không map được sản phẩm nào (data.products="' + (data.products || '') + '") — bỏ qua tạo đơn.'); return null; }

  // Tách họ tên: nếu khách chỉ gõ 1 chữ (VD "chi"), KHÔNG lặp lại chữ đó vào cả 2 ô — chỉ để 1 ô.
  var nameParts = String(data.fullname || 'Khach hang').trim().split(/\s+/).filter(Boolean);
  var firstname, lastname;
  if (nameParts.length <= 1) {
    firstname = nameParts[0] || 'Khach hang';
    lastname = '.'; // CS-Cart yêu cầu lastname khác rỗng — dùng dấu chấm thay vì lặp lại tên
  } else {
    lastname = nameParts.pop();
    firstname = nameParts.join(' ');
  }
  var phone = String(data.phone || '').replace(/\D/g, '');
  var fakeEmail = (phone || 'khach') + '@khachhang.local'; // form LP không thu email — bắt buộc phải có email nên tạo email giả theo SĐT

  // Tổng tiền đúng theo landing page báo khách. Nhờ dùng shipping_id riêng (TSC_SHIPPING_ID=9, phí
  // cứng 25k) nên CS-Cart tự tính Tổng cộng đúng luôn, không cần ghi đè gì cả — "total" dưới đây chỉ
  // còn là lớp an toàn dự phòng (đề phòng giá landing page lệch giá catalog vì lý do khác).
  var quotedTotal = Number(data.total_amount || 0);
  // "notes" = ghi chú CỦA KHÁCH (hiện ở ô "Khách hàng ghi chú") — CHỈ để đúng những gì khách tự gõ,
  // để trống nếu khách không ghi gì, không nhét thông tin hệ thống vào đây.
  var customerNote = data.note || '';
  // Không ghi chú nội bộ nào vào đơn nữa (đã bỏ theo yêu cầu) — mã đơn LP vẫn lưu đủ trong Google Sheet rồi.

  // Quà tặng áo mưa: landing page đã hứa "bill từ 549k tặng áo mưa cao cấp" (xem gift-bar trên trang).
  // total_amount = subTotal + phí ship, nhưng phí ship chỉ tính khi subTotal < 300.000đ (đã cấu hình free ship
  // ở CS-Cart) — nghĩa là khi subTotal >= 549.000đ thì ship luôn = 0, nên total_amount lúc đó CHÍNH LÀ subTotal.
  // => so sánh thẳng quotedTotal với TSC_GIFT_THRESHOLD là tương đương với logic subTotal trên landing page.
  if (quotedTotal >= TSC_GIFT_THRESHOLD) {
    products[idx++] = { product_id: TSC_GIFT_PRODUCT_ID, amount: '1', price: '0' };
    dbg('🎁 Hóa đơn ' + quotedTotal + 'đ >= ' + TSC_GIFT_THRESHOLD + 'đ — tự động thêm quà tặng Áo mưa Trung Sơn (product_id ' + TSC_GIFT_PRODUCT_ID + ', giá 0đ).');
  }

  var payload = {
    user_id: '0',
    company_id: String(TSC_COMPANY_ID || '1'), // bắt buộc — xem comment ở TSC_COMPANY_ID phía trên
    payment_id: String(data.payment_method === 'cod' ? TSC_PAYMENT_ID_COD : TSC_PAYMENT_ID_BANK),
    shipping_id: String(TSC_SHIPPING_ID),
    products: products,
    user_source: 'LandingPage', // để cột "Source" trong admin không bị trống — phân biệt với đơn Website/MobileApp thật
    notes: customerNote, // ghi chú CỦA KHÁCH — để trống nếu khách không ghi gì
    total: quotedTotal > 0 ? String(quotedTotal) : undefined, // lớp an toàn dự phòng — bình thường không cần vì shipping_id=7 đã tự tính đúng
    // Ô "Quản lý" trên đơn (issuer_id) mặc định bị CS-Cart tự gán = tài khoản API (binhlct) vì đó là tài khoản
    // xác thực khi gọi API — không phải do mình cố ý gán. Thử để trống/0 để ô đó thành placeholder, tránh
    // hiện sẵn tên admin, cho sale nào cũng bấm vào tự nhận xử lý được (KHÔNG tự phân chia cho ai cụ thể).
    issuer_id: '0',
    user_data: {
      email: fakeEmail,
      firstname: firstname, lastname: lastname,
      b_firstname: firstname, b_lastname: lastname,
      s_firstname: firstname, s_lastname: lastname,
      b_address: data.address || '', s_address: data.address || '',
      // Mã tỉnh/quận GSO chuẩn (khớp đúng hệ thống tslocation của trungsoncare.com) —
      // lấy từ select Tỉnh/Thành + Quận/Huyện trên landing page (vn-locations.json).
      // Hệ thống dùng mã số (state/district), KHÔNG dùng city dạng chữ → để trống b_city.
      // Phường/xã (ward) vẫn chưa có select riêng — khách gõ tự do trong ô "Địa chỉ chi tiết".
      b_city: '', s_city: '',
      b_district: String(data.district_code || ''), s_district: String(data.district_code || ''),
      b_ward: '', s_ward: '',
      b_country: 'VN', s_country: 'VN',
      b_state: String(data.province_code || ''), s_state: String(data.province_code || ''),
      b_zipcode: '', s_zipcode: '',
      b_phone: phone, s_phone: phone
    }
  };

  dbg('▶️ tscCreateOrder: gửi payload ' + JSON.stringify(payload).slice(0, 800));
  var path = TSC_COMPANY_ID ? ('stores/' + TSC_COMPANY_ID + '/orders') : 'orders';
  var res = tscApiPost(path, payload);
  if (res && res.order_id) {
    dbg('✅ Đã tạo đơn hệ thống #' + res.order_id + ' cho ' + (data.fullname || ''));
    dbg('ℹ️ issuer_id sau khi tạo = ' + (res.order_data && res.order_data.issuer_id) + ' (gửi issuer_id:"0" — nếu vẫn khác "0"/null nghĩa là CS-Cart bỏ qua field này, tự gán tài khoản API).');
    // Kiểm tra xem "total" và "shipping_cost" gửi kèm lúc TẠO đơn có được CS-Cart chấp nhận không.
    var createdTotal = res.order_data && res.order_data.total != null ? Number(res.order_data.total) : null;
    var createdShipCost = res.order_data && res.order_data.shipping_cost != null ? Number(res.order_data.shipping_cost) : null;
    var totalOk = quotedTotal > 0 && createdTotal !== null && Math.round(createdTotal) === Math.round(quotedTotal);
    var shipOk = createdShipCost !== null && Math.round(createdShipCost) === LP_SHIPPING_FEE;
    if (totalOk && shipOk) {
      dbg('✅ Total = ' + createdTotal + 'đ, Phí ship = ' + createdShipCost + 'đ — đã đúng ngay lúc tạo đơn, không cần sửa thêm.');
    } else {
      dbg('⚠️ Lúc tạo đơn: total=' + createdTotal + 'đ (LP báo ' + quotedTotal + 'đ), ship=' + createdShipCost + 'đ (LP báo ' + LP_SHIPPING_FEE + 'đ). Thử sửa lại bằng PUT (dự phòng — PUT có thể bị server chặn)...');
      // PUT thường bị chặn 403 ở tầng Apache trên host này — chỉ thử 1 lần cho chắc, không cần retry
      // nhiều lần vì đây là chặn hạ tầng chứ không phải lỗi tạm thời. Đã ghi rõ total đúng vào "notes"
      // ở trên rồi nên dù PUT thất bại, dược sĩ vẫn thấy số đúng để tự sửa tay khi xử lý đơn.
      var updatePayload = {};
      if (!totalOk && quotedTotal > 0) updatePayload.total = String(quotedTotal);
      if (!shipOk) updatePayload.shipping_cost = String(LP_SHIPPING_FEE);
      var updateRes = tscApiPost('orders/' + res.order_id, updatePayload, 'put');
      if (updateRes) dbg('✅ Đã sửa lại đơn #' + res.order_id + ' bằng PUT: ' + JSON.stringify(updatePayload));
      else dbg('⚠️ PUT sửa đơn #' + res.order_id + ' KHÔNG thành công (có thể bị chặn ở server) — nhưng "notes" của đơn đã ghi rõ số đúng để dược sĩ tự sửa tay.');
    }
    return res.order_id;
  }
  dbg('⚠️ tscCreateOrder: API trả về nhưng KHÔNG có order_id. res=' + JSON.stringify(res));
  return null;
}

/** Đảm bảo Sheet có đủ 3 cột mới (idempotent — gọi nhiều lần vô hại) */
function ensureExtraColumns(sh) {
  var wanted = ['Mã đơn hệ thống', 'Đã nhận tiền (CK)', 'Products raw (ẩn)',
    'Tỉnh/Thành', 'Mã Tỉnh', 'Quận/Huyện', 'Mã Quận', 'Debug (CS-Cart)', 'GA Client ID'];
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  wanted.forEach(function (name) {
    if (existing.indexOf(name) === -1) {
      var col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#d1fae5');
      existing.push(name);
    }
  });
}

/** Tìm số thứ tự cột theo tên header (dò trong row 1) */
function colIndex(sh, name) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var i = headerRow.indexOf(name);
  return i === -1 ? -1 : i + 1;
}

/**
 * Chạy mỗi 10 phút (gọi từ syncOrderStatus): quét các dòng CK mà dược sĩ đã tick "Đã nhận tiền (CK)"
 * nhưng CHƯA có "Mã đơn hệ thống" → tạo đơn thật bên hệ thống, ghi order_id vào Sheet.
 * Sau đó syncOrderStatus (chạy ngay sau) sẽ tiếp tục đối soát như bình thường.
 */
function pushConfirmedBankOrders() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return;
  ensureExtraColumns(sh);
  var colTrangThai = colIndex(sh, 'Trạng thái gọi');
  var colMaDon = colIndex(sh, 'Mã đơn');
  var colHoTen = colIndex(sh, 'Họ tên');
  var colSDT = colIndex(sh, 'SĐT');
  var colDiaChi = colIndex(sh, 'Địa chỉ');
  var colGhiChu = colIndex(sh, 'Ghi chú');
  var colMaHeThong = colIndex(sh, 'Mã đơn hệ thống');
  var colDaNhanTien = colIndex(sh, 'Đã nhận tiền (CK)');
  var colProductsRaw = colIndex(sh, 'Products raw (ẩn)');
  var colTinh = colIndex(sh, 'Tỉnh/Thành');
  var colMaTinh = colIndex(sh, 'Mã Tỉnh');
  var colQuan = colIndex(sh, 'Quận/Huyện');
  var colMaQuan = colIndex(sh, 'Mã Quận');
  var colDebug = colIndex(sh, 'Debug (CS-Cart)');

  var last = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(2, 1, last - 1, lastCol).getValues();
  var pushed = 0;
  vals.forEach(function (r, i) {
    var trangThai = String(r[colTrangThai - 1] || '');
    var daNhanTien = String(r[colDaNhanTien - 1] || '').trim();
    var maHeThong = String(r[colMaHeThong - 1] || '').trim();
    if (trangThai !== 'Chờ chuyển khoản' || !daNhanTien || maHeThong) return; // chỉ đơn CK, đã tick tiền về, chưa đẩy lần nào

    var data = {
      payment_method: 'bank', fullname: r[colHoTen - 1], phone: r[colSDT - 1],
      address: r[colDiaChi - 1], note: r[colGhiChu - 1], code: r[colMaDon - 1],
      products: colProductsRaw > 0 ? r[colProductsRaw - 1] : '',
      province_name: colTinh > 0 ? r[colTinh - 1] : '', province_code: colMaTinh > 0 ? r[colMaTinh - 1] : '',
      district_name: colQuan > 0 ? r[colQuan - 1] : '', district_code: colMaQuan > 0 ? r[colMaQuan - 1] : ''
    };
    DEBUG_LOG.length = 0;
    var orderId = null;
    try {
      orderId = tscCreateOrder(data);
    } catch (pushErr) {
      dbg('❌ tscCreateOrder EXCEPTION (CK): ' + pushErr + (pushErr && pushErr.stack ? ('\n' + pushErr.stack) : ''));
    }
    if (orderId) {
      sh.getRange(2 + i, colMaHeThong).setValue(orderId);
      pushed++;
    }
    if (colDebug > 0) sh.getRange(2 + i, colDebug).setValue(DEBUG_LOG.join(' | ').slice(0, 5000));
  });
  if (pushed) dbg('✅ Đã đẩy ' + pushed + ' đơn CK vào hệ thống.');
}

/** Map mã trạng thái (1 chữ cái) → tên trạng thái viết thường */
function tscStatusMap() {
  var data = tscFetch('statuses?type=O&items_per_page=100');
  var map = {};
  ((data && data.statuses) || []).forEach(function (s) {
    map[s.status] = String(s.description || '').toLowerCase();
  });
  return map;
}

/** Lấy đơn trên hệ thống trong khoảng thời gian */
function tscOrders(start, end) {
  var all = [], page = 1;
  while (page < 20) {
    var data = tscFetch('orders?period=C&time_from=' + Math.floor(start.getTime() / 1000) +
      '&time_to=' + Math.floor(end.getTime() / 1000) + '&items_per_page=250&page=' + page);
    var list = (data && data.orders) || [];
    all = all.concat(list);
    if (list.length < 250) break;
    page++;
  }
  return all;
}

function last9(phone) { var d = String(phone || '').replace(/\D/g, ''); return d.slice(-9); }

/** CHẠY 1 LẦN — bật lại payment method "Chuyển khoản" (payment_id 12) từ Disabled sang Active,
 * để đơn CK tạo qua API không bị hệ thống từ chối vì payment method "không available". */
function enableBankPaymentMethod() {
  var res = tscApiPost('payments/12', { status: 'A' }, 'put');
  if (res) Logger.log('✅ Đã bật Active cho payment_id=12 (Chuyển khoản). Kiểm tra lại trong Admin → Thanh toán để chắc ăn.');
  else Logger.log('❌ Không bật được — xem log lỗi phía trên (tscApiPost).');
}

/**
 * CHẠY 1 LẦN — API trungsoncare không có tham số tìm theo tên (q= bị bỏ qua, đó là lý do lần trước
 * exploreTscRefData ra cùng 1 sản phẩm cho cả 3 lần tìm). Đợt này dò đúng theo GIÁ (price_from/price_to
 * là filter thật, có trong docs CS-Cart) để tìm product_id của Hồng Sâm Dongwha (203000đ) và
 * SangSangTon Up (168000đ). Gas Whal đã biết là product_id 6373.
 */
function exploreTscProductsByPrice() {
  var prices = { 'Hồng Sâm Dongwha (203000đ)': 203000, 'SangSangTon Up (168000đ)': 168000 };
  Object.keys(prices).forEach(function (label) {
    var price = prices[label];
    var data = tscFetch('products?price_from=' + price + '&price_to=' + price + '&items_per_page=20');
    Logger.log('--- ' + label + ' ---');
    var list = (data && data.products) || [];
    if (!list.length) { Logger.log('❌ Không tìm thấy sản phẩm nào giá đúng ' + price + 'đ'); return; }
    list.forEach(function (p) {
      Logger.log('product_id=' + p.product_id + '  |  ' + p.product + '  |  price=' + p.price + '  |  status=' + p.status);
    });
  });
}

/**
 * Giá không khớp (giá LP có thể khác giá gốc lưu trong hệ thống) → quét toàn bộ catalog theo TÊN thay vì giá.
 * Tự dừng trước khi bị Google kill (giới hạn ~6 phút/lần chạy), nhớ đã quét tới trang nào — CHẠY LẠI HÀM NÀY
 * NHIỀU LẦN (chọn lại → Chạy) cho tới khi log in ra "ĐÃ QUÉT HẾT CATALOG". Log ra kết quả NGAY khi tìm thấy,
 * không đợi tới cuối, nên dù bị timeout giữa chừng cũng không mất kết quả.
 */
function exploreTscProductsByName() {
  var keywords = ['hồng sâm', 'dongwha', 'sangsangton', 'sáng sáng tôn', 'gas whal'];
  var props = PropertiesService.getScriptProperties();
  var page = Number(props.getProperty('EXPLORE_PAGE') || '1');
  var startTime = Date.now();
  var maxMs = 5 * 60 * 1000; // dừng ở phút thứ 5, chừa margin trước khi Google tự kill ở phút thứ 6

  Logger.log('▶️ Bắt đầu quét từ trang ' + page);
  while (Date.now() - startTime < maxMs) {
    var data = tscFetch('products?items_per_page=100&page=' + page + '&sort_by=product&sort_order=asc');
    var list = (data && data.products) || [];
    if (!list.length) {
      Logger.log('✅ ĐÃ QUÉT HẾT CATALOG (dừng ở trang ' + page + ')');
      props.deleteProperty('EXPLORE_PAGE');
      return;
    }
    list.forEach(function (p) {
      var name = String(p.product || '').toLowerCase();
      if (keywords.some(function (k) { return name.indexOf(k) >= 0; })) {
        Logger.log('🎯 product_id=' + p.product_id + '  |  ' + p.product + '  |  price=' + p.price + '  |  status=' + p.status);
      }
    });
    page++;
    if (page % 5 === 0) Logger.log('...đã quét tới trang ' + page);
  }
  props.setProperty('EXPLORE_PAGE', String(page));
  Logger.log('⏸️ Tạm dừng ở trang ' + page + ' (sắp hết giờ) — CHỌN LẠI exploreTscProductsByName → Chạy để tiếp tục.');
}

/** Chạy hàm này để test riêng API trungsoncare — log ra trạng thái + vài đơn gần nhất */
function testTscApi() {
  var st = tscFetch('statuses?type=O&items_per_page=100');
  Logger.log('Statuses: ' + (st ? JSON.stringify(st).slice(0, 400) : '❌ null — API lỗi hoặc sai key'));
  var now = new Date();
  var from = new Date(now.getTime() - 7 * 86400000);
  var od = tscFetch('orders?period=C&time_from=' + Math.floor(from.getTime() / 1000) + '&time_to=' + Math.floor(now.getTime() / 1000) + '&items_per_page=3');
  Logger.log('Orders 7 ngày (3 đơn mẫu): ' + (od ? JSON.stringify(od).slice(0, 600) : '❌ null'));
}

/** Chạy hàm này khi testTscApi() báo null — gọi thẳng không qua tscFetch để lộ ra ĐÚNG mã lỗi HTTP + nội dung CS-Cart trả về.
 * 401/403 = sai email hoặc key (hoặc IP bị chặn). 404 = sai đường dẫn API. 500 = lỗi phía CS-Cart. */
function testTscApiRaw() {
  Logger.log('TSC_API_EMAIL đang dùng: ' + TSC_API_EMAIL);
  Logger.log('TSC_API_KEY đang dùng (10 ký tự đầu): ' + String(TSC_API_KEY).slice(0, 10) + '...');
  var res = UrlFetchApp.fetch('https://trungsoncare.com/api/statuses?type=O&items_per_page=5', {
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(TSC_API_EMAIL + ':' + TSC_API_KEY) },
    muteHttpExceptions: true
  });
  Logger.log('HTTP status code: ' + res.getResponseCode());
  Logger.log('Response body: ' + res.getContentText().slice(0, 1000));
}

/**
 * CHẠY 1 LẦN — chỉ để dò thông tin cần thiết trước khi bật tính năng "đẩy đơn tự động vào hệ thống".
 * Lấy danh sách: payment methods (COD/CK là ID mấy), shipping methods (dùng ID nào),
 * và tìm product_id của 3 sản phẩm LP đang bán. Xem kết quả trong Execution log (Ctrl+Enter),
 * copy paste log gửi lại cho Claude để điền đúng ID vào tscCreateOrder().
 */
function exploreTscRefData() {
  Logger.log('===== PAYMENT METHODS =====');
  var pay = tscFetch('payments?items_per_page=50');
  Logger.log(pay ? JSON.stringify(pay) : '❌ null — kiểm tra lại API key');

  Logger.log('===== SHIPPING METHODS =====');
  var ship = tscFetch('shippings?items_per_page=50');
  Logger.log(ship ? JSON.stringify(ship) : '❌ null');

  Logger.log('===== PRODUCTS (tìm hong-sam / gas-whal / sangsangton) =====');
  ['hong', 'gas', 'sang'].forEach(function (kw) {
    var p = tscFetch('products?q=' + encodeURIComponent(kw) + '&items_per_page=10');
    Logger.log('--- q=' + kw + ' ---');
    Logger.log(p ? JSON.stringify(p).slice(0, 1500) : '❌ null');
  });

  Logger.log('===== 1 ĐƠN MẪU GẦN NHẤT (xem đủ field: company_id, s_country, s_state...) =====');
  var now = new Date(), from = new Date(now.getTime() - 30 * 86400000);
  var orders = tscFetch('orders?period=C&time_from=' + Math.floor(from.getTime() / 1000) + '&time_to=' + Math.floor(now.getTime() / 1000) + '&items_per_page=1');
  var list = (orders && orders.orders) || [];
  if (list[0]) {
    var full = tscFetch('orders/' + list[0].order_id);
    Logger.log(full ? JSON.stringify(full) : '❌ null');
  } else {
    Logger.log('Không có đơn nào trong 30 ngày để soi mẫu.');
  }
}

/**
 * Tổng hợp báo cáo — luật đếm:
 * - Lead COD: tính ngay khi điền form
 * - Lead CK: CHỈ tính khi hệ thống admin đã xác nhận (Xác nhận đơn / Đã thanh toán / Hoàn tất)
 * - Đơn xác nhận + doanh thu: lấy từ hệ thống, khớp 9 số cuối SĐT
 */
function buildReport(start, end, title) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues().forEach(function (r) {
      var t = r[0];
      if (!(t instanceof Date) || t < start || t > end) return;
      rows.push({ phone: last9(r[3]), amount: Number(r[9]) || 0, isCod: String(r[8]).indexOf('COD') >= 0 });
    });
  }

  // Đối soát hệ thống (đơn có thể lên trễ hơn lead → nới 1 ngày)
  var apiOk = false, confirmedPhones = {}, orders = 0, revenue = 0;
  var allPhones = {};
  rows.forEach(function (r) { if (r.phone) allPhones[r.phone] = true; });
  var statusMap = tscStatusMap();
  if (Object.keys(statusMap).length) {
    apiOk = true;
    var endExt = new Date(Math.min(end.getTime() + 86400000, Date.now()));
    tscOrders(start, endExt).forEach(function (o) {
      var name = statusMap[o.status] || '';
      if (TSC_CONFIRM_STATUSES.indexOf(name) < 0) return;
      var p = last9(o.phone || o.b_phone || o.s_phone);
      if (p && allPhones[p]) { confirmedPhones[p] = true; orders++; revenue += Number(o.total) || 0; }
    });
  }

  var leads = 0, potential = 0, cod = 0, bankOk = 0, bankPending = 0;
  rows.forEach(function (r) {
    if (r.isCod) { leads++; potential += r.amount; cod++; }
    else if (r.phone && confirmedPhones[r.phone]) { leads++; potential += r.amount; bankOk++; }
    else { bankPending++; } // CK chưa xác nhận — không tính lead
  });

  return title +
    '\n🏥 Dự án: ' + PROJECT_NAME +
    '\n————————————' +
    '\n📥 Tổng lead: ' + leads +
    '\n✅ Đơn xác nhận' + (apiOk ? ' (hệ thống)' : ' (⚠️ API lỗi)') + ': ' + orders +
    '\n💰 Doanh thu xác nhận: ' + revenue.toLocaleString('vi-VN') + 'đ' +
    '\n💵 Tổng giá trị lead: ' + potential.toLocaleString('vi-VN') + 'đ' +
    '\n📦 COD: ' + cod + '  |  🏦 CK xác nhận: ' + bankOk +
    (bankPending ? '  |  ⏳ CK đang chờ: ' + bankPending + ' (không tính)' : '');
}

function formatProducts(s) {
  var names = { 'hong-sam': 'Hồng Sâm Dongwha', 'gas-whal': 'Gas Whal', 'sangsangton': 'SangSangTon Up' };
  return String(s).split(',').filter(Boolean).map(function (p) {
    var a = p.split(':');
    return (names[a[0]] || a[0]) + ' x' + (a[1] || 1);
  }).join(', ');
}
