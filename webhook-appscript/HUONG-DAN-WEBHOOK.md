# CÀI WEBHOOK NHẬN ĐƠN VỀ GOOGLE SHEET — 5 PHÚT

## Bước 1 — Tạo Sheet
1. Vào [sheets.new](https://sheets.new) → đặt tên file: **Đơn hàng LP Dongwha**

## Bước 2 — Gắn Apps Script
1. Trong Sheet: menu **Tiện ích mở rộng (Extensions) → Apps Script**
2. Xóa code mẫu, dán toàn bộ nội dung file `Code.gs` (cùng thư mục này) vào
3. Bấm **💾 Lưu**

## Bước 3 — Cấp quyền
1. Chọn hàm `testSetup` trên thanh công cụ → bấm **Run (▶)**
2. Google hỏi quyền → **Review permissions** → chọn tài khoản → **Advanced → Go to ... (unsafe) → Allow**
   (cảnh báo "unsafe" là bình thường vì script tự viết, chưa qua Google verify)
3. Chạy xong sẽ nhận 1 email test "✅ Test webhook LP Dongwha" → tức là OK

## Bước 4 — Deploy Web App
1. Bấm **Deploy (Triển khai) → New deployment**
2. Bấm bánh răng ⚙ → chọn **Web app**
3. Cấu hình — QUAN TRỌNG, sai là không nhận được đơn:
   - Execute as: **Me** (tài khoản của bạn)
   - Who has access: **Anyone** (Bất kỳ ai)
4. Bấm **Deploy** → copy **Web app URL** (dạng `https://script.google.com/macros/s/AKfycb.../exec`)

## Bước 5 — Gửi URL cho Claude
Dán URL `/exec` vào chat → Claude thay vào `WEBHOOK_URL` trong landing page → deploy lại là chạy.

---

## Sheet nhận được gì mỗi đơn
Thời gian · Mã đơn · Họ tên · SĐT · Địa chỉ · Sản phẩm · Số lượng · Bump · COD/CK · Tổng tiền · Ghi chú · **Trạng thái gọi** (cột cho dược sĩ đánh dấu) · UTM đầy đủ · fbclid · Event ID (đối soát CAPI)

Kèm **email báo đơn mới tức thì** về `lctbinh0006@gmail.com` (đổi/tắt trong dòng `NOTIFY_EMAIL` của Code.gs).

---

# THÊM BÁO ĐƠN QUA ZALO BOT — 5 BƯỚC

## Bước 1 — Tạo bot (làm trên điện thoại)
1. Mở app **Zalo** → tìm kiếm OA **"Zalo Bot Manager"**
2. Vào chat với OA đó → bấm **Tạo bot** trong menu
3. Đặt tên bot — bắt buộc bắt đầu bằng chữ `Bot`, ví dụ: **Bot TrungSon Don Hang**
4. Tạo xong, Zalo **nhắn tin cho bạn cái Bot Token** → copy lại

## Bước 2 — Điền token vào Script Properties (KHÔNG dán thẳng vào code nữa)
Mở lại Apps Script (Sheet → Tiện ích mở rộng → Apps Script) → **dán đè toàn bộ Code.gs bản mới** (file này đã cập nhật) → bấm biểu tượng bánh răng **Project Settings** (menu bên trái) → cuộn xuống **Script Properties** → **Add script property**:
```
Property: SP_ZALO_BOT_TOKEN
Value:    <dán token vừa copy vào đây>
```
→ Bấm **Save script properties**.

⚠️ Code không còn chứa token cứng nữa — mọi secret (Zalo Bot Token, Meta CAPI Token, CS-Cart API key) đều đọc từ Script Properties, không được dán trực tiếp vào Code.gs (repo GitHub đang public).

## Bước 3 — Nhắn cho bot 1 tin
Trên Zalo, tìm bot vừa tạo → nhắn bất kỳ, ví dụ "hi" (để bot biết mày là ai)

## Bước 4 — Lấy chat_id
Trong Apps Script: chọn hàm **getZaloChatId** → **Run** → mở **Execution log** → thấy dòng `chat_id: xxxx` → copy dán vào:
```
var ZALO_CHAT_IDS = ['xxxx'];
```
(Muốn báo cho cả team: mỗi người nhắn "hi" cho bot → chạy lại getZaloChatId → thêm hết ID vào: `['id1','id2','id3']`)

## Bước 5 — Deploy phiên bản mới ⚠️ QUAN TRỌNG
**Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**
(Không làm bước này thì code mới KHÔNG chạy — URL /exec giữ nguyên, không cần sửa gì bên landing page)

## Test
Chạy hàm **testSetup** → bot Zalo phải nhắn "✅ Test: Zalo Bot đã kết nối..." vào máy mày. Xong → đặt 1 đơn thử trên web → Sheet + Email + Zalo phải nổ cùng lúc.

---

## Lưu ý
- Mỗi lần sửa Code.gs phải **Deploy → Manage deployments → Edit → New version** thì thay đổi mới ăn.
- URL `/exec` và Bot Token là bí mật vận hành — đừng đăng công khai.
