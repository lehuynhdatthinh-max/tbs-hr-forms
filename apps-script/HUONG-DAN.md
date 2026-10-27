# Cài máy chủ cho "Hồ sơ dự tuyển" (khoảng 5 phút)

Trang `ho-so-du-tuyen/` cần một Apps Script riêng (không dùng chung với khảo sát VHDN hay phiếu nghỉ việc).

1. Tạo **Google Sheet mới**, đặt tên ví dụ `Hồ sơ dự tuyển - Phản hồi`.
2. **Tiện ích mở rộng › Apps Script**. Xóa nội dung `Code.gs`, dán toàn bộ `apps-script/HoSoDuTuyen.gs`. Bấm **Lưu**.
3. Chọn hàm **`setup`** › **Chạy** › cấp quyền (Sheets, Drive, gửi email).
   Script tạo tab **Hồ sơ dự tuyển** và thư mục Drive **Hồ sơ dự tuyển - Ảnh & chữ ký** (chỉ chủ sheet xem được).
4. **Triển khai › Triển khai mới** › loại **Ứng dụng web**:
   - Thực thi dưới dạng: **Tôi**
   - Người có quyền truy cập: **Bất kỳ ai**
   Bấm **Triển khai** và sao chép URL kết thúc bằng `/exec`.
5. Gửi URL `/exec` cho Claude (hoặc tự dán vào `const WEB_APP_URL = "";` trong `ho-so-du-tuyen/index.html`).
6. (Tùy chọn) Nhận email khi có hồ sơ mới: trong hàm `setNotifyEmails` điền email, ví dụ `'tuyendung@tbslogistics.com.vn'`, rồi **Chạy** một lần.

Sửa script sau này: **Triển khai › Quản lý triển khai** › bút chì › **Phiên bản mới** › **Triển khai** (giữ nguyên URL).

## Dữ liệu nằm ở đâu

- Mỗi hồ sơ là **một dòng** trong tab **Hồ sơ dự tuyển**: các cột chính (họ tên, ngày sinh, vị trí, di động, CMND/CCCD…) để lọc nhanh, toàn bộ nội dung ở cột **Dữ liệu (JSON)**.
- Cột **Xem / in hồ sơ**: mở bản in đúng bố cục mẫu Word TBSL-NS-QT01-BM02 (có ảnh 3x4, chữ ký, phần VII để trống cho bộ phận tuyển dụng ký trên giấy). Bấm **In / Lưu PDF**.
- Cột **Trạng thái** (mặc định "Mới") và **Ghi chú HCNS** để phòng HCNS tự ghi.
- Ảnh 3x4 và chữ ký nằm trong thư mục Drive nói trên.
- Ứng viên nộp lại khi mất mạng (trang tự thử lại) không tạo dòng trùng.

## Lưu ý

- Ai có đường link "Xem / in" của một hồ sơ thì xem được hồ sơ đó (link chứa mã ngẫu nhiên, không đoán được). Chỉ gửi link cho người cần xem.
- Bản in lấy logo từ `https://lehuynhdatthinh-max.github.io/tbs-hr-forms/images/logo.png`, nên cần bật GitHub Pages cho repo.
