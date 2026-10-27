# Biểu mẫu nhân sự TBS Logistics

Trang tĩnh (GitHub Pages) cho các biểu mẫu của phòng HCNS, cùng giao diện SafetyHub / TBS Logistics 2026.

| Đường dẫn | Biểu mẫu | Máy chủ |
|---|---|---|
| `/` | Trang chọn biểu mẫu | – |
| `/ho-so-du-tuyen/` | **Hồ sơ dự tuyển** (TBSL-NS-QT01-BM02), dựng theo mẫu Word | `apps-script/HoSoDuTuyen.gs` (cài theo `apps-script/HUONG-DAN.md`) |
| `/exit/` | **Phiếu phỏng vấn nghỉ việc** (TBSL-NS-QT03-BM05), chuyển từ repo `core-value-survey`; bắt buộc ký tên | Apps Script phiếu nghỉ việc (Code.gs + pdf_view.html gửi riêng, không để trong repo công khai vì có mật khẩu quản trị) |
| `/quan-ly/` | **Trang quản lý HCNS**: gộp hồ sơ dự tuyển + phiếu nghỉ việc, lọc, xuất Excel, Xem / in từng người | Gọi API quản trị của hai Apps Script trên (đăng nhập email + mật khẩu) |

Link sau khi bật GitHub Pages (Settings › Pages › Deploy from a branch › `main` / root):

- https://lehuynhdatthinh-max.github.io/tbs-hr-forms/ho-so-du-tuyen/
- https://lehuynhdatthinh-max.github.io/tbs-hr-forms/exit/
- https://lehuynhdatthinh-max.github.io/tbs-hr-forms/quan-ly/ (HCNS)

## Hồ sơ dự tuyển

- Đủ 6 phần của mẫu giấy: I. Thông tin cá nhân (kèm ảnh 3x4), II. Gia đình, III. Học vấn (văn bằng, khóa học, ngoại ngữ, vi tính),
  IV. Kinh nghiệm làm việc, V. Thông tin khác, VI. Tham chiếu, cam kết và chữ ký. Phần VII để bộ phận tuyển dụng ký trên bản in.
- Dùng tốt trên điện thoại: chụp ảnh 3x4 (tự cắt đúng tỉ lệ), ký bằng ngón tay, tự lưu bài đang điền trên máy, tự thử lại khi mất mạng.
- Mỗi hồ sơ là một dòng trong Google Sheet; cột "Xem / in hồ sơ" mở bản in theo bố cục mẫu Word.
