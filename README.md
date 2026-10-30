# 📧 DebtMail Pro - Công Cụ Gửi Email Nhắc Nợ Hàng Loạt Qua Outlook

**DebtMail Pro** là công cụ tự động hóa gửi email thông báo / nhắc nợ quá hạn chuyên nghiệp cho doanh nghiệp, kế toán công nợ và bộ phận tài chính. Ứng dụng tích hợp trực tiếp với Microsoft Outlook Desktop thông qua COM Automation, hỗ trợ cá nhân hóa nội dung email (Mail Merge), tự động tạo bảng chi tiết công nợ từng hóa đơn và tự động ghép file đính kèm riêng cho từng khách hàng.

---

## ✨ Tính Năng Nổi Bật

- 🚀 **Tích hợp Outlook Desktop (COM Automation):**
  - Không cần cấu hình SMTP phức tạp.
  - Tự động nhận diện tài khoản Outlook và các email Alias (Send on Behalf / Proxy Addresses).
  - Hai chế độ gửi: **Hiển thị nháp trên Outlook (Display)** để kiểm tra trước hoặc **Gửi tự động ngầm (Send)**.
- 📊 **Quản lý dữ liệu công nợ:**
  - Nhập danh sách công nợ trực tiếp từ Excel (`.xlsx`).
  - Hỗ trợ chuyển đổi tự động từ định dạng xuất công nợ hệ thống ERP sang template chuẩn.
  - Bảng tổng hợp công nợ tự động tính tổng tiền phải thu, đã thu và còn lại phải thu.
- 📎 **Đính kèm tài liệu hàng loạt thông minh:**
  - **Auto-match theo Mã KH:** Tự động quét thư mục hóa đơn và ghép đúng các file PDF/Excel có chứa mã khách hàng vào từng email.
  - **Đính kèm tùy chỉnh qua Excel:** Chỉ định thư mục hoặc tên file riêng cho từng khách hàng ngay trong sheet dữ liệu.
- 📝 **Soạn thảo & Mail Merge linh hoạt:**
  - Hỗ trợ các biến động: `{customer_name}`, `{company_name}`, `{debt_date}`, `{payment_deadline}`, `{debt_table}`, `{signature}`, `{short_name}`.
  - Tự động định dạng tiêu đề email theo mẫu quy định.
- 📋 **Lịch sử & Thống kê:**
  - Theo dõi số lần gửi nhắc nợ cho từng khách hàng.
  - Xem trước email thực tế (Live Preview) trước khi gửi.

---

## 🛠️ Cấu Trúc Thư Mục

```text
├── app.py                      # Flask backend & API điều khiển
├── outlook_sender.py           # Module kết nối và gửi mail qua Outlook COM
├── email_template.py           # Module Mail Merge & tạo bảng HTML công nợ
├── convert_to_template.py      # Tiện ích chuyển đổi dữ liệu ERP / Excel gốc
├── license_check.py            # Kiểm tra thời hạn & kích hoạt ứng dụng
├── DebtMailPro.spec            # Cấu hình đóng gói PyInstaller (.exe)
├── requirements.txt            # Danh sách thư viện Python phụ thuộc
├── settings.example.json       # File cấu hình mẫu
├── Template_Mau.xlsx           # File Excel dữ liệu mẫu (sạch, không có data thực)
├── HuongDan_DinhKemHangLoat.md # Hướng dẫn chi tiết tính năng đính kèm tệp
├── static/                     # Giao diện frontend (HTML5, CSS, JS)
│   ├── app.js
│   └── style.css
└── templates/                  # Template giao diện người dùng
    └── index.html
```

---

## 🚀 Hướng Dẫn Cài Đặt & Sử Dụng

### 1. Yêu cầu hệ thống
- Hệ điều hành: **Windows 10 / 11**
- Microsoft Outlook Desktop (2016, 2019, 2021 hoặc Microsoft 365) đã đăng nhập tài khoản email.
- Python 3.9+ (nếu chạy từ mã nguồn).

### 2. Cài đặt các thư viện cần thiết
```bash
pip install -r requirements.txt
```

### 3. Chạy ứng dụng
```bash
python app.py
```
Sau khi khởi động, trình duyệt sẽ tự động mở giao diện tại địa chỉ: `http://localhost:5000`

---

## 📋 Hướng Dẫn Sử Dụng Nhanh

1. **Chuẩn bị dữ liệu Excel:**
   - Tải file mẫu từ giao diện (`Tải Excel mẫu`) hoặc dùng file [Template_Mau.xlsx](file:///d:/OneDrive%20-%20SHINE%20GROUP/00.PHAT/900.Canhan/DuAn/03.Sent_Email_CongNo/Template_Mau.xlsx).
   - Điền thông tin vào 2 sheet:
     - `DanhSach`: Thông tin khách hàng, email người nhận (`email_to`), CC (`email_cc`), BCC, hạn thanh toán,...
     - `CongNo`: Chi tiết từng hóa đơn quá hạn (tên khách hàng phải khớp chính xác với sheet `DanhSach`).
2. **Nạp dữ liệu vào ứng dụng:**
   - Nhấn **Tải file Excel lên** và chọn file dữ liệu của bạn.
   - Hoặc nếu dùng file báo cáo xuất từ hệ thống ERP, sử dụng chức năng **Chuyển đổi ERP**.
3. **Cài đặt email & chữ ký:**
   - Vào mục **Cài đặt** để cấu hình tên công ty, tiêu đề mặc định và chữ ký email.
4. **Đính kèm file (Tùy chọn):**
   - Nhập đường dẫn thư mục chứa hóa đơn PDF vào ô **Thư mục file riêng KH**. Hệ thống sẽ tự động quét và ghép file theo `customer_code`.
5. **Gửi email:**
   - Chọn khách hàng cần gửi (hoặc nhấn **Chọn tất cả**).
   - Chọn chế độ **Hiển thị trên Outlook (Display)** hoặc **Gửi trực tiếp (Send)**.
   - Nhấn **Gửi tất cả đã chọn**.

---

## 🔒 Bảo Mật & Riêng Tư

- Mã nguồn đã được chuẩn hóa để loại bỏ toàn bộ thông tin cá nhân, tài khoản thực, chữ ký nội bộ và số liệu công nợ thực tế.
- Ứng dụng chạy trực tiếp trên máy cục bộ (`localhost`), không lưu trữ hay chuyển tiếp dữ liệu qua máy chủ bên ngoài.
- Dữ liệu gửi đi được quản lý hoàn toàn qua ứng dụng Outlook Desktop đã được cấp quyền của người dùng.

---

## 📄 Bản Quyền
Dự án được phân phối cho mục đích sử dụng nội bộ và phát triển giải pháp tự động hóa văn phòng.
