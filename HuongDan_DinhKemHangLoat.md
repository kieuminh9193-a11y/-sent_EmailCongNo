# HƯỚNG DẪN GỬI EMAIL VÀ ĐÍNH KÈM FILE HÀNG LOẠT (DEBTMAIL PRO)

Tài liệu này hướng dẫn cách sử dụng công cụ **DebtMail Pro** để gửi email nhắc nợ hàng loạt, kèm theo file đính kèm (hóa đơn, biên bản công nợ, đối chiếu,...) riêng biệt cho từng khách hàng một cách tự động.

---

## 📋 BƯỚC 1: CHUẨN BỊ MẪU NHẬP LIỆU EXCEL
Khi bạn tải Excel mẫu từ ứng dụng hoặc chuyển đổi từ file ERP, sheet **DanhSach** sẽ có các cột quan trọng sau:

1. **customer_name**: Tên khách hàng (Ví dụ: `CÔNG TY CP ABC`).
2. **customer_code**: Mã khách hàng (Ví dụ: `KH001`).
3. **attachment_dir** *(Không bắt buộc)*: Đường dẫn thư mục chứa file riêng của khách hàng này (Ví dụ: `D:\CongNo\PDF`).
4. **attachment_file** *(Không bắt buộc)*: Tên file đính kèm riêng của khách hàng này (Ví dụ: `KH001_HoaDon.pdf`).

---

## ⚡ BƯỚC 2: CÁC PHƯƠNG ÁN ĐÍNH KÈM TỆP TỰ ĐỘNG

Ứng dụng hỗ trợ **2 phương án** để đính kèm file riêng của từng khách hàng. Bạn có thể chọn cách phù hợp nhất:

### PHƯƠNG ÁN 1: Tự động khớp file theo Mã KH (Khuyên dùng)
Cách này phù hợp khi bạn xuất tất cả file PDF hóa đơn/công nợ của mọi khách hàng ra **chung 1 thư mục**.

*   **Cách đặt tên file:** Đảm bảo tên file chứa **Mã khách hàng** tương ứng (không phân biệt chữ hoa/thường).
    *   *Ví dụ:* Mã KH là `KH001` -> Tên file có thể là `KH001_CongNo.pdf`, `Invoice_KH001.xlsx`, `KH001.pdf`.
*   **Thao tác trên phần mềm:**
    1. Nhập đường dẫn thư mục chứa tất cả các file đó vào ô **"📁 Thư mục file riêng KH (Auto-match theo mã)"** ở cột bên phải giao diện ứng dụng.
    2. Giao diện sẽ tự động quét thư mục và hiển thị các file khớp được thành thẻ màu xanh lá dưới nhãn **"FILE TỰ ĐỘNG KHỚP (MÃ KH)"** khi bạn chọn vào khách hàng đó.

---

### PHƯƠNG ÁN 2: Nhập đường dẫn và tên file trực tiếp trong Excel
Cách này phù hợp khi các file của bạn lưu trữ ở nhiều thư mục khác nhau hoặc tên file không có cấu trúc cố định.

*   **Cách thiết lập trong Excel:**
    *   Tại cột `attachment_dir`: Điền đường dẫn thư mục chứa file (Ví dụ: `D:\HoaDon\Thang7`).
    *   Tại cột `attachment_file`: Điền chính xác tên file kèm đuôi file (Ví dụ: `Invoice_ABC.pdf`).
    *   *Mẹo:* Bạn cũng có thể điền luôn đường dẫn đầy đủ vào cột `attachment_file` (Ví dụ: `D:\HoaDon\Thang7\Invoice_ABC.pdf`) và bỏ trống cột `attachment_dir`.
*   **Thao tác trên phần mềm:** Nạp file Excel đã điền đầy đủ thông tin này vào phần mềm. Các tệp đính kèm sẽ tự động hiển thị trên giao diện và được gộp vào danh sách đính kèm khi gửi.

---

## 📤 BƯỚC 3: TIẾN HÀNH GỬI EMAIL HÀNG LOẠT

1. **Kết nối Outlook:** Đảm bảo phần mềm **Outlook Desktop** đã mở trên máy tính và hiển thị trạng thái kết nối **Online** (chấm xanh lá cây).
2. **Chọn tài khoản gửi:** Chọn tài khoản email chính hoặc alias tương ứng trong mục **From:**.
3. **Chọn khách hàng gửi:** 
    *   Tích chọn vào các ô vuông ở đầu danh sách khách hàng để chọn những người cần gửi.
    *   Nhấn **☑ Chọn tất cả** nếu muốn gửi toàn bộ danh sách.
4. **Chọn Chế độ gửi (Send Mode):**
    *   🔍 **Hiển thị trên Outlook trước (Display):** Tool sẽ tạo sẵn email nháp và bật cửa sổ Outlook lên để bạn kiểm tra nội dung và file đính kèm trước khi bấm gửi thủ công.
    *   ⚡ **Gửi trực tiếp (Send):** Tool sẽ tự động gửi trực tiếp qua nền Outlook mà không hiện cửa sổ. Thích hợp khi đã kiểm tra thử 1-2 email đầu tiên thấy chính xác.
5. **Bắt đầu gửi:** Nhấn nút **📨 Gửi tất cả đã chọn (N)**. Hệ thống sẽ lần lượt xử lý từng email và đính kèm chính xác tệp tin của từng khách hàng.
