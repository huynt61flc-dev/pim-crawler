# Bot Tự Động Thu Thập Thông Tin Từ PIM.FRT.VN Vào Excel

Bot tự động đọc mã sản phẩm từ **Cột B** của file Excel, tìm kiếm trên hệ thống PIM nội bộ (**pim.frt.vn/product/upc-list**), trích xuất tên sản phẩm từ bảng thông tin (Right Drawer) và điền vào **Cột C**.

---

## 🌟 Tính Năng Nổi Bật

- **Tự động hóa 100%**: Tìm kiếm mã ➔ Click Tên UPC ➔ Mở Drawer bên phải ➔ Trích xuất tên sản phẩm sạch giữa dấu `-` và chữ `Đã duyệt` ➔ Ghi vào Cột C.
- **Lưu phiên đăng nhập**: Dữ liệu đăng nhập (SSO / OTP) được lưu vào thư mục `browser_profile` cục bộ, chỉ cần đăng nhập 1 lần duy nhất.
- **Bảo toàn dữ liệu**: Tự động tạo bản sao lưu (`_backup.xlsx`) trước khi chạy, tự động lưu sau mỗi dòng, bỏ qua các dòng đã có tên nếu chạy lại.
- **Chạy nhanh bằng 1 click**: Hỗ trợ file `run_bot.bat` hoặc lệnh `/pim`.

---

## 🛠️ Cài Đặt

1. **Cài đặt thư viện Python**:
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```

2. **Cấu hình file `config.json`**:
   ```json
   {
     "data_source": "excel",
     "excel_path": "sample_input.xlsx",
     "sheet_name": "",
     "start_row": 2,
     "code_column": "B",
     "result_column": "C",
     "pim_url": "https://pim.frt.vn/product/upc-list",
     "headless": false,
     "request_delay_seconds": 1.5
   }
   ```

---

## 🚀 Hướng Dẫn Sử Dụng

1. Đặt file Excel của bạn (có chứa mã sản phẩm ở Cột B) vào thư mục dự án (hoặc dùng `sample_input.xlsx`).
2. Đóng file Excel trong Microsoft Excel trước khi chạy.
3. Chạy bot:
   - **Cách 1**: Click đúp vào file `run_bot.bat`.
   - **Cách 2**: Chạy lệnh `python main.py` trong terminal.
4. Đăng nhập tài khoản FPT + OTP trên trình duyệt lần đầu tiên.
5. Nhấn **ENTER** trong terminal để bot tự động cào và cập nhật dữ liệu.
