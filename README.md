# VNPost Huế - VIP Customer Dashboard (V2.0)

Hệ thống theo dõi và báo cáo CSKH VIP chuyên sâu, sử dụng Python Backend và SQLite.

## Hướng dẫn cài đặt & Khởi chạy

### 1. Khởi chạy nhanh
- Nhấp đúp chuột vào file `run_server.bat` ở thư mục gốc.
- Một cửa sổ CMD sẽ hiện ra và thông báo địa chỉ IP của máy bạn (Ví dụ: `http://192.168.1.10:8088`).
- Các máy khác trong mạng nội bộ chỉ cần nhập địa chỉ này vào trình duyệt để xem Dashboard.

### 2. Cài đặt tự động khởi động cùng Windows (Autorun)
Để Dashboard luôn sẵn sàng ngay khi bạn mở máy:
1. Nhấn phím `Windows + R`, gõ `shell:startup` và nhấn Enter.
2. Cửa sổ thư mục `Startup` của Windows sẽ hiện ra.
3. Nhấp chuột phải vào file `run_server.bat` của dự án, chọn **Create shortcut**.
4. Kéo shortcut vừa tạo vào thư mục `Startup` đã mở ở bước 2.
5. Xong! Từ nay mỗi khi bạn mở máy, server sẽ tự động chạy ngầm.

## Tính năng chính
- **Nhận diện VIP**: Tự động lấy tên khách hàng từ dòng 1 của file Excel.
- **Đếm chính xác**: Xử lý logic "Thành công" chuẩn xác từ dữ liệu thực tế (Đã sửa lỗi đếm 0).
- **SQLite Persistence**: Dữ liệu được lưu trữ an toàn trong file `backend/cskh_vip.db`.
- **LAN Ready**: Cho phép toàn bộ nhân viên trong phòng truy cập báo cáo qua mạng nội bộ.

## Hướng dẫn sử dụng
1. Nhấp đúp vào [run_server.bat](file:///d:/Antigravity%20-%20Project%20-%20TTVH/CSKH/run_server.bat).
2. Mở trình duyệt và truy cập địa chỉ được hiển thị (thường là `http://localhost:8088`).
3. Nhấn **NHẬP EXCEL** và chọn file của bạn để xem kết quả tức thì.

## Yêu cầu hệ thống
- Python 3.10+
- Thư viện: `pip install fastapi uvicorn pandas openpyxl` (Đã được cài đặt trong quá trình setup).
