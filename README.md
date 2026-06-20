# 🛡️ Hệ Thống Kiểm Định Chứng Từ Xuất Nhập Khẩu E54/E23

Ứng dụng nội bộ dành cho phòng Logistics, tự động hóa toàn diện quy trình đối chiếu chéo (Cross-check) giữa chứng từ Hải quan (Invoice, Packing List, Chỉ định giao hàng) và Hệ thống quản lý kho (SAP/ERP).

## 🌟 Tính Năng Cốt Lõi
- **Đa định dạng đầu vào:** Đọc mượt mà dữ liệu từ Excel, CSV, PDF (bảng biểu) và thậm chí là **Ảnh chụp/Scan (AI OCR)**.
- **Tối ưu hóa phần cứng:** Tự động nén ảnh đầu vào và sử dụng AI bản CPU siêu nhẹ để ngăn chặn lỗi tràn RAM trên máy chủ Cloud.
- **Bảo mật tuyệt đối:** Tích hợp cổng kiểm tra mật khẩu nội bộ trước khi cấp quyền sử dụng. (Streamlit Secrets).
- **Double Check & Xử lý rác:** Tự động loại bỏ các dòng Header thừa của biểu mẫu, có tính năng *Preview* để nhân viên kiểm tra lại file trước khi merge.
- **Đánh giá logic thông minh (HITL):** Highlight các dòng số liệu bị vênh màu sắc trực quan (Xanh: Khớp, Đỏ: Lệch file gốc, Cam: Lệch ERP).
- **Xuất báo cáo tự động:** Trả về file Excel dữ liệu sạch để nạp trực tiếp vào ECUS5/VNACCS và biên bản PDF.

---

## 💻 Hướng Dẫn Chạy Tool Nội Bộ (Cho Nhân Viên)

Ứng dụng này đã được cấu hình chạy trên nền tảng Web. Bạn không cần cài đặt bất cứ phần mềm nào.

1. Xin cấp mật khẩu truy cập từ Quản trị viên (Phòng XNK).
2. Truy cập vào đường link hệ thống.
3. Nhập mật khẩu.
4. Kéo thả 4 tệp dữ liệu vào các ô tương ứng.
5. Nhập số dòng tiêu đề thừa (Header) cần bỏ qua.
6. Bấm **BẮT ĐẦU ĐỐI CHIẾU** và xem kết quả.

---

## 🛠️ Hướng Dẫn Cài Đặt (Dành Cho IT / Admin)

Nếu bạn muốn chạy thử hệ thống này trên máy tính cá nhân (Local) để test code trước khi đưa lên Cloud, vui lòng làm theo các bước sau:

### Bước 1: Khởi tạo môi trường
Yêu cầu hệ