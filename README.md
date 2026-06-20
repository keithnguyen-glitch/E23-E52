# 🏢 HỆ THỐNG KIỂM ĐỊNH CHỨNG TỪ XNK TỔNG HỢP (E54 / E23)

**Phiên bản:** 2.0 (Enterprise Dashboard)
**Nghiệp vụ:** Logistics / Xuất Nhập Khẩu (Khai báo Hải quan, Quyết toán Gia công)
**Đơn vị áp dụng:** Ching Luh & Fu-Luh

---

## 🌟 TỔNG QUAN HỆ THỐNG
Đây là hệ thống phần mềm tự động hóa (Framework) được thiết kế chuyên biệt để đối chiếu chéo (Cross-check) dữ liệu Xuất Nhập Khẩu. Thay vì dùng hàm VLOOKUP thủ công trên Excel, hệ thống sử dụng Python và AI để đọc, làm sạch, quy đổi và đối chiếu hàng ngàn dòng vật tư chỉ trong vài giây.

### ✨ CÁC TÍNH NĂNG ĐỘT PHÁ:
1. **Đối chiếu liên thông 5 chiều:** Tự động so khớp dữ liệu giữa `Invoice` ↔ `Packing List` ↔ `Chỉ Định Giao Hàng` ↔ `SAP/ERP` ↔ `Tờ khai ECUS`.
2. **Quy đổi Master Data (HSQĐ):** Hệ thống ngầm tự động dò tìm Hệ số quy đổi (Ví dụ: YD sang MTK) và đồng bộ Đơn vị tính (UOM) về chuẩn VNACCS của Hải quan.
3. **Đọc đa định dạng bằng AI (OCR):** Hỗ trợ đọc dữ liệu bảng biểu từ `Excel`, `CSV`, `PDF` và bóc tách chữ từ `Hình ảnh chụp/Scan` (JPG, PNG).
4. **Bộ Lọc Lỗi Thông Minh (Tolorence):** - Tự động bỏ qua các sai số làm tròn (Dung sai) đối với hàng cân ký (KGM) hoặc mét vuông (MTK).
   - Kiểm tra logic toán học: `Số lượng` x `Đơn giá` = `Thành tiền`.
   - Nút gạt (Toggle) giúp ẩn nhanh các mã đã khớp, chỉ hiển thị dòng bị lỗi.
5. **Trích xuất 1-Click:** Xuất File Excel dữ liệu đã chuẩn hóa (để nạp thẳng vào phần mềm ECUS) và Biên bản báo cáo file PDF.

---

## 📁 CẤU TRÚC THƯ MỤC DỰ ÁN

Để hệ thống hoạt động, thư mục của bạn cần có cấu trúc như sau:

```text
📁 E54_CHECKER_FRAMEWORK/
 ├── 📄 app.py               # Mã nguồn chính của giao diện Dashboard
 ├── 📄 requirements.txt     # Danh sách thư viện Python
 ├── 📄 packages.txt         # Cấu hình gói thư viện Linux (Chỉ dùng khi up lên Cloud)
 ├── 📄 START_APP.bat        # File khởi động 1-click dành cho Staff (Chạy Local)
 ├── 📄 README.md            # Tài liệu hướng dẫn sử dụng này
 └── 📁 .streamlit/
      └── 📄 secrets.toml    # File chứa mật khẩu bảo mật (Không up lên mạng)
