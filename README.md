# 🏢 HỆ THỐNG KIỂM ĐỊNH CHỨNG TỪ XNK TỔNG HỢP (E54 / E23)

**Phiên bản:** 3.0 (Enterprise Architecture & Logic Engine)
**Nghiệp vụ áp dụng:** Logistics / Xuất Nhập Khẩu (Khai báo Hải quan, Quyết toán Gia công)
**Đơn vị áp dụng:** Ching Luh & Fu-Luh



[Image of data integration architecture diagram]


---

## 🌟 TỔNG QUAN HỆ THỐNG (SYSTEM OVERVIEW)
Đây là hệ thống phần mềm tự động hóa (Framework) được thiết kế chuyên biệt để đối chiếu chéo (Cross-check) dữ liệu Xuất Nhập Khẩu. Hệ thống giải quyết triệt để bài toán sai lệch số liệu giữa Chứng từ nội bộ, Sổ sách Kế toán kho (SAP/ERP) và Dữ liệu khai báo Hải quan (ECUS/VNACCS).

Thay vì sử dụng hàm VLOOKUP thủ công và dễ sai sót trên Excel, hệ thống sử dụng Python (Pandas) kết hợp AI (OCR) để thiết lập một luồng xử lý tự động từ khâu bóc tách, chuẩn hóa, quy đổi đến đánh giá dung sai.

---

## 🧠 GIẢI THÍCH LUỒNG LOGIC THUẬT TOÁN (CORE LOGIC ENGINE)

Toàn bộ sức mạnh của dự án nằm ở **Quy trình xử lý 6 Pha (6-Phase Pipeline)**. Dưới đây là giải thích chi tiết cách máy tính tư duy và xử lý dữ liệu:

### Pha 1: Trích xuất & Làm sạch thông minh (Smart Extraction & Cleaning)
- **Đa định dạng:** Đọc mượt mà dữ liệu từ `Excel`, `CSV`, bảng biểu `PDF`. Tích hợp công nghệ **EasyOCR** để bóc tách text từ ảnh chụp/scan (`JPG/PNG`).
- **Dynamic Header (Bỏ qua dòng rác):** Các biểu mẫu XNK thường có phần thông tin công ty, địa chỉ ở đầu. Hệ thống cho phép cắt bỏ linh hoạt số dòng rác (`skiprows`) để đưa bảng dữ liệu về chuẩn cấu trúc hàng/cột.
- **Dọn dẹp:** Tự động loại bỏ các cột không tên (`Unnamed`), các dòng trống (`NaN`) trước khi đưa vào bộ nhớ.

### Pha 2: Ánh xạ và Đồng bộ (Mapping & Normalization)
- **Thuật toán `get_col_fallback`:** Không hardcode (gắn cứng) tên cột. Máy tính sẽ quét bằng từ khóa (Ví dụ: tìm cột `Material` hoặc `Mã NL`), giúp hệ thống vẫn chạy đúng dù đối tác có thay đổi tên cột trên biểu mẫu.
- **Chuẩn hóa UOM (Đơn vị tính):** Tự động quy đổi các đơn vị viết tắt nội bộ về chuẩn mã VNACCS của Hải quan thông qua một từ điển (Dictionary) định sẵn. *(Ví dụ: `PC`, `PCS` -> `PCE`; `KG`, `KGS` -> `KGM`; `YD` -> `YRD`; `M2` -> `MTK`)*.
- **Làm sạch chuỗi (String Clean):** Ép toàn bộ mã vật tư về chữ IN HOA và xóa khoảng trắng thừa (`.strip().upper()`), ngăn chặn lỗi vênh mã do lỗi đánh máy.

### Pha 3: Động cơ Quy đổi Master Data (HSQĐ Conversion Engine)
- Dữ liệu nguyên phụ liệu (đặc biệt là Vải/Mesh) thường được mua theo Yard/Inch nhưng Hải quan yêu cầu khai theo Mét vuông (MTK).
- Thuật toán hoạt động như một hàm VLOOKUP động: Quét Mã vật tư trên Invoice/PKL, nếu tồn tại trong file `Bảng Master HSQĐ`, hệ thống ngầm thực hiện phép tính: **`Số lượng thực tế` x `Hệ số quy đổi`** và tự động đổi Đơn vị báo quan thành chuẩn (MTK).

### Pha 4: Đối chiếu liên thông 5 chiều (5-Way Cross-Check & Full Outer Join)
Hệ thống sử dụng Mã vật tư (`Ma_Vat_Tu`) làm Trục xương sống (Key) để thực hiện lệnh nối bảng `Outer Join` trên Pandas.
- **Bóc tách ECUS:** Tờ khai Hải quan ECUS thường trộn lẫn mã nguyên liệu vào trong chuỗi Mô tả (Ví dụ: *NL32#&Hạt nhựa...*). Hệ thống dùng thuật toán Substring Matching quét đối chiếu với danh sách mã gốc để bóc tách chính xác Mã NL ra một cột riêng để merge.
- Cuối cùng, hệ thống dùng thuật toán GroupBy để cộng dồn số lượng nếu một mã vật tư bị chia ra làm nhiều dòng đóng gói (trên PKL).

### Pha 5: Bộ Đánh giá Logic Nghiệp vụ (HITL Rule Engine)
Hệ thống tính toán các độ lệch (Difference) và đưa qua bộ màng lọc Logic (Tolerance Rule):
1. **Lỗi Khuyết mã:** Có mã trên Invoice nhưng mất tích trên Tờ khai, hoặc ngược lại.
2. **Lỗi Nhân tiền:** Kiểm tra chéo `Số Lượng x Đơn Giá = Thành Tiền` (Cho phép sai số làm tròn 0.05 USD).
3. **Chấp nhận Dung sai (Tolerance):** Nhận diện hàng Cân ký (`KGM`) hoặc Mét vuông (`MTK`). Nếu độ lệch giữa thực tế cân đo (PKL/ERP) và chứng từ (Invoice) ≤ `0.2` thì đánh giá là **KHỚP (Vàng)** do làm tròn số thập phân.
4. **Phân loại Trách nhiệm (Color Coding):** - 🔴 **ĐỎ:** Lệch chứng từ (Lỗi do nhà cung cấp/phòng Mua hàng).
   - 🟠 **CAM:** Lệch SAP/ERP (Lỗi do bộ phận Kho thực tế).

### Pha 6: Output & Báo cáo (Reporting)
- Render giao diện DataFrame tích hợp bộ lọc nhanh (chỉ hiện lỗi).
- Xuất file Excel định dạng chuẩn, có thể nạp thẳng vào phần mềm VNACCS/ECUS.
- Render Biên bản đối chiếu tự động bằng PDF.

---

## 📁 CẤU TRÚC DỰ ÁN (PROJECT STRUCTURE)

```text
📁 E54_CHECKER_FRAMEWORK/
 ├── 📄 app.py               # Chứa luồng xử lý UI (Streamlit) và 6 Pha thuật toán
 ├── 📄 requirements.txt     # Các gói thư viện Python yêu cầu
 ├── 📄 packages.txt         # Cấu hình gói đồ họa Linux (dành cho OCR trên Cloud)
 ├── 📄 START_APP.bat        # Mã kịch bản (Script) khởi động cục bộ cho nhân viên
 ├── 📄 README.md            # Tài liệu kiến trúc và hướng dẫn sử dụng (File này)
 └── 📁 .streamlit/
      └── 📄 secrets.toml    # Lưu trữ mật khẩu an toàn (Bỏ qua bởi Git)
