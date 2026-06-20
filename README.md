🏢 CLG SCM EXIM VN - E23 & E54 Reconciliation Tool
Hệ thống kiểm soát và đối chiếu liên thông chứng từ xuất nhập khẩu (Loại hình E23 - Nhập gia công / E54 - Xuất gia công). Ứng dụng giúp loại bỏ hoàn toàn các thao tác VLOOKUP thủ công, tự động phát hiện sai lệch số lượng và tự động sinh chuỗi dữ liệu khai báo Hải quan.

🌟 Tính năng Cốt lõi (Key Features)
Smart File Extractor: Tự động quét và tìm dòng tiêu đề (Header) của các file Excel/CSV/PDF, nhân viên không cần thiết lập skiprows thủ công. Hỗ trợ đọc cả file PDF và Ảnh (bằng AI OCR).

Data Sanitization: Tự động khử nhiễu dữ liệu. Loại bỏ chữ/ký tự lạ dính trong cột số lượng; Xóa bỏ các đuôi thập phân lỗi (.0) và khoảng trắng trong Mã vật tư để khớp lệnh chính xác tuyệt đối.

Multi-dimensional Cross-Check: Đối chiếu ma trận 5 chiều (Invoice ↔ Packing List ↔ Chỉ Định ↔ SAP ZMM12 ↔ ECUS Hải Quan).

Customs String Auto-Generator: Tự động lắp ghép chuỗi mô tả truyền tờ khai Hải Quan (Đầu 30+) theo đúng cú pháp [Mã NL]#&[Mô tả].

Dynamic Tolerance: Thanh trượt tùy chỉnh dung sai cho phép lệch (ví dụ: ±0.2 KG cho hàng cân ký, 0 cho hàng đếm chiếc).

Multi-language & Export: Hỗ trợ 3 ngôn ngữ (VI, EN, ZH). Xuất báo cáo Excel (tự động căn cột, tô màu lỗi) và biên bản Word trình ký.

🔄 Luồng Vận hành (Workflow)
Hệ thống được thiết kế theo luồng tác chiến tuần tự của nhân viên kho:

Phase 0 (SAP Query Helper): * Tải chứng từ thô (Invoice/PKL) lên.

Hệ thống bóc tách, lọc trùng và nhả ra danh sách Mã vật tư (R-codes) sạch.

Mục đích: Nhân viên copy list này dán vào SAP để kết xuất báo cáo ZMM12 nhanh chóng.

Phase 1 (Actual Delivery Audit): * Tải bộ chứng từ chính thức (Invoice, PKL, Chỉ định) + Báo cáo SAP ZMM12 vừa tải về.

Hệ thống trừ cấn trừ số lượng, phát hiện chênh lệch đóng gói và hao hụt tồn kho.

Phase 2 (Customs Sync & Export): * Tải file Tờ khai Hải quan (ECUS - Tab HANG).

Đối chiếu chốt chặn cuối cùng. Tải báo cáo Excel & Biên bản Word.

🛠️ Cài đặt & Triển khai (Installation)
1. Yêu cầu hệ thống (Requirements)
Cài đặt các thư viện cần thiết thông qua tệp requirements.txt:

Bash
pip install -r requirements.txt
2. Cấu hình Bảo mật (Security Secrets)
Ứng dụng sử dụng cơ chế bảo mật Password Gate. Bắt buộc phải tạo file secrets trước khi chạy để tránh lỗi KeyError.

Chạy Local: Tạo thư mục .streamlit/ ở thư mục gốc của project, tạo file secrets.toml bên trong và thêm dòng sau:

Ini, TOML
app_password = "Mat_khau_cua_ban"
Chạy trên Streamlit Cloud: Vào Settings của App -> mục Secrets -> Dán cấu hình trên vào.

3. Khởi chạy Ứng dụng

Bash
streamlit run app.py
🔒 Bảo mật dữ liệu (Data Privacy)
In-Memory Processing: Ứng dụng xử lý mọi tệp chứng từ trực tiếp trên RAM (Volatility). Không có bất kỳ tệp dữ liệu gốc hay báo cáo nào của công ty được lưu trữ xuống ổ cứng máy chủ.

Zero-Hardcode: Mật khẩu truy cập được cô lập hoàn toàn khỏi mã nguồn trên GitHub.

Developed for internal workflow optimization at CLG SCM EXIM VN.
