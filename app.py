import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import pdfplumber
import easyocr
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import difflib
import plotly.express as px
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. GIAO DIỆN CHUẨN DOANH NGHIỆP & BẢO MẬT
# ==========================================
st.set_page_config(layout="wide", page_title="CLG SCM EXIM VN E54-E23 checker", page_icon="🌐", initial_sidebar_state="expanded")

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .stDeployButton {display:none;}
            .block-container {padding-top: 1rem; padding-bottom: 2rem;}
            [data-testid="stSidebar"] {background-color: #f8fafc;}
            .metric-card {background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); border-left: 5px solid #2563eb; margin-bottom: 15px;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

def check_password():
    if "auth" not in st.session_state: st.session_state["auth"] = False
    if not st.session_state["auth"]:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.markdown("<h2 style='text-align: center; color: #1E3A8A; padding-top: 50px;'>🌐 Chingluh Group - SCM EXIM VN - IMPORT TEAM 🌐</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center;'>Hệ thống liên thông dữ liệu Ching Luh & Fu-Luh khi kiểm tra E54 & E23 </p>", unsafe_allow_html=True)
            pwd = st.text_input("Nhập mã truy cập an ninh:", type="password")
            if pwd:
                if pwd == st.secrets.get("app_password"):
                    st.session_state["auth"] = True
                    st.rerun()
                else: st.error("❌ Mã truy cập không hợp lệ!")
        return False
    return True

if not check_password(): st.stop()

@st.cache_resource
def load_ocr_reader(): return easyocr.Reader(['vi', 'en'], gpu=False)
reader = load_ocr_reader()

# ==========================================
# 2. BỘ KHỬ NHIỄU & LÀM SẠCH DỮ LIỆU ĐỘC QUYỀN
# ==========================================
def clean_numeric_string(series):
    """Lọc bỏ toàn bộ dấu phẩy, chữ, khoảng trắng, chỉ giữ lại số thuần túy"""
    if series is None: return 0
    cleaned = series.astype(str).str.replace(',', '', regex=False)
    cleaned = cleaned.str.replace(r'[^\d\.]', '', regex=True)
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

def purify_code(s):
    """San phẳng chuỗi mã vật tư: Xóa bỏ dấu cách, gạch ngang để chống lệch VLOOKUP"""
    if pd.isna(s): return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).strip().upper())

def universal_file_extractor(uploaded_file, skiprows=0):
    if uploaded_file is None: return pd.DataFrame()
    ext = uploaded_file.name.split('.')[-1].lower()
    try:
        if ext in ['xlsx', 'xls', 'csv']:
            df = pd.read_csv(uploaded_file, skiprows=skiprows, on_bad_lines='skip') if ext == 'csv' else pd.read_excel(uploaded_file, skiprows=skiprows)
            # Khắc phục triệt để lỗi tên cột kiểu float/NaN của biểu thuế
            df.columns = [str(c).strip() for c in df.columns]
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all')
            return df
        elif ext == 'pdf':
            all_text_rows = []
            with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for row in table:
                                cr = [str(c).replace('\n', ' ') if c else '' for c in r]
                                if any(cr): all_text_rows.append(cr)
            if all_text_rows: return pd.DataFrame(all_text_rows[1:], columns=[str(c) for c in all_text_rows[0]])
    except Exception as e: st.error(f"Lỗi đọc tệp {uploaded_file.name}: {e}")
    return pd.DataFrame()

def get_col(df, kws, fb_idx):
    for c in df.columns:
        for k in kws:
            if k.lower() in str(c).lower(): return c
    return df.columns[fb_idx] if len(df.columns) > fb_idx else None

def fz_match(code, masters, master_purified_dict):
    p_code = purify_code(code)
    if p_code in master_purified_dict: return master_purified_dict[p_code]
    if code in masters: return code
    m = difflib.get_close_matches(code, masters, n=1, cutoff=0.82)
    return m[0] if m else code

# =====================================================================
# 3. EXPORT ENGINE (BỘ TRÍCH XUẤT ĐA ĐỊNH DẠNG AUTOMATION)
# =====================================================================
class ExportEngine:
    @staticmethod
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Sheet 1: Toàn bộ bảng dữ liệu sạch đã khớp
            df.to_excel(writer, index=False, sheet_name='Cross_Check_Result')
            wb = writer.book
            ws = writer.sheets['Cross_Check_Result']
            
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1, 'align': 'center'})
            fmt_red = wb.add_format({'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'border': 1})
            fmt_green = wb.add_format({'bg_color': '#D1FAE5', 'font_color': '#065F46', 'border': 1})
            fmt_num = wb.add_format({'num_format': '#,##0.00', 'border': 1})
            
            # Tự động căn chỉnh độ rộng của cột theo độ dài chuỗi ký tự (Chống lỗi ###)
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)
                max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 3
                ws.set_column(col_idx, col_idx, max_len)
                if df[col_name].dtype in [np.float64, np.int64]: ws.set_column(col_idx, col_idx, max_len, fmt_num)
            
            status_col_idx = len(df.columns) - 1
            ws.set_column(status_col_idx, status_col_idx, 40)
            for row_idx, status in enumerate(df.iloc[:, status_col_idx]):
                fmt = fmt_green if "🟢" in str(status) or "🟡" in str(status) else fmt_red
                ws.write(row_idx + 1, status_col_idx, str(status), fmt)
                
            # Sheet 2: Tách riêng các dòng BỊ LỖI để staff xử lý nhanh (Action Required)
            err_df = df[~df.iloc[:, status_col_idx].str.contains("🟢|🟡", regex=True)]
            if not err_df.empty:
                err_df.to_excel(writer, index=False, sheet_name='Action_Required_Errors')
                ws_err = writer.sheets['Action_Required_Errors']
                for col_idx, col_name in enumerate(err_df.columns):
                    ws_err.write(0, col_idx, col_name, fmt_header)
                    max_len = max(err_df[col_name].astype(str).map(len).max(), len(col_name)) + 3
                    ws_err.set_column(col_idx, col_idx, max_len)
        return output.getvalue()

    @staticmethod
    def to_word(df):
        doc = Document()
        doc.add_heading('BIÊN BẢN ĐỐI CHIẾU VÀ XÁC NHẬN SỐ LIỆU LÔ HÀNG', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Thời gian lập hồ sơ hệ thống: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        
        doc.add_heading('Danh sách các mặt hàng phát hiện sai lệch dữ liệu hệ thống cần làm rõ:', level=1)
        status_col_idx = len(df.columns) - 1
        err_df = df[~df.iloc[:, status_col_idx].str.contains("🟢|🟡", regex=True)]
        if err_df.empty:
            doc.add_paragraph("Kết quả đối chiếu: Tuyệt vời! Toàn bộ chứng từ gốc, sổ sách kho SAP và tờ khai hoàn toàn trùng khớp số lượng.")
        else:
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Mã Vật Tư (Material)', 'Số Lượng Invoice', 'Trạng Thái Sai Lệch'
            for _, r in err_df.head(25).iterrows():
                row = table.add_row().cells
                row[0].text, row[1].text, row[2].text = str(r.iloc[0]), str(r.iloc[2]), str(r.iloc[-1])
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()

# ==========================================
# 4. GIAO DIỆN ĐIỀU HÀNH MỘT MÀN HÌNH DUY NHẤT
# ==========================================
with st.sidebar:
    st.markdown("### ⚖️ THIẾT LẬP DUNG SAI SỐ THẬP PHÂN")
    tol_weight = st.slider("Dung sai hàng Vải/Cân ký (KGM, MTK):", min_value=0.0, max_value=1.0, value=0.2, step=0.05)
    tol_count = st.slider("Dung sai hàng Đếm chiếc (PCE, PRS, PR):", min_value=0.0, max_value=0.1, value=0.01, step=0.01)
    st.markdown("---")
    st.markdown("### ⚙️ CẤU HÌNH DÒNG TIÊU ĐỀ THỪA (SKIPROWS)")
    skip_inv = st.number_input("Dòng thừa Invoice/PKL:", value=15)
    skip_cd  = st.number_input("Dòng thừa Chỉ Định Giao Hàng:", value=17)
    skip_erp = st.number_input("Dòng thừa SAP (ZMM12):", value=0)
    skip_ecus = st.number_input("Dòng thừa Tờ Khai (Tab HANG):", value=7)

st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>💎 TRUNG TÂM KIỂM ĐỊNH LIÊN THÔNG CHỨNG TỪ XNK</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.05em;'>Hỗ trợ tự động hóa khép kín: Tách mã liệu R ➔ Quét sổ kho SAP ZMM12 ➔ Ráp chuỗi Hải quan #& ➔ Xuất form tờ khai đầu 30.</p>", unsafe_allow_html=True)
st.markdown("---")

# ==============================================================================
# PHASE 0: SAP QUERY HELPER (TRÍCH XUẤT MÃ LIỆU R TIỀN TRẠM)
# ==============================================================================
st.markdown("### ⚡ PHASE 0: TRÍCH XUẤT MÃ LIỆU R TIỀN TRẠM (SAP QUERY HELPER)")
st.caption("Nhận chứng từ thô từ kho gửi ➔ Thảy vào đây để lấy danh sách mã sạch đem đi truy vấn SAP ZMM12.")
col_f1, col_f2 = st.columns(2)
with col_f1: f_pre_inv = st.file_uploader("Tải lên Invoice nhận từ kho gửi:", type=["xlsx","csv","pdf"])
with col_f2: f_pre_pkl = st.file_uploader("Tải lên Packing List nhận từ kho gửi:", type=["xlsx","csv","pdf"])

if f_pre_inv or f_pre_pkl:
    with st.spinner("Đang gom mã vật tư và lọc bỏ trùng lặp..."):
        df_temp_i = universal_file_extractor(f_pre_inv, skip_inv)
        df_temp_p = universal_file_extractor(f_pre_pkl, skip_inv)
        
        raw_codes = []
        for df in [df_temp_i, df_temp_p]:
            if not df.empty:
                col_mat = get_col(df, ['Material code', 'Material', 'Mã'], None)
                if col_mat: raw_codes.extend(df[col_mat].dropna().astype(str).str.strip().str.upper().tolist())
        
        unique_r_codes = sorted(list(set([c for c in raw_codes if c and c != "NAN" and len(c) > 3])))
        if unique_r_codes:
            st.success(f"🎉 Hệ thống đã gom và lọc sạch trùng lặp! Trích xuất được {len(unique_r_codes)} Mã liệu R độc nhất của lô hàng.")
            codes_string = "\n".join(unique_r_codes)
            st.text_area("👇 Bấm VÀO Ô DƯỚI, nhấn Ctrl+A rồi Ctrl+C để copy danh sách đem dán thẳng vào ô Multi-Selection trên SAP:", value=codes_string, height=120)
        else: st.warning("Chưa quét được mã vật tư. Vui lòng kiểm tra lại cấu hình số dòng thừa (Skiprows).")

st.markdown("---")

# ==============================================================================
# PHASE 1 & 2: MULTI-DIMENSIONAL CROSS-CHECK ENGINE
# ==============================================================================
st.markdown("### 📥 PHASE 1 & 2: ĐỐI CHIẾU DIỆN RỘNG VÀ TỰ ĐỘNG GENERATE CHUỖI #& HẢI QUAN")
st.caption("Kéo thả tất cả các file sau khi đã lấy được báo cáo SAP ZMM12 và Tờ khai thô về máy.")

c1, c2, c3 = st.columns(3)
with c1:
    st.info("📄 **CHỨNG TỪ THƯƠNG MẠI GỐC**")
    f_inv_real = st.file_uploader("Upload Invoice Khai Báo Chính Thức:", type=["xlsx","csv"])
    f_pkl_real = st.file_uploader("Upload Packing List Khai Báo Chính Thức:", type=["xlsx","csv"])
    f_cd_real  = st.file_uploader("Upload Chỉ Định Giao Hàng (Nếu có):", type=["xlsx","csv"])
with c2:
    st.warning("💻 **SỔ SÁCH NỘI BỘ (SAP)**")
    f_sap_zmm = st.file_uploader("Upload Báo Cáo SAP ZMM12 vừa tải về:", type=["xlsx","csv"])
with c3:
    st.success("🏛️ **TỜ KHAI HẢI QUAN (ECUS)**")
    f_ecus_hang = st.file_uploader("Upload File Tờ khai Hải quan (Tab HANG):", type=["xlsx","csv","xls"])

if st.button("🚀 KÍCH HOẠT QUY TRÌNH ĐỐI CHIẾU TỔNG LỰC", type="primary", use_container_width=True):
    if not (f_inv_real and f_sap_zmm): st.error("⚠️ Điểm kiểm soát yêu cầu tối thiểu phải có Invoice chính thức và báo cáo SAP ZMM12.")
    else:
        with st.spinner("Hệ thống đang đồng bộ định dạng chuỗi và khớp nối ma trận mảng..."):
            r_inv = universal_file_extractor(f_inv_real, skip_inv)
            r_pkl = universal_file_extractor(f_pkl_real, skip_inv)
            r_cd  = universal_file_extractor(f_cd_real, skip_cd)
            r_sap = universal_file_extractor(f_sap_zmm, skip_erp)
            r_hq  = universal_file_extractor(f_ecus_hang, skip_ecus)

            # 1. Trích xuất Invoice (Trục Key gốc)
            c_inv_mat = get_col(r_inv, ['Material code', 'Material', 'Mã'], 1)
            df_inv = r_inv[[c_inv_mat, get_col(r_inv, ['Quantity Customs', 'Quantity', 'Số lượng'], 6), get_col(r_inv, ['Unit Customs', 'Unit', 'ĐVT'], 5), get_col(r_inv, ['Unit Price', 'DonGia', 'Price'], 7), get_col(r_inv, ['Amount', 'TriGia', 'Thành tiền'], 8)]].dropna(subset=[c_inv_mat]).copy()
            df_inv.columns = ['Ma_Vat_Tu', 'SL_INV', 'UOM_INV', 'DonGia', 'TriGia']
            df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].astype(str).str.strip().str.upper()
            for c in ['SL_INV', 'DonGia', 'TriGia']: df_inv[c] = clean_numeric_string(df_inv[c])
            
            # Khử nhiễu trùng dòng, cộng dồn Invoice trước khi check tránh mù dòng lặp
            df_inv = df_inv.groupby(['Ma_Vat_Tu', 'UOM_INV', 'DonGia'], as_index=False)[['SL_INV', 'TriGia']].sum()
            masters = df_inv['Ma_Vat_Tu'].tolist()
            master_purified_dict = {purify_code(m): m for m in masters}

            # 2. Trích xuất Packing List (Cộng dồn tự động theo Pallet)
            p_mat = get_col(r_pkl, ['Material', 'Mã'], 0)
            p_qty = get_col(r_pkl, ["Q'TY", 'Quantity', 'Số lượng'], 5)
            df_pkl = pd.DataFrame()
            df_pkl['Ma_Vat_Tu'] = r_pkl[p_mat].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
            df_pkl['SL_PKL'] = clean_numeric_string(r_pkl[p_qty])
            df_pkl = df_pkl.groupby('Ma_Vat_Tu', as_index=False)['SL_PKL'].sum()

            # 3. Trích xuất Chỉ Định Giao Hàng
            df_cd = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ChiDinh'])
            if not r_cd.empty:
                cd_mat = get_col(r_cd, ['Mã NL', 'Material', 'Mã'], 1)
                cd_qty = get_col(r_cd, ['Số Lượng', 'Quantity'], 4)
                df_cd = pd.DataFrame()
                df_cd['Ma_Vat_Tu'] = r_cd[cd_mat].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
                df_cd['SL_ChiDinh'] = clean_numeric_string(r_cd[cd_qty])
                df_cd = df_cd.groupby('Ma_Vat_Tu', as_index=False)['SL_ChiDinh'].sum()

            # 4. Trích xuất SAP ZMM12 (Tự động bóc tách Mã NL và Mô tả chuẩn từ SAP hệ thống)
            df_sap = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_SAP_ZMM', 'Ma_NL_HQ', 'Customs_Desc'])
            if not r_sap.empty:
                c_sap_mat = get_col(r_sap, ['Material'], 10)
                c_sap_qty = get_col(r_sap, ['PO NL QTY', 'In NL QTY', 'In Qty', 'PO Qty'], 13)
                c_sap_nl = get_col(r_sap, ['NL#'], 22)
                c_sap_desc = get_col(r_sap, ['CUSTOMS DESC'], 23)
                
                df_sap_raw = r_sap[[c_sap_mat, c_sap_qty, c_sap_nl, c_sap_desc]].dropna(subset=[c_sap_mat]).copy()
                df_sap_raw.columns = ['Ma_Vat_Tu', 'SL_SAP_ZMM', 'Ma_NL_HQ', 'Customs_Desc']
                df_sap_raw['Ma_Vat_Tu'] = df_sap_raw['Ma_Vat_Tu'].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
                df_sap_raw['SL_SAP_ZMM'] = clean_numeric_string(df_sap_raw['SL_SAP_ZMM'])
                df_sap = df_sap_raw.groupby(['Ma_Vat_Tu', 'Ma_NL_HQ', 'Customs_Desc'], as_index=False)['SL_SAP_ZMM'].sum()

            # 5. Trích xuất Tờ Khai Hải Quan thực tế (Tab HANG)
            df_hq = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_HAI_QUAN'])
            if not r_hq.empty:
                c_hq_mat = get_col(r_hq, ['Mã số hàng hóa', 'Mã số'], 0)
                c_hq_qty = get_col(r_hq, ['Số lượng (1)', 'Lượng'], 3)
                # Parse logic cho file dạng tờ khai Hải quan
                df_hq = r_hq[[c_hq_mat, c_hq_qty]].dropna().copy()
                df_hq.columns = ['Ma_Vat_Tu', 'SL_HAI_QUAN']
                df_hq['SL_HAI_QUAN'] = clean_numeric_string(df_hq['SL_HAI_QUAN'])
                df_hq = df_hq.groupby('Ma_Vat_Tu', as_index=False)['SL_HAI_QUAN'].sum()

            # LẬP MA TRẬN ĐỐI CHIẾU ĐA CHIỀU (FULL OUTER JOIN)
            mg = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_cd, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_sap, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_hq, on='Ma_Vat_Tu', how='outer').fillna(0)

            # THUẬT TOÁN AUTOMATION: TỰ ĐỘNG LẮP GHÉP CHUỖI #& HẢI QUAN CHO TỜ KHAI ĐẦU 30
            # Sẽ lấy Mã NL (ví dụ: NL17) ghép với dấu #& và Customs Desc (ví dụ: Miếng lót đệm)
            mg['MÔ TẢ TOÀN DIỆN HẢI QUAN (FORM 30+)'] = mg['Ma_NL_HQ'].astype(str) + "#&" + mg['Customs_Desc'].astype(str)

            # THUẬT TOÁN ĐÁNH GIÁ TRẠNG THÁI THEO DUNG SAI ĐỘNG
            def evaluate_row(r):
                if r['SL_INV'] == 0: return "❌ LỖI: CHỈ CÓ TRÊN ĐẦU KHO (THIẾU TRÊN INV)"
                tol = tol_weight if r['UOM_INV'] in ['KGM', 'MTK'] else tol_count
                
                if abs(r['SL_INV'] - r['SL_PKL']) > tol: return "🔴 LỆCH SỐ LIỆU ĐÓNG GÓI PACKING LIST"
                if 'SL_ChiDinh' in r and r['SL_ChiDinh'] > 0 and abs(r['SL_INV'] - r['SL_ChiDinh']) > tol: return "🔴 LỆCH CHỈ ĐỊNH GIAO HÀNG"
                if 'SL_SAP_ZMM' in r and abs(r['SL_INV'] - r['SL_SAP_ZMM']) > tol: return "🚨 LỖI THẤT THOÁT: VÊNH TỒN KHO NỘI BỘ SAP"
                if 'SL_HAI_QUAN' in r and r['SL_HAI_QUAN'] > 0 and abs(r['SL_INV'] - r['SL_HAI_QUAN']) > tol: return "🔴 LỆCH THỰC TẾ KHAI TỜ KHAI HQ"
                return "🟢 CHỨNG TỪ KHỚP HOÀN TOÀN"

            mg['KẾT LUẬN KIỂM CHÉO'] = mg.apply(evaluate_row, axis=1)
            mg = mg.sort_values(by='KẾT LUẬN KIỂM CHÉO', ascending=False)

            # HIỂN THỊ KẾT QUẢ RA MÀN HÌNH CHÍNH
            st.markdown("---")
            st.markdown("### 📊 BẢNG VÀNG ĐỐI CHIẾU LIÊN THÔNG TOÀN CỤC")
            err_filter = st.toggle("🚨 BẬT BỘ LỌC LỖI (Ẩn toàn bộ dòng đúng để tập trung sửa dòng sai)", value=False)
            final_board = mg[~mg['KẾT LUẬN KIỂM CHÉO'].str.contains("🟢", regex=False)] if err_filter else mg
            
            st.data_editor(
                final_board.style.applymap(lambda x: 'background-color:#d1fae5; color:#065f46' if '🟢' in str(x) else 'background-color:#7f1d1d; color:#fca5a5; font-weight:bold' if '🚨' in str(x) else 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🔴' in str(x) or '❌' in str(x) else '', subset=['KẾT LUẬN KIỂM CHÉO']),
                use_container_width=True, hide_index=True, height=400
            )

            # BỘ NÚT TRÍCH XUẤT ĐA ĐỊNH DẠNG TẬP TRUNG (ĂN NGAY)
            st.markdown("#### 📥 TRÍCH XUẤT HỒ SƠ LÔ HÀNG ĐA ĐỊNH DẠNG")
            d1, d2 = st.columns(2)
            with d1:
                st.download_button("📊 TẢI EXCEL ĐỐI CHIẾU CHÉO (Có Summary & Sheet Lỗi riêng)", ExportEngine.to_excel(mg), "KetQua_CrossCheck_Master.xlsx", "primary", use_container_width=True)
            with d2:
                st.download_button("📝 TẢI BIÊN BẢN WORD KÝ DUYỆT TRÌNH GIÁM ĐỐC", ExportEngine.to_word(mg), "BienBan_DoiChieu_SCM.docx", use_container_width=True)
