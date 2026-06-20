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
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN CHUẨN ĐƠN TRANG (LIGHT MODE CỐ ĐỊNH)
# ==========================================
st.set_page_config(layout="wide", page_title="CLG SCM EXIM VN E23 E54 Checker", page_icon="🌐", initial_sidebar_state="expanded")

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .stDeployButton {display:none;}
            .block-container {padding-top: 1rem; padding-bottom: 2rem; background-color: #ffffff; color: #0f172a;}
            .stApp {background-color: #ffffff; color: #0f172a;}
            
            [data-testid="stSidebar"] {background-color: #f1f5f9 !important;}
            [data-testid="stSidebar"] p, 
            [data-testid="stSidebar"] span, 
            [data-testid="stSidebar"] label, 
            [data-testid="stSidebar"] h3 {
                color: #0f172a !important;
                font-weight: 600;
            }
            
            .metric-card {background: #f8fafc; padding: 15px; border-radius: 10px; border: 1px solid #e2e8f0; border-left: 5px solid #2563eb; margin-bottom: 15px; color: #0f172a;}
            .phase-box {border: 1px solid #e2e8f0; padding: 20px; border-radius: 8px; background-color: #f8fafc; margin-bottom: 25px;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

def check_password():
    if "auth" not in st.session_state: st.session_state["auth"] = False
    if not st.session_state["auth"]:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.markdown("<h3 style='text-align: center; color: #1E3A8A; padding-top: 50px;'>🌐 CLG SCM EXIM VN INTERNALS</h3>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; font-size:0.9em; color:#555;'>Công cụ đối soát liên thông chứng từ phục vụ nhóm nghiệp vụ nhập kho</p>", unsafe_allow_html=True)
            pwd = st.text_input("Nhập mã an ninh nội bộ:", type="password")
            if pwd:
                if pwd == st.secrets.get("app_password", "ChingLuh@2026"):
                    st.session_state["auth"] = True
                    st.rerun()
                else: st.error("❌ Mã truy cập không hợp lệ!")
        return False
    return True

if not check_password(): st.stop()

@st.cache_resource
def load_ocr_reader(): return easyocr.Reader(['vi', 'en'], gpu=False)
reader = load_ocr_reader()

if 'master_hsqd' not in st.session_state: st.session_state.master_hsqd = pd.DataFrame()
if 'master_thue' not in st.session_state: st.session_state.master_thue = pd.DataFrame()
if 'baseline_e21_pool' not in st.session_state: st.session_state.baseline_e21_pool = pd.DataFrame()
if 'data_phase1' not in st.session_state: st.session_state.data_phase1 = pd.DataFrame()

# ==========================================
# 2. ĐỘNG CƠ KHỬ NHIỄU & RÀO TRƯỚC LỖI FORMAT CỰC ĐOAN
# ==========================================
class DataSanitizer:
    @staticmethod
    def clean_numeric(series):
        if series is None: return 0
        cleaned = series.astype(str).str.replace(',', '', regex=False)
        cleaned = cleaned.str.replace(r'[^\d\.]', '', regex=True)
        return pd.to_numeric(cleaned, errors='coerce').fillna(0)

    @staticmethod
    def purify_material_code(val):
        """RÀO TRƯỚC LỖI FORMAT: Loại bỏ khoảng trắng, ép viết hoa, phá đuôi .0 thập phân lỗi của Excel"""
        if pd.isna(val): return ""
        s = str(val).strip().split('.')[0] # Cắt phăng phần thập phân .0 nếu bị nhận diện nhầm thành số float
        return re.sub(r'[^A-Z0-9]', '', s.upper())

class SmartFileReader:
    """Động cơ đọc vạn năng: Tự căn dòng tiêu đề chuẩn xác, không cần nhân viên chọn Skiprows thủ công"""
    @staticmethod
    def read_smart(uploaded_file, target_keywords):
        if uploaded_file is None: return pd.DataFrame()
        ext = uploaded_file.name.split('.')[-1].lower()
        try:
            if ext in ['xlsx', 'xls', 'csv']:
                # Bước 1: Đọc thô từ dòng đầu tiên không bỏ sót
                df_raw = pd.read_csv(uploaded_file, skiprows=0, on_bad_lines='skip') if ext == 'csv' else pd.read_excel(uploaded_file, skiprows=0)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                
                # Bước 2: Thuật toán quét 25 dòng đầu tìm hàng tiêu đề thực tế chứa từ khóa gốc
                header_row_idx = None
                for i, row in df_raw.head(25).iterrows():
                    combined_row_str = " ".join(row.dropna().astype(str).lower())
                    if any(kw.lower() in combined_row_str for kw in target_keywords):
                        header_row_idx = i
                        break
                
                # Bước 3: Tái cấu trúc lại bảng dữ liệu theo dòng tiêu đề vừa tìm được
                if header_row_idx is not None:
                    actual_headers = df_raw.iloc[header_row_idx].astype(str).str.strip().tolist()
                    df_clean = df_raw.iloc[header_row_idx + 1:].copy()
                    df_clean.columns = actual_headers
                    df_clean = df_clean.loc[:, ~df_clean.columns.str.contains('^nan|^Unnamed', case=False)].dropna(how='all')
                    return df_clean
                return df_raw.dropna(how='all')
                
            elif ext == 'pdf':
                all_rows = []
                with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        if tables:
                            for table in tables:
                                for row in table:
                                    cr = [str(c).strip().replace('\n', ' ') if c else '' for c in row]
                                    if any(cr): all_rows.append(cr)
                if all_rows: return pd.DataFrame(all_rows[1:], columns=[str(c) for c in all_rows[0]])
        except Exception as e:
            st.error(f"Lỗi phân rã cấu trúc tệp {uploaded_file.name}: {e}")
        return pd.DataFrame()

def get_col(df, kws, fb_idx):
    for c in df.columns:
        for kw in kws:
            if kw.lower() in str(c).lower(): return c
    return df.columns[fb_idx] if len(df.columns) > fb_idx else None

def fz_match(code, masters, master_purified_dict):
    p_code = DataSanitizer.purify_material_code(code)
    if p_code in master_purified_dict: return master_purified_dict[p_code]
    if code in masters: return code
    m = difflib.get_close_matches(code, masters, n=1, cutoff=0.82)
    return m[0] if m else code

def parse_ecus(raw_df, master_codes, prefix):
    if raw_df.empty: return pd.DataFrame(columns=['Ma_Vat_Tu', f'SL_{prefix}', 'HS_Code'])
    d_col = get_col(raw_df, ['Mô tả', 'Tên hàng'], 1)
    q_col = get_col(raw_df, ['Số lượng', 'Lượng tính thuế'], 3)
    h_col = get_col(raw_df, ['Mã số', 'HS'], 0)
    
    tmp = raw_df[[d_col, q_col, h_col]].dropna(subset=[d_col]).copy()
    codes = []
    for _, r in tmp.iterrows():
        desc = str(r[d_col]).upper()
        found = "UNK"
        for c in master_codes:
            if purify_code_local := DataSanitizer.purify_material_code(c):
                if purify_code_local in re.sub(r'[^A-Z0-9]', '', desc):
                    found = c; break
        codes.append(found)
    tmp['Ma_Vat_Tu'] = codes
    res = tmp[tmp['Ma_Vat_Tu'] != 'UNK'].copy()
    res[f'SL_{prefix}'] = DataSanitizer.clean_numeric(res[q_col])
    res.rename(columns={h_col: 'HS_Code'}, inplace=True)
    return res.groupby(['Ma_Vat_Tu', 'HS_Code'], as_index=False)[f'SL_{prefix}'].sum()

class ExportEngine:
    @staticmethod
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Consolidated_Audit_Result')
            wb = writer.book
            ws = writer.sheets['Consolidated_Audit_Result']
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1, 'align': 'center'})
            fmt_red = wb.add_format({'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'border': 1})
            fmt_green = wb.add_format({'bg_color': '#D1FAE5', 'font_color': '#065F46', 'border': 1})
            fmt_num = wb.add_format({'num_format': '#,##0.00', 'border': 1})
            
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)
                max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 3
                ws.set_column(col_idx, col_idx, max_len)
                if df[col_name].dtype in [np.float64, np.int64]: ws.set_column(col_idx, col_idx, max_len, fmt_num)
            
            status_col_idx = len(df.columns) - 1
            ws.set_column(status_col_idx, status_col_idx, 45)
            for row_idx, status in enumerate(df.iloc[:, status_col_idx]):
                fmt = fmt_green if "🟢" in str(status) or "🟡" in str(status) else fmt_red
                ws.write(row_idx + 1, status_col_idx, str(status), fmt)
        return output.getvalue()

# Thanh dung sai tại Sidebar
with st.sidebar:
    st.markdown("### ⚖️ THIẾT LẬP DUNG SAI ĐỘNG")
    tol_weight = st.slider("Hàng Vải/Cân ký (KGM, MTK):", min_value=0.0, max_value=1.0, value=0.2, step=0.05)
    tol_count = st.slider("Hàng Đếm chiếc (PCE, PRS, PR):", min_value=0.0, max_value=0.1, value=0.01, step=0.01)

# ==============================================================================
# PHASE 0: THIẾT LẬP BASELINE THANH KHOẢN (E21 IMPORT POOL) & TRÍCH XUẤT MÃ R
# ==============================================================================
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown("#### ⚡ PHASE 0: PHÒNG VỆ SỐ LIỆU TỒN KHO GỐC E21 & TRÍCH XUẤT MÃ LIỆU R TIỀN TRẠM")
st.caption("Ứng dụng tự động quét tìm dòng tiêu đề. Điền đầy đủ dữ liệu mốc E21 ban đầu để tạo lá chắn thanh khoản.")

c0_1, c0_2, c0_3 = st.columns(3)
with c0_1:
    f_pre_inv = st.file_uploader("1. Tải lên Invoice kho gửi (Tệp thô):", type=["xlsx","csv","pdf"])
    f_pre_pkl = st.file_uploader("2. Tải lên Packing List kho gửi (Tệp thô):", type=["xlsx","csv","pdf"])
with c0_2:
    f_ref_e21 = st.file_uploader("3. BÁO CÁO TỜ KHAI NHẬP KHẨU E21 GỐC BAN ĐẦU (Quản trị thanh khoản):", type=["xlsx","csv"])
with c0_3:
    f_hsqd = st.file_uploader("4. Bảng hệ số quy đổi định mức (HSQĐ):", type=["xlsx","csv"])
    f_thue = st.file_uploader("5. Bảng biểu thuế (Tra cứu mã HS):", type=["xlsx","csv"])

# Tự động lưu Master Data
if f_hsqd:
    df_h = SmartFileReader.read_smart(f_hsqd, ['Mã', 'Hệ số', 'Đơn vị'])
    if not df_h.empty:
        df_h = df_h[[get_col(df_h, ['Mã'], 1), get_col(df_h, ['Hệ số'], 6), get_col(df_h, ['Đơn vị'], 8)]].dropna()
        df_h.columns = ['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan']
        df_h['Ma_Vat_Tu'] = df_h['Ma_Vat_Tu'].apply(DataSanitizer.purify_material_code)
        st.session_state.master_hsqd = df_h

if f_thue:
    df_t = SmartFileReader.read_smart(f_thue, ['Mã hàng', 'HS', 'Thuế suất'])
    if not df_t.empty:
        df_t = df_t[[get_col(df_t, ['Mã hàng', 'HS'], 1), get_col(df_t, ['Thuế suất'], 5)]].dropna()
        df_t.columns = ['HS_Code', 'Thue_Suat']
        df_t['HS_Code'] = df_t['HS_Code'].astype(str).str.replace('.', '', regex=False).str.strip()
        st.session_state.master_thue = df_t

# Kiểm soát mốc tồn Hải quan E21 gốc ban đầu
if f_ref_e21:
    df_e21_raw = SmartFileReader.read_smart(f_ref_e21, ['Mã NPL', 'Material', 'Mã hàng', 'Số lượng'])
    if not df_e21_raw.empty:
        c_e21_mat = get_col(df_e21_raw, ['Mã NPL', 'Material', 'Mã vật tư'], 1)
        c_e21_qty = get_col(df_e21_raw, ['Tổng số lượng', 'Quantity', 'Số lượng'], 3)
        df_e21_clean = df_e21_raw[[c_e21_mat, c_e21_qty]].dropna().copy()
        df_e21_clean.columns = ['Ma_Vat_Tu', 'SL_E21_GOC']
        df_e21_clean['Ma_Vat_Tu'] = df_e21_clean['Ma_Vat_Tu'].apply(DataSanitizer.purify_material_code)
        df_e21_clean['SL_E21_GOC'] = DataSanitizer.clean_numeric(df_e21_clean['SL_E21_GOC'])
        st.session_state.baseline_e21_pool = df_e21_clean.groupby('Ma_Vat_Tu', as_index=False)['SL_E21_GOC'].sum()
        st.toast("✅ ĐÃ ĐÓNG BĂNG BỂ CHỨA TỒN THANH KHOẢN GỐC E21!")

# Output nhả List mã R không trùng lắp để staff đưa vào SAP
if f_pre_inv or f_pre_pkl:
    df_temp_i = SmartFileReader.read_smart(f_pre_inv, ['Material code', 'Material', 'Mã'])
    df_temp_p = SmartFileReader.read_smart(f_pre_pkl, ['Material code', 'Material', 'Mã'])
    raw_codes = []
    for d_node in [df_temp_i, df_temp_p]:
        if not d_node.empty:
            col_mat = get_col(d_node, ['Material code', 'Material', 'Mã'], None)
            if col_mat: raw_codes.extend(d_node[col_mat].dropna().astype(str).tolist())
    unique_r_codes = sorted(list(set([DataSanitizer.purify_material_code(c) for c in raw_codes if c and str(c).lower() != "nan"])))
    if unique_r_codes:
        st.success(f"📋 Động cơ dò tìm tự động phát hiện {len(unique_r_codes)} Mã liệu R độc nhất từ chứng từ thô.")
        st.text_area("👉 Nhấn Ctrl+A rồi Ctrl+C để copy danh sách mã liệu R dán vào mục Multi-Selection trên SAP (ZMM12):", value="\n".join(unique_r_codes), height=100)
st.markdown("</div>", unsafe_allow_html=True)

# ==============================================================================
# PHASE 1: ĐỐI SOÁT MA TRẬN GIAO NHẬN THỰC TẾ & BÓC TÁCH CHÊNH LỆCH CHI TIẾT
# ==============================================================================
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown("#### 📦 PHASE 1: ĐỐI SOÁT MA TRẬN GIAO NHẬN THỰC TẾ & KHẤP CHÉO SỔ KHO")
st.caption("Hệ thống tự động chạy phép trừ số lượng đa chiều để phát hiện chính xác khối lượng hao hụt.")

c1_1, c1_2, c1_3 = st.columns(3)
with c1_1:
    f_inv_real = st.file_uploader("1. INVOICE KHAI BÁO CHÍNH THỨC:", type=["xlsx","csv"])
    f_pkl_real = st.file_uploader("2. PACKING LIST KHAI BÁO CHÍNH THỨC:", type=["xlsx","csv"])
with c1_2:
    f_cd_real = st.file_uploader("3. CHỈ ĐỊNH GIAO HÀNG (Lịch điều xe):", type=["xlsx","csv"])
with c1_3:
    f_sap_zmm = st.file_uploader("4. BÁO CÁO KHO NỘI BỘ SAP ZMM12 VỀ MÁY:", type=["xlsx","csv"])

if f_inv_real and f_sap_zmm:
    with st.spinner("Đang chạy đối soát đa diện toán học cực đoan..."):
        r_inv = SmartFileReader.read_smart(f_inv_real, ['Material code', 'Material', 'Quantity Customs', 'Unit Price'])
        r_pkl = SmartFileReader.read_smart(f_pkl_real, ['Material', 'Quantity', "Q'TY"])
        r_cd  = SmartFileReader.read_smart(f_cd_real, ['Mã NL', 'Material', 'Số Lượng'])
        r_sap = SmartFileReader.read_smart(f_sap_zmm, ['PO Number', 'Material', 'PO Qty', 'NL#'])

        # Invoice
        c_inv_mat = get_col(r_inv, ['Material code', 'Material', 'Mã'], 1)
        df_inv = r_inv[[c_inv_mat, get_col(r_inv, ['Quantity', 'Số lượng'], 6), get_col(r_inv, ['Unit', 'ĐVT'], 5), get_col(r_inv, ['Price', 'Đơn giá'], 7), get_col(r_inv, ['Amount', 'Thành tiền'], 8)]].dropna(subset=[c_inv_mat]).copy()
        df_inv.columns = ['Ma_Vat_Tu', 'SL_INV', 'UOM_INV', 'DonGia', 'TriGia']
        df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].apply(DataSanitizer.purify_material_code)
        for c in ['SL_INV', 'DonGia', 'TriGia']: df_inv[c] = DataSanitizer.clean_numeric(df_inv[c])
        df_inv = df_inv.groupby(['Ma_Vat_Tu', 'UOM_INV', 'DonGia'], as_index=False)[['SL_INV', 'TriGia']].sum()
        
        masters = df_inv['Ma_Vat_Tu'].tolist()
        master_purified_dict = {purify_code(m): m for m in masters}

        # Kiểm định tính đồng nhất toán học trên Invoice
        df_inv['Math_TriGia_Check'] = np.where(abs((df_inv['SL_INV'] * df_inv['DonGia']) - df_inv['TriGia']) > 0.05, "🚨 SAI TRỊ GIÁ", "🟢 Khớp giá")

        # Packing List
        p_mat = get_col(r_pkl, ['Material', 'Mã'], 0)
        p_qty = get_col(r_pkl, ["Q'TY", 'Quantity'], 5)
        df_pkl = pd.DataFrame()
        df_pkl['Ma_Vat_Tu'] = r_pkl[p_mat].astype(str).str.strip().upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
        df_pkl['SL_PKL'] = DataSanitizer.clean_numeric(r_pkl[p_qty])
        df_pkl = df_pkl.groupby('Ma_Vat_Tu', as_index=False)['SL_PKL'].sum()

        # Chỉ định giao hàng
        df_cd = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ChiDinh'])
        if not r_cd.empty:
            cd_mat = get_col(r_cd, ['Mã NL', 'Material'], 1)
            cd_qty = get_col(r_cd, ['Số Lượng', 'Quantity'], 4)
            df_cd = pd.DataFrame()
            df_cd['Ma_Vat_Tu'] = r_cd[cd_mat].astype(str).str.strip().upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
            df_cd['SL_ChiDinh'] = DataSanitizer.clean_numeric(r_cd[cd_qty])
            df_cd = df_cd.groupby('Ma_Vat_Tu', as_index=False)['SL_ChiDinh'].sum()

        # SAP ZMM12
        c_sap_mat = get_col(r_sap, ['Material'], 10)
        c_sap_qty = get_col(r_sap, ['PO NL QTY', 'In NL QTY', 'In Qty'], 13)
        c_sap_nl = get_col(r_sap, ['NL#'], 22)
        c_sap_desc = get_col(r_sap, ['CUSTOMS DESC'], 23)
        df_sap_raw = r_sap[[c_sap_mat, c_sap_qty, c_sap_nl, c_sap_desc]].dropna(subset=[c_sap_mat]).copy()
        df_sap_raw.columns = ['Ma_Vat_Tu', 'SL_SAP_ZMM', 'Ma_NL_HQ', 'Customs_Desc']
        df_sap_raw['Ma_Vat_Tu'] = df_sap_raw['Ma_Vat_Tu'].apply(DataSanitizer.purify_material_code)
        df_sap_raw['SL_SAP_ZMM'] = DataSanitizer.clean_numeric(df_sap_raw['SL_SAP_ZMM'])
        df_sap = df_sap_raw.groupby(['Ma_Vat_Tu', 'Ma_NL_HQ', 'Customs_Desc'], as_index=False)['SL_SAP_ZMM'].sum()

        # Gộp bảng diện rộng
        mg_p1 = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
        mg_p1 = pd.merge(mg_p1, df_cd, on='Ma_Vat_Tu', how='outer')
        mg_p1 = pd.merge(mg_p1, df_sap, on='Ma_Vat_Tu', how='outer')
        if not st.session_state.baseline_e21_pool.empty:
            mg_p1 = pd.merge(mg_p1, st.session_state.baseline_e21_pool, on='Ma_Vat_Tu', how='left')
        mg_p1 = mg_p1.fillna(0)

        # PHÂN TÍCH ĐI SÂU CHI TIẾT CHECK: Đập nhỏ cột số liệu chênh lệch trực diện
        mg_p1['Chênh_Lệch_PKL'] = mg_p1['SL_PKL'] - mg_p1['SL_INV']
        mg_p1['Chênh_Lệch_SAP'] = mg_p1['SL_SAP_ZMM'] - mg_p1['SL_INV']
        
        if 'SL_E21_GOC' in mg_p1.columns:
            mg_p1['Dư_ThanhKhoản_E21'] = mg_p1['SL_E21_GOC'] - mg_p1['SL_INV']

        def eval_p1(r):
            if r['SL_INV'] == 0: return "❌ LỖI: THIẾU TRÊN INVOICE GỐC"
            if "🚨" in str(r['Math_TriGia_Check']): return "🚨 FATAL: SAI PHÉP NHÂN TRỊ GIÁ TRÊN INVOICE"
            tol = tol_weight if r['UOM_INV'] in ['KGM', 'MTK'] else tol_count
            if abs(r['Chênh_Lệch_PKL']) > tol: return "🔴 LỆCH SỐ LIỆU ĐÓNG GÓI PACKING LIST"
            if abs(r['Chênh_Lệch_SAP']) > tol: return "🚨 LỖI THẤT THOÁT: VÊNH TỒN KHO NỘI BỘ SAP"
            if 'Dư_ThanhKhoản_E21' in r and r['Dư_ThanhKhoản_E21'] < 0: return "🚨 NGUY HIỂM: VƯỢT ĐỊNH MỨC THANH KHOẢN E21 GỐC"
            return "🟢 GIAO NHẬN THỰC TẾ KHỚP HOÀN TOÀN"

        mg_p1['KẾT LUẬN PHASE 1'] = mg_p1.apply(eval_p1, axis=1)
        st.markdown("##### 📊 KẾT QUẢ PHÂN TÍCH CHÊNH LỆCH GIAO NHẬN THỰC TẾ & THANH KHOẢN GỐC")
        st.dataframe(mg_p1, use_container_width=True, hide_index=True)
        st.session_state.data_p1 = mg_p1
st.markdown("</div>", unsafe_allow_html=True)

# ==============================================================================
# PHASE 2: ĐỐI SOÁT TỜ KHAI HẢI QUAN & TỰ ĐỘNG GENERATE CHUỖI MÔ TẢ ĐẦU 30
# ==============================================================================
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown("#### 🏛️ PHASE 2: ĐỒNG BỘ TỜ KHAI HẢI QUAN & TRÍCH XUẤT FILE REPORT")
st.caption("Cross-check chiều cuối cùng với dữ liệu tờ khai Hải quan để xuất báo cáo trình ký.")

f_ecus_hang = st.file_uploader("1. TẢI LÊN FILE TỜ KHAI HẢI QUAN (TAB HANG):", type=["xlsx","csv","xls"])

if f_ecus_hang and not st.session_state.data_p1.empty:
    with st.spinner("Đang chạy đối toán luồng Hải quan..."):
        mg_p2_goc = st.session_state.data_p1.copy()
        r_hq = SmartFileReader.read_smart(f_ecus_hang, ['Mã số hàng hóa', 'Mô tả hàng hóa', 'Số lượng (1)'])
        
        masters_p2 = mg_p2_goc['Ma_Vat_Tu'].tolist()
        df_hq = parse_ecus(r_hq, masters_p2, "HAI_QUAN")

        mg_final = pd.merge(mg_p2_goc, df_hq, on='Ma_Vat_Tu', how='outer').fillna(0)

        # THUẬT TOÁN AUTOMATION ĐỘC QUYỀN: TỰ ĐỘNG RÁP CHUỖI MÔ TẢ TỜ KHAI ĐẦU 30 CHO STAFF
        mg_final['MÔ TẢ TOÀN DIỆN HẢI QUAN (FORM 30+)'] = mg_final['Ma_NL_HQ'].astype(str) + "#&" + mg_final['Customs_Desc'].astype(str)
        mg_final['Chênh_Lệch_KhaiBáo_HQ'] = mg_final['SL_HAI_QUAN'] - mg_final['SL_INV']

        mg_final['Thue_Du_Kien'] = 0.0
        if not st.session_state.master_thue.empty and 'HS_Code' in mg_final.columns:
            tax_dict = dict(zip(st.session_state.master_thue['HS_Code'], st.session_state.master_thue['Thue_Suat']))
            mg_final['Thue_NK_Suat'] = mg_final['HS_Code'].astype(str).str.replace('.','', regex=False).map(tax_dict).fillna(0)
            mg_final['Thue_Du_Kien'] = mg_final['TriGia'] * pd.to_numeric(mg_final['Thue_NK_Suat'], errors='coerce').fillna(0) / 100

        def eval_p2(r):
            if "❌" in str(r['KẾT LUẬN PHASE 1']) or "🔴" in str(r['KẾT LUẬN PHASE 1']) or "🚨" in str(r['KẾT LUẬN PHASE 1']):
                return r['KẾT LUẬN PHASE 1']
            tol = tol_weight if r['UOM_INV'] in ['KGM', 'MTK'] else tol_count
            if abs(r['Chênh_Lệch_KhaiBáo_HQ']) > tol: return "🔴 LỆCH SỐ LIỆU KHAI BÁO TRÊN TỜ KHAI HQ"
            return "🟢 TOÀN BỘ CHỨNG TỪ KHỚP HOÀN TOÀN (AN TOÀN THÔNG QUAN)"

        mg_final['TRẠNG THÁI CUỐI CÙNG'] = mg_final.apply(eval_p2, axis=1)
        mg_final = mg_final.sort_values(by='TRẠNG THÁI CUỐI CÙNG', ascending=False)

        st.markdown("##### 📊 BẢNG VÀNG KIỂM ĐỊNH LIÊN THÔNG TOÀN DIỆN")
        err_filter = st.toggle("🚨 Chỉ hiển thị các dòng phát hiện chênh lệch lỗi số lượng", value=False)
        board_display = mg_final[~mg_final['TRẠNG THÁI CUỐI CÙNG'].str.contains("🟢", regex=False)] if err_filter else mg_final
        
        st.data_editor(
            board_display.style.applymap(lambda x: 'background-color:#d1fae5; color:#065f46' if '🟢' in str(x) else 'background-color:#7f1d1d; color:#fca5a5; font-weight:bold' if '🚨' in str(x) else 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🔴' in str(x) or '❌' in str(x) else '', subset=['TRẠNG THÁI CUỐI CÙNG']),
            use_container_width=True, hide_index=True
        )

        tong_thue = mg_final['Thue_Du_Kien'].sum() if 'Thue_Du_Kien' in mg_final.columns else 0
        if tong_thue > 0: st.info(f"💸 **Dự toán Thuế lô hàng Nhập Khẩu:** Khoảng ${tong_thue:,.2f} USD dựa trên Biểu thuế.")

        st.markdown("#### 📥 TRÍCH XUẤT OUTPUT BÁO CÁO")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button("📊 TẢI EXCEL ĐỐI CHIẾU CHÉO (Có phân loại sheet lỗi)", ExportEngine.to_excel(mg_final), "KetQua_QuyetToan_AnToan.xlsx", "primary", use_container_width=True)
        with d2:
            doc = Document()
            doc.add_heading('BIÊN BẢN ĐỐI CHIẾU LIÊN THÔNG CHỨNG TỪ SCM EXIM VN', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"Hệ thống kiểm chéo tự động ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            err_rows = mg_final[~mg_final['TRẠNG THÁI CUỐI CÙNG'].str.contains("🟢", regex=False)]
            if err_rows.empty:
                doc.add_paragraph("Kết quả kiểm tra: Hồ sơ hoàn hảo, khớp số liệu trên mọi mặt trận.")
            else:
                table = doc.add_table(rows=1, cols=3)
                table.style = 'Table Grid'
                hdr = table.rows[0].cells
                hdr[0].text, hdr[1].text, hdr[2].text = 'Mã Vật Tư', 'Số Lượng INV', 'Chi Tiết Sai Lệch'
                for _, r_row in err_rows.head(20).iterrows():
                    row_cells = table.add_row().cells
                    row_cells[0].text, row_cells[1].text, row_cells[2].text = str(r_row['Ma_Vat_Tu']), str(r_row['SL_INV']), str(r_row['TRẠNG THÁI CUỐI CÙNG'])
            word_buf = io.BytesIO()
            doc.save(word_buf)
            st.download_button("📝 TẢI BIÊN BẢN WORD KÝ DUYỆT TRÌNH GIÁM ĐỐC", word_buf.getvalue(), "BienBan_Audit_Exim.docx", use_container_width=True)

elif f_ecus_hang and st.session_state.data_p1.empty:
    st.error("⚠️ Bạn đang đi ngược quy trình! Vui lòng hoàn thành nạp file và đối chiếu thực tế ở Phase 1 trước khi tải file Hải quan vào Phase 2.")
st.markdown("</div>", unsafe_allow_html=True)
