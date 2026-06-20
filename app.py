import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import pdfplumber
import easyocr
from PIL import Image
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import difflib
import plotly.express as px
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & BẢO MẬT (WHITE-LABEL)
# ==========================================
st.set_page_config(layout="wide", page_title="CLG SCM EXIM VN E23 E54 checker", page_icon="💎", initial_sidebar_state="expanded")

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .stDeployButton {display:none;}
            .block-container {padding-top: 1rem; padding-bottom: 2rem;}
            /* Custom Sidebar */
            [data-testid="stSidebar"] {background-color: #f8fafc;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

def check_password():
    if "auth" not in st.session_state: st.session_state["auth"] = False
    if not st.session_state["auth"]:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.markdown("<h2 style='text-align: center; color: #1E3A8A; padding-top: 50px;'>🌐 CHINGLUH GROUP SCM EXIM VN - IMPORT TEAM</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center;'>Hệ thống liên thông dữ liệu Ching Luh & Fu-Luh khi khai báo E54 và E23</p>", unsafe_allow_html=True)
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
# 2. KHỞI TẠO BỘ NHỚ TRUNG TÂM (DATABASE ẢO)
# ==========================================
if 'master_hsqd' not in st.session_state: st.session_state.master_hsqd = pd.DataFrame()
if 'master_thue' not in st.session_state: st.session_state.master_thue = pd.DataFrame()

# ==========================================
# 3. LÕI THUẬT TOÁN ĐỌC VÀ LÀM SẠCH DỮ LIỆU CAO CẤP
# ==========================================
def clean_numeric_string(series):
    """Cải tiến: Khử nhiễu toàn bộ ký tự chữ, đơn vị tính dính vào số bằng Regex"""
    if series is None: return 0
    cleaned = series.astype(str).str.replace(',', '', regex=False)
    cleaned = cleaned.str.replace(r'[^\d\.]', '', regex=True) # Chỉ giữ lại số và dấu chấm thập phân
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

def purify_code(s):
    """Cải tiến: San phẳng chuỗi ký tự, xóa bỏ dấu gạch ngang, gạch chéo để khớp tuyệt đối"""
    if pd.isna(s): return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).strip().upper())

def universal_file_extractor(uploaded_file, skiprows=0):
    if uploaded_file is None: return pd.DataFrame()
    ext = uploaded_file.name.split('.')[-1].lower()
    try:
        if ext in ['xlsx', 'xls', 'csv']:
            df = pd.read_csv(uploaded_file, skiprows=skiprows, on_bad_lines='skip') if ext == 'csv' else pd.read_excel(uploaded_file, skiprows=skiprows)
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
                                cleaned_row = [str(c).replace('\n', ' ') if c else '' for c in row]
                                if any(cleaned_row): all_text_rows.append(cleaned_row)
                    else:
                        text = page.extract_text()
                        if text:
                            for line in text.split('\n'): all_text_rows.append([line])
            if all_text_rows: 
                return pd.DataFrame(all_text_rows[1:], columns=[str(c) for c in all_text_rows[0]])
        elif ext in ['png', 'jpg', 'jpeg']:
            img = Image.open(uploaded_file).convert('RGB')
            img.thumbnail((1600, 1600), Image.Resampling.LANCZOS)
            ocr_results = reader.readtext(np.array(img), detail=0)
            return pd.DataFrame(ocr_results, columns=["Du_Lieu_Anh_OCR"])
    except Exception as e: st.error(f"Lỗi đọc tệp {uploaded_file.name}: {e}")
    return pd.DataFrame()

def norm_uom(uom):
    if pd.isna(uom) or str(uom).strip() == '': return "PCE"
    d = {'PC':'PCE', 'UNIT':'PCE', 'PCS':'PCE', 'PR':'PRS', 'PAIRS':'PRS', 'PAIR':'PRS', 'KG':'KGM', 'KGS':'KGM', 'M':'MTR', 'YD':'YRD', 'YDS':'YRD', 'M2':'MTK'}
    return d.get(str(uom).strip().upper(), str(uom).strip().upper())

def get_col(df, kws, fb_idx):
    for c in df.columns:
        for k in kws:
            if k.lower() in str(c).lower(): return c
    return df.columns[fb_idx] if len(df.columns) > fb_idx else None

def fz_match(code, masters, master_purified_dict):
    """Cải tiến: Kết hợp Fuzzy Match và Kỹ thuật san phẳng chuỗi mã vật tư"""
    p_code = purify_code(code)
    if p_code in master_purified_dict:
        return master_purified_dict[p_code]
    if code in masters: return code
    m = difflib.get_close_matches(code, masters, n=1, cutoff=0.82)
    return m[0] if m else code

def parse_ecus(raw_df, master_codes, prefix):
    if raw_df.empty: return pd.DataFrame(columns=['Ma_Vat_Tu', f'SL_{prefix}', 'HS_Code'])
    d_col = get_col(raw_df, ['Mô tả', 'Tên hàng'], 1)
    q_col = get_col(raw_df, ['Số lượng', 'Lượng tính thuế'], 3)
    h_col = get_col(raw_col_hs := get_col(raw_df, ['Mã số', 'HS'], 0), ['Mã số', 'HS'], 0)
    
    tmp = raw_df[[d_col, q_col, h_col]].dropna(subset=[d_col]).copy()
    codes = []
    for _, r in tmp.iterrows():
        desc = str(r[d_col]).upper()
        found = "UNK"
        for c in master_codes:
            if c in desc: found = c; break
        codes.append(found)
    tmp['Ma_Vat_Tu'] = codes
    res = tmp[tmp['Ma_Vat_Tu'] != 'UNK'].copy()
    res[f'SL_{prefix}'] = clean_numeric_string(res[q_col])
    res.rename(columns={h_col: 'HS_Code'}, inplace=True)
    return res.groupby(['Ma_Vat_Tu', 'HS_Code'], as_index=False)[f'SL_{prefix}'].sum()

# =====================================================================
# 4. EXPORT ENGINE CƯỜNG HÓA (XUẤT FILE CHO SẾP XEM ĐẠT CHUẨN)
# =====================================================================
class ExportEngine:
    @staticmethod
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Sheet 1: Tổng hợp
            df.to_excel(writer, index=False, sheet_name='Cross_Check_Result')
            wb = writer.book
            ws = writer.sheets['Cross_Check_Result']
            
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1, 'align': 'center'})
            fmt_red = wb.add_format({'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'border': 1})
            fmt_green = wb.add_format({'bg_color': '#D1FAE5', 'font_color': '#065F46', 'border': 1})
            fmt_num = wb.add_format({'num_format': '#,##0.00', 'border': 1})
            
            # Đắp độ rộng cột tự động theo ký tự dài nhất
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)
                max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 3
                ws.set_column(col_idx, col_idx, max_len)
                if df[col_name].dtype in [np.float64, np.int64]:
                    ws.set_column(col_idx, col_idx, max_len, fmt_num)
            
            status_col_idx = len(df.columns) - 1
            for row_idx, status in enumerate(df.iloc[:, status_col_idx]):
                fmt = fmt_green if "🟢" in str(status) or "🟡" in str(status) else fmt_red
                ws.write(row_idx + 1, status_col_idx, str(status), fmt)
                
            # Sheet 2: Chỉ tách riêng lỗi hành động nhanh (Action Required)
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
        doc.add_heading('BIÊN BẢN KIỂM ĐỊNH CHỨNG TỪ LIÊN THÔNG ĐA CHIỀU', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Thời gian lập biên bản: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        
        doc.add_heading('Danh sách các mặt hàng phát hiện sai lệch dữ liệu:', level=1)
        err_df = df[~df.iloc[:, -1].str.contains("🟢|🟡", regex=True)]
        if err_df.empty:
            doc.add_paragraph("Kết quả: Số liệu đồng nhất giữa tất cả các chứng từ Giao - Nhận.")
        else:
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Mã Vật Tư', 'Số Lượng INV', 'Trạng Thái Sai Lệch'
            for _, r in err_df.head(25).iterrows():
                row = table.add_row().cells
                row[0].text, row[1].text, row[2].text = str(r.iloc[0]), str(r.iloc[2]), str(r.iloc[-1])
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()

# ==========================================
# 5. GIAO DIỆN ĐIỀU HÀNH TÁC CHIẾN ĐƠN TRANG
# ==========================================
with st.sidebar:
    st.markdown("### ⚖️ THANH DUNG SAI LINH HOẠT")
    tol_weight = st.slider("Dung sai hàng Cân ký/Mét (KGM, MTK):", min_value=0.0, max_value=1.0, value=0.2, step=0.05)
    tol_count = st.slider("Dung sai hàng Đếm chiếc (PCE, PRS):", min_value=0.0, max_value=0.1, value=0.01, step=0.01)
    st.markdown("---")
    st.markdown("### ⚙️ CẤU HÌNH DÒNG THỪA (SKIPROWS)")
    skip_inv = st.number_input("Dòng thừa Invoice/PKL:", value=15)
    skip_cd  = st.number_input("Dòng thừa Chỉ Định:", value=17)
    skip_erp = st.number_input("Dòng thừa SAP (Xuất/Nhập):", value=0)
    skip_ecus = st.number_input("Dòng thừa Tờ Khai (E54/E23):", value=18)

st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>💎 TRUNG TÂM KIỂM ĐỊNH LIÊN THÔNG CHỨNG TỪ XNK</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.1em;'>Hỗ trợ đọc vạn năng Excel, PDF, Ảnh Scan. Tự động cross-check luồng Giao - Nhận 2 chiều tại 1 màn hình duy nhất.</p>", unsafe_allow_html=True)
st.markdown("---")

# Module Master Data thu gọn dưới dạng Expander ẩn
with st.expander("⚙️ NẠP DỮ LIỆU NỀN QUY ĐỔI & BIỂU THUẾ TRUNG TÂM"):
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        f_hsqd = st.file_uploader("Bảng Hệ số quy đổi (HSQĐ):", type=["xlsx","csv"])
        if f_hsqd:
            df = universal_file_extractor(f_hsqd, 1)
            if not df.empty:
                df = df[[get_col(df, ['Mã'], 1), get_col(df, ['Hệ số'], 6), get_col(df, ['Đơn vị'], 8)]].dropna()
                df.columns = ['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan']
                df['Ma_Vat_Tu'] = df['Ma_Vat_Tu'].apply(purify_code)
                df['He_So_QD'] = pd.to_numeric(df['He_So_QD'], errors='coerce').fillna(1.0)
                st.session_state.master_hsqd = df
                st.success("✅ Đã đồng bộ cấu hình hệ số quy đổi.")
    with m_col2:
        f_thue = st.file_uploader("Bảng Biểu thuế XNK (Tra cứu mã HS):", type=["xlsx","csv"])
        if f_thue:
            df = universal_file_extractor(f_thue, 2)
            if not df.empty:
                df = df[[get_col(df, ['Mã hàng', 'HS'], 1), get_col(df, ['Thuế suất'], 5)]].dropna()
                df.columns = ['HS_Code', 'Thue_Suat']
                df['HS_Code'] = df['HS_Code'].astype(str).str.replace('.', '', regex=False).str.strip()
                st.session_state.master_thue = df
                st.success("✅ Đã cập nhật biểu thuế.")

# Ô kéo thả file nghiệp vụ
st.markdown("### 📥 TẢI LÊN TOÀN BỘ CHỨNG TỪ CỦA LÔ HÀNG")
col1, col2, col3 = st.columns(3)
with col1:
    st.info("📄 **CHỨNG TỪ THƯƠNG MẠI GỐC**")
    f_inv = st.file_uploader("1. INVOICE KHAI BÁO (Excel/PDF/Ảnh)", type=["xlsx","csv","pdf","jpg","png"])
    f_pkl = st.file_uploader("2. PACKING LIST (Excel/PDF/Ảnh)", type=["xlsx","csv","pdf","jpg","png"])
    f_cd  = st.file_uploader("3. CHỈ ĐỊNH GIAO HÀNG (Vận chuyển)", type=["xlsx","csv","pdf"])
with col2:
    st.warning("💻 **SỔ SÁCH KHO NỘI BỘ (SAP/ERP)**")
    f_sap_out = st.file_uploader("4. SAP XUẤT KHO (Bên Giao hàng)", type=["xlsx","csv"])
    f_sap_in  = st.file_uploader("5. SAP NHẬP KHO (Bên Nhận hàng)", type=["xlsx","csv"])
with col3:
    st.success("🏛️ **HẢI QUAN ĐỐI ỨNG (ECUS)**")
    f_e54 = st.file_uploader("6. TỜ KHAI XUẤT (E54)", type=["xlsx","csv"])
    f_e23 = st.file_uploader("7. TỜ KHAI NHẬP (E23)", type=["xlsx","csv"])

# ENGINE ĐỐI CHIẾU CHÉO TẬP TRUNG
if st.button("🚀 KÍCH HOẠT SIÊU HỆ THỐNG ĐỐI CHIẾU TOÀN DIỆN", type="primary", use_container_width=True):
    if not (f_inv and f_pkl):
        st.error("⚠️ Điểm chốt chặn yêu cầu tối thiểu file Invoice và Packing List để kích hoạt máy.")
    else:
        with st.spinner("Động cơ AI và Thuật toán liên thông đang quét chứng từ..."):
            r_inv = universal_file_extractor(f_inv, skip_inv)
            r_pkl = universal_file_extractor(f_pkl, skip_inv)
            r_cd  = universal_file_extractor(f_cd, skip_cd)
            r_sap_out = universal_file_extractor(f_sap_out, skip_erp)
            r_sap_in  = universal_file_extractor(f_sap_in, skip_erp)
            r_e54 = universal_file_extractor(f_e54, skip_ecus)
            r_e23 = universal_file_extractor(f_e23, skip_ecus)

            # 1. Invoice (Trục xương sống)
            c_inv_mat = get_col(r_inv, ['Material', 'Mã'], 1)
            df_inv = r_inv[[c_inv_mat, get_col(r_inv, ['Quantity', 'Số lượng'], 6), get_col(r_inv, ['Unit', 'ĐVT'], 5), get_col(r_inv, ['Price', 'Đơn giá'], 7), get_col(r_inv, ['Amount', 'Thành tiền'], 8)]].dropna(subset=[c_inv_mat]).copy()
            df_inv.columns = ['Ma_Vat_Tu', 'SL_INV', 'UOM_INV', 'DonGia', 'TriGia']
            df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].astype(str).str.strip().str.upper()
            for c in ['SL_INV', 'DonGia', 'TriGia']: df_inv[c] = clean_numeric_string(df_inv[c])
            
            # Áp dụng HSQĐ
            if not st.session_state.master_hsqd.empty:
                df_inv['Ma_Purified'] = df_inv['Ma_Vat_Tu'].apply(purify_code)
                df_inv = pd.merge(df_inv, st.session_state.master_hsqd, left_on='Ma_Purified', right_on='Ma_Vat_Tu', how='left', suffixes=('', '_hsqd'))
                df_inv['He_So_QD'] = df_inv['He_So_QD'].fillna(1.0)
                df_inv['SL_INV'] = df_inv['SL_INV'] * df_inv['He_So_QD']
                df_inv['UOM_INV'] = df_inv['DVT_Bao_Quan'].combine_first(df_inv['UOM_INV'])
                df_inv.drop(columns=['He_So_QD', 'DVT_Bao_Quan', 'Ma_Purified', 'Ma_Vat_Tu_hsqd'], errors='ignore', inplace=True)
            df_inv['UOM_INV'] = df_inv['UOM_INV'].apply(norm_uom)
            
            masters = df_inv['Ma_Vat_Tu'].tolist()
            master_purified_dict = {purify_code(m): m for m in masters} # Từ điển phục vụ ép mã nhanh

            # 2. Packing List
            p_mat = get_col(r_pkl, ['Material', 'Mã'], 0)
            p_qty = get_col(r_pkl, ["Q'TY", 'Quantity'], 5)
            p_uom = get_col(r_pkl, ['Unit', 'ĐVT'], 2)
            df_pkl_raw = pd.DataFrame()
            df_pkl_raw['Ma_Vat_Tu'] = r_pkl[p_mat].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
            df_pkl_raw['SL_PKL'] = clean_numeric_string(r_pkl[p_qty])
            df_pkl_raw['UOM_PKL'] = r_pkl[p_uom].apply(norm_uom)
            df_pkl = apply_hsqd(df_pkl_raw, 'SL_PKL', 'UOM_PKL')
            df_pkl = df_pkl.groupby('Ma_Vat_Tu', as_index=False)['SL_PKL'].sum()

            # 3. Chỉ định giao hàng
            df_cd = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ChiDinh'])
            if not r_cd.empty:
                cd_mat = get_col(r_cd, ['Mã NL', 'Material'], 1)
                cd_qty = get_col(r_cd, ['Số Lượng', 'Quantity'], 4)
                df_cd = pd.DataFrame()
                df_cd['Ma_Vat_Tu'] = r_cd[cd_mat].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
                df_cd['SL_ChiDinh'] = clean_numeric_string(r_cd[cd_qty])
                df_cd = df_cd.groupby('Ma_Vat_Tu', as_index=False)['SL_ChiDinh'].sum()

            # 4. SAP Xuất (Bên Giao)
            df_sap_out = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_SAP_XUAT'])
            if not r_sap_out.empty:
                df_sap_out = r_sap_out[[r_sap_out.columns[0], r_sap_out.columns[1]]].dropna().copy()
                df_sap_out.columns = ['Ma_Vat_Tu', 'SL_SAP_XUAT']
                df_sap_out['Ma_Vat_Tu'] = df_sap_out['Ma_Vat_Tu'].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
                df_sap_out['SL_SAP_XUAT'] = clean_numeric_string(df_sap_out['SL_SAP_XUAT'])
                df_sap_out = df_sap_out.groupby('Ma_Vat_Tu', as_index=False)['SL_SAP_XUAT'].sum()

            # 5. SAP Nhập (Bên Nhận)
            df_sap_in = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_SAP_NHAP'])
            if not r_sap_in.empty:
                df_sap_in = r_sap_in[[r_sap_in.columns[0], r_sap_in.columns[1]]].dropna().copy()
                df_sap_in.columns = ['Ma_Vat_Tu', 'SL_SAP_NHAP']
                df_sap_in['Ma_Vat_Tu'] = df_sap_in['Ma_Vat_Tu'].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
                df_sap_in['SL_SAP_NHAP'] = clean_numeric_string(df_sap_in['SL_SAP_NHAP'])
                df_sap_in = df_sap_in.groupby('Ma_Vat_Tu', as_index=False)['SL_SAP_NHAP'].sum()

            # 6. Tờ khai Hải quan đối ứng E54 và E23
            df_e54 = parse_ecus(r_e54, masters, "E54")
            df_e23 = parse_ecus(r_e23, masters, "E23")

            # FULL OUTER JOIN GỘP BẢNG TRUNG TÂM
            mg = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_cd, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_sap_out, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_sap_in, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_e54, on='Ma_Vat_Tu', how='outer')
            if not df_e23.empty:
                mg = pd.merge(mg, df_e23, on='Ma_Vat_Tu', how='outer', suffixes=('', '_drop'))
                if 'HS_Code_drop' in mg.columns:
                    mg['HS_Code'] = mg['HS_Code'].combine_first(mg['HS_Code_drop'])
                    mg.drop(columns=['HS_Code_drop'], inplace=True)
            mg = mg.fillna(0)

            # Thuế Dự Kiến
            mg['Thue_Du_Kien'] = 0.0
            if not st.session_state.master_thue.empty and 'HS_Code' in mg.columns:
                t_dict = dict(zip(st.session_state.master_thue['HS_Code'], st.session_state.master_thue['Thue_Suat']))
                mg['Thue_NK_Suat'] = mg['HS_Code'].astype(str).str.replace('.','', regex=False).map(t_dict).fillna(0)
                mg['Thue_Du_Kien'] = mg['TriGia'] * pd.to_numeric(mg['Thue_NK_Suat'], errors='coerce').fillna(0) / 100

            # LỘI SỬA BÀI TOÁN DUNG SAI ĐỘNG TỪ SIDEBAR
            def evaluate_all(r):
                if r['SL_INV'] == 0: return "❌ LỖI: THIẾU TRÊN INVOICE GỐC"
                tol = tol_weight if r['UOM_INV'] in ['KGM', 'MTK'] else tol_count
                
                if 'SL_E54' in r and 'SL_E23' in r and r['SL_E54'] > 0 and r['SL_E23'] > 0 and abs(r['SL_E54'] - r['SL_E23']) > tol:
                    return "🚨 FATAL: LỆCH GIỮA TỜ KHAI XUẤT VÀ TỜ KHAI NHẬP!"
                if abs(r['SL_INV'] - r['SL_PKL']) > tol: return "🔴 LỆCH SỐ LIỆU ĐÓNG GÓI PACKING LIST"
                if 'SL_ChiDinh' in r and r['SL_ChiDinh'] > 0 and abs(r['SL_INV'] - r['SL_ChiDinh']) > tol: return "🔴 LỆCH CHỈ ĐỊNH GIAO HÀNG"
                if 'SL_SAP_XUAT' in r and r['SL_SAP_XUAT'] > 0 and abs(r['SL_INV'] - r['SL_SAP_XUAT']) > tol: return "🟠 LỆCH KHO XUẤT (BÊN GIAO)"
                if 'SL_SAP_NHAP' in r and r['SL_SAP_NHAP'] > 0 and abs(r['SL_INV'] - r['SL_SAP_NHAP']) > tol: return "🚨 LỖI THẤT THOÁT: KHO NHẬP BỊ THIẾU"
                if 'SL_E54' in r and r['SL_E54'] > 0 and abs(r['SL_INV'] - r['SL_E54']) > tol: return "🔴 LỆCH KHAI BÁO TỜ KHAI XUẤT E54"
                if 'SL_E23' in r and r['SL_E23'] > 0 and abs(r['SL_INV'] - r['SL_E23']) > tol: return "🔴 LỆCH KHAI BÁO TỜ KHAI NHẬP E23"
                return "🟢 CHỨNG TỪ KHỚP HOÀN TOÀN"

            mg['KẾT LUẬN KIỂM CHÉO'] = mg.apply(evaluate_all, axis=1)
            mg = mg.sort_values(by='KẾT LUẬN KIỂM CHÉO', ascending=False)

            # RENDER GIAO DIỆN KẾT QUẢ ĐÃ ĐƯỢC LỌC NHANH
            st.markdown("---")
            st.markdown("### 📊 BẢNG KẾT QUẢ ĐỐI CHIẾU LIÊN THÔNG TOÀN CỤC")
            err_filter = st.toggle("🚨 BẬT BỘ LỌC LỖI (Ẩn dòng khớp hoàn toàn)", value=False)
            final_board = mg[~mg['KẾT LUẬN KIỂM CHÉO'].str.contains("🟢", regex=False)] if err_filter else mg
            
            st.data_editor(
                final_board.style.applymap(lambda x: 'background-color:#d1fae5; color:#065f46' if '🟢' in str(x) else 'background-color:#7f1d1d; color:#fca5a5; font-weight:bold' if '🚨' in str(x) else 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🔴' in str(x) or '❌' in str(x) else '', subset=['KẾT LUẬN KIỂM CHÉO']),
                use_container_width=True, hide_index=True, height=450
            )

            # NÚT TẢI FILE TẬP TRUNG NGAY DƯỚI BẢNG KẾT QUẢ
            st.markdown("#### 📥 TRÍCH XUẤT KẾT QUẢ ĐA ĐỊNH DẠNG TỨC THÌ")
            d1, d2 = st.columns(2)
            with d1:
                st.download_button("📊 XUẤT FILE EXCEL ĐỐI CHIẾU CHÉO (Có Summary & Tô màu)", ExportEngine.to_excel(mg), "KetQua_CrossCheck_Master.xlsx", "primary", use_container_width=True)
            with d2:
                st.download_button("📝 XUẤT BIÊN BẢN WORD KÝ DUYỆT TRÌNH SẾP", ExportEngine.to_word(mg), "BienBan_DoiChieu_SCM.docx", use_container_width=True)
