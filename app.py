import streamlit as st
import pandas as pd
import io
import pdfplumber
import easyocr
import numpy as np
from PIL import Image
from fpdf import FPDF
import difflib
import plotly.express as px

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & BẢO MẬT
# ==========================================
st.set_page_config(layout="wide", page_title="CLG SCM EXIM VN E23 E54 checker", page_icon="💎", initial_sidebar_state="expanded")

def check_password():
    if "password_correct" not in st.session_state: st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>🔒 CỔNG AN NINH PHÒNG XNK</h2>", unsafe_allow_html=True)
        pwd = st.text_input("Nhập mật khẩu truy cập hệ thống:", type="password")
        if pwd:
            if pwd == st.secrets.get("app_password", "ChingLuh@2026"):
                st.session_state["password_correct"] = True
                st.rerun()
            else: st.error("❌ Mật khẩu không chính xác!")
        return False
    return True

if not check_password(): st.stop()

@st.cache_resource
def load_ocr_reader(): return easyocr.Reader(['vi', 'en'], gpu=False)
reader = load_ocr_reader()

# ==========================================
# 2. HÀM BỔ TRỢ NGHIỆP VỤ XNK & FUZZY MATCH
# ==========================================
def extract_and_clean_data(uploaded_file, skiprows_count=0):
    if uploaded_file is None: return None
    ext = uploaded_file.name.split('.')[-1].lower()
    try:
        if ext in ['xlsx', 'xls', 'csv']:
            df = pd.read_csv(uploaded_file, skiprows=skiprows_count, on_bad_lines='skip') if ext == 'csv' else pd.read_excel(uploaded_file, skiprows=skiprows_count)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df = df.dropna(how='all')
            df.columns = [str(c).strip() for c in df.columns]
            return df
        elif ext == 'pdf':
            all_rows = []
            with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                for page in pdf.pages:
                    for table in page.extract_tables():
                        for row in table:
                            cr = [str(cell).replace('\n', ' ') if cell else '' for cell in row]
                            if any(cr): all_rows.append(cr)
            if all_rows: return pd.DataFrame(all_rows[1:], columns=all_rows[0])
        elif ext in ['png', 'jpg', 'jpeg']:
            img = Image.open(uploaded_file)
            img.thumbnail((1600, 1600), Image.Resampling.LANCZOS)
            return pd.DataFrame(reader.readtext(np.array(img), detail=0), columns=["Dữ_Liệu"])
    except Exception as e: st.error(f"Lỗi file {uploaded_file.name}: {e}")
    return None

def normalize_uom(uom):
    if pd.isna(uom) or str(uom).strip() == '': return "PCE"
    uom_dict = {'PC': 'PCE', 'UNIT': 'PCE', 'PCS': 'PCE', 'PR': 'PRS', 'PAIRS': 'PRS', 'PAIR': 'PRS', 'KG': 'KGM', 'KGS': 'KGM', 'M': 'MTR', 'YD': 'YRD', 'YDS': 'YRD', 'M2': 'MTK', 'MTK': 'MTK'}
    return uom_dict.get(str(uom).strip().upper(), str(uom).strip().upper())

def get_col_fallback(df, keywords, fallback_idx):
    for col in df.columns:
        for kw in keywords:
            if kw.lower() in col.lower(): return col
    return df.columns[fallback_idx] if len(df.columns) > fallback_idx else None

def fuzzy_match_code(target_code, master_codes_list):
    """Thuật toán Fuzzy Matching: Tự động gom mã nếu gõ sai chính tả nhẹ (Sai dấu cách, gạch nối)"""
    if target_code in master_codes_list: return target_code, False
    matches = difflib.get_close_matches(target_code, master_codes_list, n=1, cutoff=0.85)
    if matches: return matches[0], True # Trả về mã chuẩn đã match và Cờ báo hiệu (Fuzzy = True)
    return target_code, False

# ==========================================
# 3. GIAO DIỆN UPLOAD FILE (7 MODULES)
# ==========================================
st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>💎 XNK ULTIMATE AI CHECKER (LEVEL MAX)</h1>", unsafe_allow_html=True)
st.markdown("---")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.info("📦 **CHỨNG TỪ GỐC**")
    file_inv = st.file_uploader("1. INVOICE KHAI BÁO", type=["xlsx", "csv", "pdf"])
    file_pkl = st.file_uploader("2. PACKING LIST", type=["xlsx", "csv", "pdf"])
with c2:
    st.warning("🚛 **SỔ SÁCH & KHO**")
    file_cd  = st.file_uploader("3. CHỈ ĐỊNH GIAO HÀNG", type=["xlsx", "csv", "pdf"])
    file_erp = st.file_uploader("4. BÁO CÁO XUẤT KHO SAP", type=["xlsx", "csv"])
with c3:
    st.success("🏛️ **HẢI QUAN ECUS**")
    file_ecus = st.file_uploader("5. TỜ KHAI HẢI QUAN HÀNG", type=["xlsx", "csv"])
    file_thue = st.file_uploader("6. BẢNG BIỂU THUẾ (Mã HS)", type=["xlsx", "csv"])
with c4:
    st.error("⚙️ **MASTER DATA**")
    file_hsqd = st.file_uploader("7. BẢNG HỆ SỐ QUY ĐỔI", type=["xlsx", "csv"])

with st.sidebar:
    st.markdown("### ⚙️ CẤU HÌNH HEADER (SKIPROWS)")
    skip_inv = st.number_input("Header Invoice:", value=15)
    skip_pkl = st.number_input("Header PKL:", value=15)
    skip_cd  = st.number_input("Header Chỉ Định:", value=17)
    skip_erp = st.number_input("Header SAP/ERP:", value=0)
    skip_ecus = st.number_input("Header ECUS:", value=18)
    skip_thue = st.number_input("Header Biểu Thuế:", value=2)
    skip_hsqd = st.number_input("Header Bảng HSQĐ:", value=1)

# ==========================================
# 4. CORE ENGINE XỬ LÝ
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 KÍCH HOẠT SIÊU HỆ THỐNG KIỂM ĐỊNH", type="primary", use_container_width=True):
    if not (file_inv and file_pkl):
        st.error("⚠️ Hệ thống yêu cầu tải lên tối thiểu INVOICE và PACKING LIST.")
    else:
        with st.spinner("🔄 Hệ thống đang chạy Fuzzy Match, tính N.W/G.W, và dò mã HS Biểu Thuế..."):
            
            raw_inv = extract_and_clean_data(file_inv, skip_inv)
            raw_pkl = extract_and_clean_data(file_pkl, skip_pkl)
            raw_cd  = extract_and_clean_data(file_cd, skip_cd)
            raw_erp = extract_and_clean_data(file_erp, skip_erp) if file_erp else pd.DataFrame()
            raw_ecus = extract_and_clean_data(file_ecus, skip_ecus) if file_ecus else pd.DataFrame()
            raw_thue = extract_and_clean_data(file_thue, skip_thue) if file_thue else pd.DataFrame()
            raw_hsqd = extract_and_clean_data(file_hsqd, skip_hsqd) if file_hsqd else pd.DataFrame()

            def clean_code(s): return str(s).strip().upper()

            # --- SETUP HSQĐ ---
            df_hsqd = pd.DataFrame(columns=['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan'])
            if not raw_hsqd.empty:
                df_hsqd = raw_hsqd[[get_col_fallback(raw_hsqd, ['Mã số', 'Material'], 1), get_col_fallback(raw_hsqd, ['Hệ số', 'Rate'], 6), get_col_fallback(raw_hsqd, ['Đơn vị báo quan'], 8)]].dropna()
                df_hsqd.columns = ['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan']
                df_hsqd['Ma_Vat_Tu'] = df_hsqd['Ma_Vat_Tu'].apply(clean_code)
                df_hsqd['He_So_QD'] = pd.to_numeric(df_hsqd['He_So_QD'], errors='coerce').fillna(1.0)
                df_hsqd['DVT_Bao_Quan'] = df_hsqd['DVT_Bao_Quan'].apply(normalize_uom)

            def apply_hsqd_engine(df, qty_col, uom_col):
                if not df_hsqd.empty and not df.empty:
                    df = pd.merge(df, df_hsqd, on='Ma_Vat_Tu', how='left')
                    df['He_So_QD'] = df['He_So_QD'].fillna(1.0)
                    df[qty_col] = df[qty_col] * df['He_So_QD']
                    if uom_col in df.columns: df[uom_col] = df['DVT_Bao_Quan'].combine_first(df[uom_col])
                    return df.drop(columns=['He_So_QD', 'DVT_Bao_Quan'], errors='ignore')
                return df

            # --- 1. INVOICE ---
            i_mat = get_col_fallback(raw_inv, ['Material code', 'Mã'], 1)
            df_inv = raw_inv[[i_mat, get_col_fallback(raw_inv, ['Quantity', 'Số lượng'], 6), get_col_fallback(raw_inv, ['Unit', 'ĐVT'], 5), get_col_fallback(raw_inv, ['Unit Price'], 7), get_col_fallback(raw_inv, ['Amount'], 8)]].dropna(subset=[i_mat]).copy()
            df_inv.columns = ['Ma_Vat_Tu', 'SL_Invoice', 'UOM_Invoice', 'Price_Inv', 'Amount_Inv']
            df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].apply(clean_code)
            for c in ['SL_Invoice', 'Price_Inv', 'Amount_Inv']: df_inv[c] = pd.to_numeric(df_inv[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_inv = apply_hsqd_engine(df_inv, 'SL_Invoice', 'UOM_Invoice')
            df_inv['UOM_Invoice'] = df_inv['UOM_Invoice'].apply(normalize_uom)
            master_codes = df_inv['Ma_Vat_Tu'].tolist() # Tập mã chuẩn để Fuzzy Match

            # --- 2. PACKING LIST (Tính N.W / G.W) ---
            p_mat = get_col_fallback(raw_pkl, ['Material', 'Mã'], 0)
            p_qty = get_col_fallback(raw_pkl, ["Q'TY", 'Quantity'], 5)
            p_uom = get_col_fallback(raw_pkl, ['Unit', 'ĐVT'], 2)
            nw_col = get_col_fallback(raw_pkl, ['NW', 'Net Weight', 'N.W'], -1)
            gw_col = get_col_fallback(raw_pkl, ['GW', 'Gross Weight', 'G.W'], -1)
            
            df_pkl_raw = raw_pkl[[p_mat, p_qty, p_uom]].copy()
            if nw_col: df_pkl_raw['NW'] = pd.to_numeric(raw_pkl[nw_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            if gw_col: df_pkl_raw['GW'] = pd.to_numeric(raw_pkl[gw_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            df_pkl_raw = df_pkl_raw.dropna(subset=[p_mat])
            df_pkl_raw.columns = ['Ma_Vat_Tu', 'SL_PKL', 'UOM_PKL'] + ([c for c in ['NW', 'GW'] if c in df_pkl_raw.columns])
            df_pkl_raw['Ma_Vat_Tu'] = df_pkl_raw['Ma_Vat_Tu'].apply(clean_code)
            df_pkl_raw['SL_PKL'] = pd.to_numeric(df_pkl_raw['SL_PKL'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            # Fuzzy Match cho PKL
            df_pkl_raw['Ma_Vat_Tu'] = df_pkl_raw['Ma_Vat_Tu'].apply(lambda x: fuzzy_match_code(x, master_codes)[0])
            
            df_pkl = apply_hsqd_engine(df_pkl_raw, 'SL_PKL', 'UOM_PKL')
            df_pkl['UOM_PKL'] = df_pkl['UOM_PKL'].apply(normalize_uom)
            df_pkl = df_pkl.groupby(['Ma_Vat_Tu', 'UOM_PKL'], as_index=False)['SL_PKL'].sum()

            # --- 3. ERP (Fuzzy Match) ---
            df_erp = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ERP'])
            if not raw_erp.empty:
                df_erp = raw_erp[[raw_erp.columns[0], raw_erp.columns[1]]].dropna().copy()
                df_erp.columns = ['Ma_Vat_Tu', 'SL_ERP']
                df_erp['Ma_Vat_Tu'] = df_erp['Ma_Vat_Tu'].apply(clean_code)
                # Fuzzy Match
                df_erp['Fuzzy_Flag'] = df_erp['Ma_Vat_Tu'].apply(lambda x: fuzzy_match_code(x, master_codes)[1])
                df_erp['Ma_Vat_Tu'] = df_erp['Ma_Vat_Tu'].apply(lambda x: fuzzy_match_code(x, master_codes)[0])
                df_erp['SL_ERP'] = pd.to_numeric(df_erp['SL_ERP'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_erp['UOM_ERP'] = None
                df_erp = apply_hsqd_engine(df_erp, 'SL_ERP', 'UOM_ERP').drop(columns=['UOM_ERP'])

            # --- 4. ECUS & BIỂU THUẾ ---
            df_ecus = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ECUS', 'HS_Code'])
            if not raw_ecus.empty:
                ecus_mat = get_col_fallback(raw_ecus, ['Mô tả', 'Tên hàng'], 1)
                ecus_qty = get_col_fallback(raw_ecus, ['Số lượng (1)', 'Lượng tính thuế'], 3)
                ecus_hs = get_col_fallback(raw_ecus, ['Mã số hàng hóa'], 0) # Lấy cột mã HS
                
                raw_e = raw_ecus[[ecus_mat, ecus_qty, ecus_hs]].dropna(subset=[ecus_mat]).copy()
                extracted_codes = []
                for _, r in raw_e.iterrows():
                    desc = str(r[ecus_mat]).upper()
                    found = "UNKNOWN"
                    for code in master_codes:
                        if code in desc: found = code; break
                    extracted_codes.append(found)
                
                raw_e['Ma_Vat_Tu'] = extracted_codes
                df_ecus = raw_e[raw_e['Ma_Vat_Tu'] != "UNKNOWN"].copy()
                df_ecus['SL_ECUS'] = pd.to_numeric(df_ecus[ecus_qty].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_ecus.rename(columns={ecus_hs: 'HS_Code'}, inplace=True)
                df_ecus = df_ecus.groupby(['Ma_Vat_Tu', 'HS_Code'], as_index=False)['SL_ECUS'].sum()

            # Merge All
            merged = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
            merged = pd.merge(merged, df_erp, on='Ma_Vat_Tu', how='outer')
            merged = pd.merge(merged, df_ecus, on='Ma_Vat_Tu', how='outer').fillna(0)

            # --- KIỂM TRA MÃ HS VỚI BIỂU THUẾ ---
            merged['Lỗi_Thuế'] = ""
            if not raw_thue.empty and not df_ecus.empty:
                thue_hs_col = get_col_fallback(raw_thue, ['Mã hàng', 'HS'], 1)
                thue_nk_col = get_col_fallback(raw_thue, ['NK', 'Thuế'], 5)
                df_thue = raw_thue[[thue_hs_col, thue_nk_col]].dropna().copy()
                df_thue[thue_hs_col] = df_thue[thue_hs_col].astype(str).str.replace('.', '').str.strip()
                hs_valid_list = df_thue[thue_hs_col].tolist()
                
                def check_hs(hs):
                    if hs != 0 and str(hs).strip() not in hs_valid_list:
                        return "🔴 MÃ HS TRÊN TỜ KHAI KHÔNG TỒN TẠI TRONG BIỂU THUẾ"
                    return ""
                merged['Lỗi_Thuế'] = merged['HS_Code'].apply(check_hs)

            # --- ĐÁNH GIÁ TRẠNG THÁI ---
            def evaluate_row(r):
                if r['SL_Invoice'] == 0: return "❌ KHUYẾT MÃ INVOICE"
                if r.get('Lỗi_Thuế', "") != "": return r['Lỗi_Thuế']
                if r.get('Fuzzy_Flag', False): return "🟡 CẢNH BÁO TÊN MÃ (AI TỰ GỘP)"
                
                if r['UOM_Invoice'] != 0 and r['UOM_PKL'] != 0 and r['UOM_Invoice'] != r['UOM_PKL']:
                    return "🔴 LỖI ĐVT"
                if r['Price_Inv'] > 0 and r['Amount_Inv'] > 0 and abs((r['SL_Invoice'] * r['Price_Inv']) - r['Amount_Inv']) > 0.1:
                    return "🔴 LỖI NHÂN TRỊ GIÁ"

                s_pkl = abs(r['SL_Invoice'] - r['SL_PKL'])
                s_erp = abs(r['SL_Invoice'] - r['SL_ERP']) if 'SL_ERP' in r else 0
                s_hq = abs(r['SL_Invoice'] - r['SL_ECUS']) if 'SL_ECUS' in r else 0
                
                tol = 0.2 if r['UOM_Invoice'] in ['KGM', 'MTK'] else 0.01
                if s_pkl <= tol and s_erp <= tol and s_hq <= tol: return "🟢 KHỚP HOÀN TOÀN"
                if s_pkl <= tol and s_erp <= tol and s_hq <= tol: return "🟡 KHỚP (DUNG SAI)"

                if s_pkl > tol: return "🔴 LỆCH PACKING LIST"
                if not raw_ecus.empty and s_hq > tol: return "🔴 LỆCH TỜ KHAI ECUS"
                return "🟠 LỆCH SỔ SÁCH ERP"

            merged['ĐÁNH GIÁ (STATUS)'] = merged.apply(evaluate_row, axis=1)
            
            # Tổ chức lại bảng để hiển thị
            display_cols = ['Ma_Vat_Tu', 'UOM_Invoice', 'SL_Invoice', 'SL_PKL']
            if not raw_erp.empty: display_cols.append('SL_ERP')
            if not raw_ecus.empty: display_cols.extend(['HS_Code', 'SL_ECUS'])
            display_cols.append('ĐÁNH GIÁ (STATUS)')
            
            merged_display = merged[display_cols].sort_values(by='ĐÁNH GIÁ (STATUS)', ascending=False)

            # ==========================================
            # 5. RENDER CÁC TAB KẾT QUẢ (UI/UX)
            # ==========================================
            tab1, tab2, tab3 = st.tabs(["📊 BẢNG ĐỐI CHIẾU (LIVE EDITOR)", "⚖️ TRỌNG LƯỢNG (N.W/G.W)", "📈 DASHBOARD THỐNG KÊ LỖI"])

            with tab1:
                st.markdown("### 📝 CHỈNH SỬA TRỰC TIẾP (LIVE DATA EDITOR)")
                st.caption("Double-click vào các ô bị lệch để sửa lại số liệu trực tiếp trên Web. Hệ thống sẽ áp dụng dữ liệu bạn sửa để tải file.")
                
                def color_matrix(val):
                    if "🟢 KHỚP" in str(val) or "🟡 KHỚP" in str(val): return 'background-color: #d1fae5; color: #065f46'
                    if "🔴" in str(val): return 'background-color: #fee2e2; color: #991b1b; font-weight: bold'
                    if "❌" in str(val): return 'background-color: #fce7f3; color: #9d174d; font-weight: bold'
                    return 'background-color: #fef3c7; color: #b45309; font-weight: bold'

                # SỬ DỤNG DATA EDITOR THAY VÌ DATAFRAME TĨNH
                edited_df = st.data_editor(merged_display.style.map(color_matrix, subset=['ĐÁNH GIÁ (STATUS)']), use_container_width=True, height=500, num_rows="dynamic")

                st.markdown("<br>", unsafe_allow_html=True)
                ex1, ex2 = st.columns(2)
                with ex1:
                    excel_buf = io.BytesIO()
                    with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
                        edited_df.to_excel(writer, index=False, sheet_name='E54_Checked')
                    st.download_button("📥 TẢI FILE EXCEL SAU KHI SỬA", data=excel_buf.getvalue(), file_name="E54_Output_Edited.xlsx", type="primary", use_container_width=True)
                with ex2:
                    st.download_button("📑 TẢI BIÊN BẢN BÁO CÁO (PDF)", data=b"Demo PDF", file_name="E54_Report.pdf", use_container_width=True)

            with tab2:
                st.markdown("### ⚖️ KIỂM TRA TRỌNG LƯỢNG TỪ PACKING LIST")
                if 'NW' in df_pkl_raw.columns or 'GW' in df_pkl_raw.columns:
                    w1, w2 = st.columns(2)
                    total_nw = df_pkl_raw['NW'].sum() if 'NW' in df_pkl_raw.columns else 0
                    total_gw = df_pkl_raw['GW'].sum() if 'GW' in df_pkl_raw.columns else 0
                    with w1: st.metric("Tổng Trọng Lượng Tịnh (Net Weight - N.W)", f"{total_nw:,.2f} KGM")
                    with w2: st.metric("Tổng Trọng Lượng Cộp (Gross Weight - G.W)", f"{total_gw:,.2f} KGM")
                    st.info("💡 So sánh hai con số này với Footer (Tổng cộng) trên file Packing List giấy để đảm bảo nhân viên không kéo Excel sót dòng.")
                else:
                    st.warning("Packing List tải lên không có cột NW (Net Weight) hoặc GW (Gross Weight).")

            with tab3:
                st.markdown("### 📈 THỐNG KÊ TỶ LỆ LỖI CHỨNG TỪ")
                status_counts = merged_display['ĐÁNH GIÁ (STATUS)'].value_counts().reset_index()
                status_counts.columns = ['Trạng Thái', 'Số Lượng Vật Tư']
                
                fig = px.pie(status_counts, values='Số Lượng Vật Tư', names='Trạng Thái', title='Phân bổ Tình trạng Đối chiếu',
                             color='Trạng Thái', color_discrete_map={
                                 "🟢 KHỚP HOÀN TOÀN": "green", "🟡 KHỚP (DUNG SAI)": "lightgreen",
                                 "🔴 LỆCH PACKING LIST": "red", "🟠 LỆCH SỔ SÁCH ERP": "orange", "❌ KHUYẾT MÃ INVOICE": "purple"
                             })
                st.plotly_chart(fig, use_container_width=True)
