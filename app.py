import streamlit as st
import pandas as pd
import io
import pdfplumber
import easyocr
import numpy as np
from PIL import Image
from fpdf import FPDF

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & BẢO MẬT
# ==========================================
st.set_page_config(layout="wide", page_title="XNK Master Checker", page_icon="🏢", initial_sidebar_state="expanded")

def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>🔒 CỔNG AN NINH PHÒNG XNK</h2>", unsafe_allow_html=True)
        pwd = st.text_input("Nhập mật khẩu truy cập hệ thống:", type="password", key="pwd_input")
        if pwd:
            if pwd == st.secrets.get("app_password", "ChingLuh@2026"):
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ Mật khẩu không chính xác!")
        return False
    return True

if not check_password():
    st.stop()

@st.cache_resource
def load_ocr_reader():
    return easyocr.Reader(['vi', 'en'], gpu=False)

reader = load_ocr_reader()

# ==========================================
# 2. HÀM BỔ TRỢ NGHIỆP VỤ XNK
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
    except Exception as e:
        st.error(f"Lỗi file {uploaded_file.name}: {e}")
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

# ==========================================
# 3. GIAO DIỆN DASHBOARD CHÍNH (UI)
# ==========================================
st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>🏢 HỆ THỐNG KIỂM ĐỊNH CHỨNG TỪ XNK TỔNG HỢP</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.1em; color: #555;'>Module tự động bóc tách, quy đổi HSQĐ và đối chiếu liên thông 5 chiều: INVOICE ↔ PKL ↔ CHỈ ĐỊNH ↔ SAP ↔ ECUS</p>", unsafe_allow_html=True)
st.markdown("---")

# Khu vực Upload - Chia làm 3 khối nghiệp vụ
col_a, col_b, col_c = st.columns(3)

with col_a:
    st.info("📦 **NHÓM CHỨNG TỪ HẢI QUAN**")
    file_inv = st.file_uploader("1. INVOICE KHAI BÁO (Bắt buộc)", type=["xlsx", "csv", "pdf"])
    file_pkl = st.file_uploader("2. PACKING LIST (Bắt buộc)", type=["xlsx", "csv", "pdf"])

with col_b:
    st.warning("🚛 **NHÓM LƯU THÔNG & SỔ SÁCH**")
    file_cd  = st.file_uploader("3. CHỈ ĐỊNH GIAO NHẬN (Bắt buộc)", type=["xlsx", "csv", "pdf"])
    file_erp = st.file_uploader("4. BÁO CÁO XUẤT KHO SAP", type=["xlsx", "csv"])

with col_c:
    st.success("🏛️ **NHÓM QUY ĐỔI & THÔNG QUAN**")
    file_hsqd = st.file_uploader("5. BẢNG MASTER HSQĐ", type=["xlsx", "csv"])
    file_ecus = st.file_uploader("6. DỮ LIỆU TỜ KHAI ECUS", type=["xlsx", "csv"])

# Khu vực Sidebar Cấu Hình
with st.sidebar:
    st.markdown("### ⚙️ CẤU HÌNH HEADER (SKIPROWS)")
    st.caption("Khai báo số dòng tiêu đề thừa của từng file")
    skip_inv = st.number_input("Dòng thừa Invoice:", value=15, min_value=0)
    skip_pkl = st.number_input("Dòng thừa PKL:", value=15, min_value=0)
    skip_cd  = st.number_input("Dòng thừa Chỉ Định:", value=17, min_value=0)
    skip_erp = st.number_input("Dòng thừa SAP/ERP:", value=0, min_value=0)
    skip_hsqd = st.number_input("Dòng thừa Bảng HSQĐ:", value=1, min_value=0)
    skip_ecus = st.number_input("Dòng thừa ECUS:", value=18, min_value=0)

# ==========================================
# 4. ENGINE XỬ LÝ & ĐỐI CHIẾU CHÉO
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 KÍCH HOẠT HỆ THỐNG ĐỐI CHIẾU TOÀN DIỆN", type="primary", use_container_width=True):
    if not (file_inv and file_pkl and file_cd):
        st.error("⚠️ Lỗi: Hệ thống yêu cầu tải lên tối thiểu 3 chứng từ gốc: INVOICE, PACKING LIST và CHỈ ĐỊNH GIAO HÀNG.")
    else:
        with st.spinner("🔄 Đang nạp dữ liệu, quét OCR và chạy Engine quy đổi..."):
            
            # Đọc Raw Data
            raw_inv = extract_and_clean_data(file_inv, skip_inv)
            raw_pkl = extract_and_clean_data(file_pkl, skip_pkl)
            raw_cd  = extract_and_clean_data(file_cd, skip_cd)
            raw_erp = extract_and_clean_data(file_erp, skip_erp) if file_erp else pd.DataFrame()
            raw_hsqd = extract_and_clean_data(file_hsqd, skip_hsqd) if file_hsqd else pd.DataFrame()
            raw_ecus = extract_and_clean_data(file_ecus, skip_ecus) if file_ecus else pd.DataFrame()

            def clean_code(s): return str(s).strip().upper()

            # --- SETUP BẢNG QUY ĐỔI HSQĐ ---
            df_hsqd = pd.DataFrame(columns=['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan'])
            if not raw_hsqd.empty:
                h_mat = get_col_fallback(raw_hsqd, ['Mã số', 'Vật liệu', 'Material'], 1)
                h_rate = get_col_fallback(raw_hsqd, ['Hệ số', 'Quy đổi', 'Rate'], 6)
                h_uom = get_col_fallback(raw_hsqd, ['Đơn vị báo quan', 'Chuẩn'], 8)
                df_hsqd = raw_hsqd[[h_mat, h_rate, h_uom]].dropna(subset=[h_mat])
                df_hsqd.columns = ['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan']
                df_hsqd['Ma_Vat_Tu'] = df_hsqd['Ma_Vat_Tu'].apply(clean_code)
                df_hsqd['He_So_QD'] = pd.to_numeric(df_hsqd['He_So_QD'], errors='coerce').fillna(1.0)
                df_hsqd['DVT_Bao_Quan'] = df_hsqd['DVT_Bao_Quan'].apply(normalize_uom)

            def apply_hsqd_engine(df, qty_col, uom_col):
                if not df_hsqd.empty and not df.empty:
                    df = pd.merge(df, df_hsqd, on='Ma_Vat_Tu', how='left')
                    df['He_So_QD'] = df['He_So_QD'].fillna(1.0)
                    df[qty_col] = df[qty_col] * df['He_So_QD']
                    if uom_col in df.columns:
                        df[uom_col] = df['DVT_Bao_Quan'].combine_first(df[uom_col])
                    return df.drop(columns=['He_So_QD', 'DVT_Bao_Quan'], errors='ignore')
                return df

            # --- TRÍCH XUẤT 5 CHIỀU ---
            # 1. Invoice
            i_mat = get_col_fallback(raw_inv, ['Material code', 'Mã'], 1)
            i_qty = get_col_fallback(raw_inv, ['Quantity', 'Số lượng'], 6)
            i_uom = get_col_fallback(raw_inv, ['Unit', 'ĐVT'], 5)
            i_pri = get_col_fallback(raw_inv, ['Unit Price', 'Đơn giá'], 7)
            i_amt = get_col_fallback(raw_inv, ['Amount', 'Thành tiền'], 8)
            df_inv = raw_inv[[i_mat, i_qty, i_uom, i_pri, i_amt]].dropna(subset=[i_mat]).copy()
            df_inv.columns = ['Ma_Vat_Tu', 'SL_Invoice', 'UOM_Invoice', 'Price_Inv', 'Amount_Inv']
            df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].apply(clean_code)
            for c in ['SL_Invoice', 'Price_Inv', 'Amount_Inv']: df_inv[c] = pd.to_numeric(df_inv[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_inv = apply_hsqd_engine(df_inv, 'SL_Invoice', 'UOM_Invoice')
            df_inv['UOM_Invoice'] = df_inv['UOM_Invoice'].apply(normalize_uom)

            # 2. PKL
            p_mat = get_col_fallback(raw_pkl, ['Material', 'Mã'], 0)
            p_qty = get_col_fallback(raw_pkl, ["Q'TY", 'Quantity'], 5)
            p_uom = get_col_fallback(raw_pkl, ['Unit', 'ĐVT'], 2)
            df_pkl = raw_pkl[[p_mat, p_qty, p_uom]].dropna(subset=[p_mat]).copy()
            df_pkl.columns = ['Ma_Vat_Tu', 'SL_PKL', 'UOM_PKL']
            df_pkl['Ma_Vat_Tu'] = df_pkl['Ma_Vat_Tu'].apply(clean_code)
            df_pkl['SL_PKL'] = pd.to_numeric(df_pkl['SL_PKL'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_pkl = apply_hsqd_engine(df_pkl, 'SL_PKL', 'UOM_PKL')
            df_pkl['UOM_PKL'] = df_pkl['UOM_PKL'].apply(normalize_uom)
            df_pkl = df_pkl.groupby(['Ma_Vat_Tu', 'UOM_PKL'], as_index=False)['SL_PKL'].sum()

            # 3. Chỉ Định
            cd_mat = get_col_fallback(raw_cd, ['Mã NL', 'Material'], 1)
            cd_qty = get_col_fallback(raw_cd, ['Số Lượng', 'Quantity'], 4)
            df_cd = raw_cd[[cd_mat, cd_qty]].dropna(subset=[cd_mat]).copy()
            df_cd.columns = ['Ma_Vat_Tu', 'SL_ChiDinh']
            df_cd['Ma_Vat_Tu'] = df_cd['Ma_Vat_Tu'].apply(clean_code)
            df_cd['SL_ChiDinh'] = pd.to_numeric(df_cd['SL_ChiDinh'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_cd['UOM_CD'] = None
            df_cd = apply_hsqd_engine(df_cd, 'SL_ChiDinh', 'UOM_CD').drop(columns=['UOM_CD'])
            df_cd = df_cd.groupby('Ma_Vat_Tu', as_index=False)['SL_ChiDinh'].sum()

            # 4. ERP
            df_erp = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ERP'])
            if not raw_erp.empty:
                df_erp = raw_erp[[raw_erp.columns[0], raw_erp.columns[1]]].dropna().copy()
                df_erp.columns = ['Ma_Vat_Tu', 'SL_ERP']
                df_erp['Ma_Vat_Tu'] = df_erp['Ma_Vat_Tu'].apply(clean_code)
                df_erp['SL_ERP'] = pd.to_numeric(df_erp['SL_ERP'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_erp['UOM_ERP'] = None
                df_erp = apply_hsqd_engine(df_erp, 'SL_ERP', 'UOM_ERP').drop(columns=['UOM_ERP'])

            # 5. ECUS
            df_ecus = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ECUS'])
            if not raw_ecus.empty:
                ecus_mat = get_col_fallback(raw_ecus, ['Mô tả', 'Tên hàng'], 1)
                ecus_qty = get_col_fallback(raw_ecus, ['Số lượng (1)', 'Lượng tính thuế'], 3)
                raw_e = raw_ecus[[ecus_mat, ecus_qty]].dropna().copy()
                
                # Bóc tách mã NL từ mô tả ECUS
                known_codes = set(df_inv['Ma_Vat_Tu']).union(set(df_hsqd['Ma_Vat_Tu']))
                extracted_codes = []
                for _, r in raw_e.iterrows():
                    desc = str(r[ecus_mat]).upper()
                    found = "UNKNOWN"
                    for code in known_codes:
                        if code in desc:
                            found = code; break
                    extracted_codes.append(found)
                
                raw_e['Ma_Vat_Tu'] = extracted_codes
                df_ecus = raw_e[raw_e['Ma_Vat_Tu'] != "UNKNOWN"].copy()
                df_ecus['SL_ECUS'] = pd.to_numeric(df_ecus[ecus_qty].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_ecus = df_ecus.groupby('Ma_Vat_Tu', as_index=False)['SL_ECUS'].sum()

            # --- MERGE VÀ ĐÁNH GIÁ LỖI ---
            merged = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
            merged = pd.merge(merged, df_cd, on='Ma_Vat_Tu', how='outer')
            merged = pd.merge(merged, df_erp, on='Ma_Vat_Tu', how='outer')
            merged = pd.merge(merged, df_ecus, on='Ma_Vat_Tu', how='outer').fillna(0)

            merged['Lệch_PKL'] = merged['SL_Invoice'] - merged['SL_PKL']
            merged['Lệch_CD'] = merged['SL_Invoice'] - merged['SL_ChiDinh']
            merged['Lệch_ERP'] = merged['SL_Invoice'] - merged['SL_ERP']
            merged['Lệch_ECUS'] = merged['SL_Invoice'] - merged['SL_ECUS']

            def evaluate_row(r):
                if r['SL_Invoice'] == 0: return "❌ THIẾU MÃ TRÊN INVOICE"
                if r['UOM_Invoice'] != 0 and r['UOM_PKL'] != 0 and r['UOM_Invoice'] != r['UOM_PKL']:
                    return f"🔴 LỖI ĐVT: {r['UOM_Invoice']} != {r['UOM_PKL']}"
                if r['Price_Inv'] > 0 and r['Amount_Inv'] > 0:
                    if abs((r['SL_Invoice'] * r['Price_Inv']) - r['Amount_Inv']) > 0.1:
                        return "🔴 LỖI NHÂN TRỊ GIÁ INVOICE"

                s_pkl, s_cd, s_erp, s_hq = abs(r['Lệch_PKL']), abs(r['Lệch_CD']), abs(r['Lệch_ERP']), abs(r['Lệch_ECUS'])
                
                if s_pkl == 0 and s_cd == 0 and (s_erp == 0 or raw_erp.empty) and (s_hq == 0 or raw_ecus.empty):
                    return "🟢 KHỚP HOÀN TOÀN"
                
                # Dung sai
                tol = 0.2 if r['UOM_Invoice'] in ['KGM', 'MTK'] else 0.01
                if s_pkl <= tol and s_cd <= tol and s_erp <= tol and s_hq <= tol:
                    return "🟡 KHỚP (DUNG SAI LÀM TRÒN)"

                if s_pkl > tol: return "🔴 LỆCH PACKING LIST"
                if s_cd > tol: return "🔴 LỆCH CHỈ ĐỊNH GIAO HÀNG"
                if not raw_ecus.empty and s_hq > tol: return "🔴 LỆCH TỜ KHAI ECUS"
                return "🟠 LỆCH SỔ SÁCH ERP"

            merged['ĐÁNH GIÁ (STATUS)'] = merged.apply(evaluate_row, axis=1)
            
            # Lọc cột hiển thị
            display_cols = ['Ma_Vat_Tu', 'UOM_Invoice', 'SL_Invoice', 'SL_PKL', 'SL_ChiDinh']
            if not raw_erp.empty: display_cols.extend(['SL_ERP', 'Lệch_ERP'])
            if not raw_ecus.empty: display_cols.extend(['SL_ECUS', 'Lệch_ECUS'])
            display_cols.append('ĐÁNH GIÁ (STATUS)')
            
            merged_display = merged[display_cols].sort_values(by='ĐÁNH GIÁ (STATUS)', ascending=False)

            st.markdown("---")
            st.markdown("### 📊 DASHBOARD KẾT QUẢ ĐỐI CHIẾU")
            
            # Toggle Filter
            err_filter = st.toggle("🚨 BẬT BỘ LỌC LỖI (Chỉ hiển thị các mã bị vênh số lượng / ĐVT)", value=False)
            if err_filter:
                final_board = merged_display[~merged_display['ĐÁNH GIÁ (STATUS)'].str.contains("KHỚP")]
            else:
                final_board = merged_display

            def color_matrix(val):
                if "🟢 KHỚP" in str(val) or "🟡 KHỚP" in str(val): return 'background-color: #d1fae5; color: #065f46; font-weight: bold'
                if "🔴 LỆCH" in str(val) or "🔴 LỖI" in str(val): return 'background-color: #fee2e2; color: #991b1b; font-weight: bold'
                if "❌" in str(val): return 'background-color: #fce7f3; color: #9d174d; font-weight: bold'
                return 'background-color: #fef3c7; color: #b45309; font-weight: bold'

            st.dataframe(final_board.style.map(color_matrix, subset=['ĐÁNH GIÁ (STATUS)']), use_container_width=True, height=500)

            # --- XUẤT OUTPUT ---
            st.markdown("<br>", unsafe_allow_html=True)
            exp1, exp2 = st.columns(2)
            with exp1:
                excel_buf = io.BytesIO()
                with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
                    merged_display.to_excel(writer, index=False, sheet_name='E54_Checked')
                    writer.sheets['E54_Checked'].set_column(0, 0, 25)
                st.download_button("📊 XUẤT FILE EXCEL (CHUẨN HÓA)", data=excel_buf.getvalue(), file_name="E54_CrossCheck_Output.xlsx", type="primary", use_container_width=True)
            with exp2:
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=14, style='B')
                pdf.cell(200, 10, txt="BIEN BAN KIEM TRA CHUNG TU E54 / E23", ln=1, align="C")
                pdf.set_font("Arial", size=10)
                err_count = len(merged_display[~merged_display['ĐÁNH GIÁ (STATUS)'].str.contains("KHỚP")])
                pdf.cell(200, 10, txt=f"Phat hien {err_count} dong vat tu bi lech so lieu.", ln=2)
                st.download_button("📑 XUẤT BIÊN BẢN BÁO CÁO (PDF)", data=bytes(pdf.output()), file_name="E54_Report.pdf", use_container_width=True)
