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
from datetime import datetime, timedelta

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & BẢO MẬT (WHITE-LABEL)
# ==========================================
st.set_page_config(layout="wide", page_title="Hệ Sinh Thái XNK E54-E23", page_icon="🌐", initial_sidebar_state="expanded")

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
            st.markdown("<h2 style='text-align: center; color: #1E3A8A; padding-top: 50px;'>🌐 SCM EXIM - Team Nhập</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center;'>Hệ thống liên thông dữ liệu Ching Luh & Fu-Luh khi khai báo E54 E23</p>", unsafe_allow_html=True)
            pwd = st.text_input("Nhập mã truy cập an ninh:", type="password")
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

# ==========================================
# 2. KHỞI TẠO BỘ NHỚ TRUNG TÂM (DATABASE ẢO)
# ==========================================
if 'master_hsqd' not in st.session_state: st.session_state.master_hsqd = pd.DataFrame()
if 'master_thue' not in st.session_state: st.session_state.master_thue = pd.DataFrame()
if 'chingluh_export_data' not in st.session_state: st.session_state.chingluh_export_data = pd.DataFrame()
if 'invoice_date' not in st.session_state: st.session_state.invoice_date = datetime.today()

# ==========================================
# 3. LÕI THUẬT TOÁN (ENGINE)
# ==========================================
def extract_and_clean_data(uploaded_file, skiprows_count=0):
    if uploaded_file is None: return None
    ext = uploaded_file.name.split('.')[-1].lower()
    try:
        if ext in ['xlsx', 'xls', 'csv']:
            df = pd.read_csv(uploaded_file, skiprows=skiprows_count, on_bad_lines='skip') if ext == 'csv' else pd.read_excel(uploaded_file, skiprows=skiprows_count)
            # SỬA LỖI TẠI ĐÂY: Ép kiểu string cho toàn bộ tên cột trước khi lọc 'Unnamed'
            df.columns = [str(c).strip() for c in df.columns]
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df = df.dropna(how='all')
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
            return pd.DataFrame(reader.readtext(np.array(img), detail=0), columns=["Dữ_Liệu_OCR"])
    except Exception as e: st.error(f"Lỗi file {uploaded_file.name}: {e}")
    return None

def norm_uom(uom):
    if pd.isna(uom) or str(uom).strip() == '': return "PCE"
    d = {'PC':'PCE', 'UNIT':'PCE', 'PCS':'PCE', 'PR':'PRS', 'PAIRS':'PRS', 'KG':'KGM', 'KGS':'KGM', 'M':'MTR', 'YD':'YRD', 'YDS':'YRD', 'M2':'MTK'}
    return d.get(str(uom).strip().upper(), str(uom).strip().upper())

def get_col(df, kws, fb_idx):
    for c in df.columns:
        for k in kws:
            if k.lower() in str(c).lower(): return c
    if fb_idx is not None and len(df.columns) > fb_idx:
        return df.columns[fb_idx]
    return None

def fz_match(code, masters):
    if code in masters: return code
    m = difflib.get_close_matches(code, masters, n=1, cutoff=0.85)
    return m[0] if m else code

def apply_hsqd(df, qty_col, uom_col):
    hsqd = st.session_state.master_hsqd
    if not hsqd.empty and not df.empty:
        df = pd.merge(df, hsqd, on='Ma_Vat_Tu', how='left')
        df['He_So_QD'] = df['He_So_QD'].fillna(1.0)
        df[qty_col] = df[qty_col] * df['He_So_QD']
        if uom_col in df.columns: df[uom_col] = df['DVT_Bao_Quan'].combine_first(df[uom_col])
        return df.drop(columns=['He_So_QD', 'DVT_Bao_Quan'], errors='ignore')
    return df

# ==========================================
# 4. HỆ THỐNG ĐIỀU HƯỚNG TỪNG PHÂN HỆ
# ==========================================
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a7/React-icon.svg/120px-React-icon.svg.png", width=50) 
st.sidebar.markdown("## 🌐 ĐIỀU HÀNH XNK")
menu = st.sidebar.radio("CHỌN KHÔNG GIAN LÀM VIỆC:", [
    "1️⃣ BÊN XUẤT: CHING LUH (E54)", 
    "2️⃣ BÊN NHẬP: FU-LUH (E23)", 
    "3️⃣ MASTER DATA & BÁO CÁO"
])

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ CẤU HÌNH HEADER FILE")
skip_inv = st.sidebar.number_input("Invoice/PKL:", value=15)
skip_cd  = st.sidebar.number_input("Chỉ Định:", value=17)
skip_erp = st.sidebar.number_input("SAP/ERP:", value=0)
skip_ecus = st.sidebar.number_input("Tờ Khai ECUS:", value=18)

# ==============================================================================
# PHÂN HỆ 3: MASTER DATA
# ==============================================================================
if menu == "3️⃣ MASTER DATA & BÁO CÁO":
    st.markdown("<h2 style='color: #4c1d95;'>⚙️ QUẢN LÝ DỮ LIỆU NỀN (MASTER DATA)</h2>", unsafe_allow_html=True)
    st.info("💡 Tải lên Bảng Hệ số Quy đổi và Biểu thuế tại đây. Dữ liệu sẽ được dùng chung cho cả Ching Luh và Fu-Luh.")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        f_hsqd = st.file_uploader("📥 BẢNG HỆ SỐ QUY ĐỔI (HSQĐ)", type=["xlsx", "csv"])
        if f_hsqd:
            raw = extract_and_clean_data(f_hsqd, 1)
            if not raw.empty:
                df = raw[[get_col(raw, ['Mã'], 1), get_col(raw, ['Hệ số'], 6), get_col(raw, ['Đơn vị báo quan'], 8)]].dropna()
                df.columns = ['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan']
                df['Ma_Vat_Tu'] = df['Ma_Vat_Tu'].astype(str).str.strip().str.upper()
                df['He_So_QD'] = pd.to_numeric(df['He_So_QD'], errors='coerce').fillna(1.0)
                df['DVT_Bao_Quan'] = df['DVT_Bao_Quan'].apply(norm_uom)
                st.session_state.master_hsqd = df
                st.success(f"✅ Đã nạp thành công {len(df)} mã HSQĐ vào bộ nhớ.")

    with col_m2:
        f_thue = st.file_uploader("📥 BẢNG BIỂU THUẾ NĂM", type=["xlsx", "csv"])
        if f_thue:
            raw = extract_and_clean_data(f_thue, 2)
            if not raw.empty:
                df = raw[[get_col(raw, ['Mã hàng', 'HS'], 1), get_col(raw, ['Thuế suất', 'NK'], 5)]].dropna()
                df.columns = ['HS_Code', 'Thue_Suat']
                df['HS_Code'] = df['HS_Code'].astype(str).str.replace('.', '').str.strip()
                st.session_state.master_thue = df
                st.success(f"✅ Đã nạp thành công {len(df)} mã Biểu thuế vào bộ nhớ.")

    st.markdown("---")
    st.markdown("### 📈 THỐNG KÊ TỔNG QUAN")
    if not st.session_state.chingluh_export_data.empty:
        st.write("Dữ liệu Hàng đang đi đường (Transit):")
        st.dataframe(st.session_state.chingluh_export_data.head())
    else:
        st.caption("Chưa có lô hàng nào được xuất từ hệ thống Ching Luh.")

# ==============================================================================
# PHÂN HỆ 1: CHING LUH (XUẤT - E54)
# ==============================================================================
elif menu == "1️⃣ BÊN XUẤT: CHING LUH (E54)":
    st.markdown("<h2 style='color: #1d4ed8;'>📤 PHÂN HỆ XUẤT HÀNG (CHING LUH - E54)</h2>", unsafe_allow_html=True)
    
    st.session_state.invoice_date = st.date_input("🗓️ Ngày lập Invoice / Xuất hàng:", datetime.today())
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("##### 📦 1. Bộ Chứng Từ")
        f_inv = st.file_uploader("INVOICE KHAI BÁO", type=["xlsx", "csv"])
        f_pkl = st.file_uploader("PACKING LIST", type=["xlsx", "csv"])
        f_cd = st.file_uploader("CHỈ ĐỊNH GIAO HÀNG", type=["xlsx", "csv"])
    with c2:
        st.markdown("##### 💻 2. Kho Xuất (Ching Luh)")
        f_erp_out = st.file_uploader("SAP/ERP XUẤT KHO", type=["xlsx", "csv"])
    with c3:
        st.markdown("##### 🏛️ 3. Hải Quan E54")
        f_e54 = st.file_uploader("TỜ KHAI XUẤT E54", type=["xlsx", "csv"])

    if st.button("🔄 ĐỐI CHIẾU XUẤT & CHỐT DỮ LIỆU ĐI ĐƯỜNG", type="primary", use_container_width=True):
        if not (f_inv and f_pkl):
            st.error("⚠️ Bắt buộc phải có INVOICE và PACKING LIST!")
        else:
            with st.spinner("Đang tổng hợp chứng từ xuất..."):
                raw_inv = extract_and_clean_data(f_inv, skip_inv)
                raw_pkl = extract_and_clean_data(f_pkl, skip_inv)
                raw_erp = extract_and_clean_data(f_erp_out, skip_erp) if f_erp_out else pd.DataFrame()
                raw_e54 = extract_and_clean_data(f_e54, skip_ecus) if f_e54 else pd.DataFrame()

                # Xử lý Invoice
                df_inv = raw_inv[[get_col(raw_inv, ['Material', 'Mã'], 1), get_col(raw_inv, ['Quantity'], 6), get_col(raw_inv, ['Unit'], 5), get_col(raw_inv, ['Price'], 7), get_col(raw_inv, ['Amount'], 8)]].dropna().copy()
                df_inv.columns = ['Ma_Vat_Tu', 'SL_INV', 'UOM_INV', 'DonGia', 'TriGia']
                df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].astype(str).str.strip().str.upper()
                for c in ['SL_INV', 'DonGia', 'TriGia']: df_inv[c] = pd.to_numeric(df_inv[c].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                df_inv = apply_hsqd(df_inv, 'SL_INV', 'UOM_INV')
                df_inv['UOM_INV'] = df_inv['UOM_INV'].apply(norm_uom)
                masters = df_inv['Ma_Vat_Tu'].tolist()

                # --- SỬA LỖI 2 TẠI ĐÂY: XỬ LÝ PACKING LIST AN TOÀN TUYỆT ĐỐI ---
                p_mat = get_col(raw_pkl, ['Material', 'Mã'], 0)
                p_qty = get_col(raw_pkl, ["Q'TY", 'Quantity'], 5)
                p_uom = get_col(raw_pkl, ['Unit', 'ĐVT'], 2)
                nw_col = get_col(raw_pkl, ['NW', 'Net Weight', 'N.W'], None)
                gw_col = get_col(raw_pkl, ['GW', 'Gross Weight', 'G.W'], None)
                
                df_pkl_raw = pd.DataFrame()
                df_pkl_raw['Ma_Vat_Tu'] = raw_pkl[p_mat].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters))
                df_pkl_raw['SL_PKL'] = pd.to_numeric(raw_pkl[p_qty].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                df_pkl_raw['UOM_PKL'] = raw_pkl[p_uom].apply(norm_uom)
                
                if nw_col and nw_col in raw_pkl.columns:
                    df_pkl_raw['NW'] = pd.to_numeric(raw_pkl[nw_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                if gw_col and gw_col in raw_pkl.columns:
                    df_pkl_raw['GW'] = pd.to_numeric(raw_pkl[gw_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                
                df_pkl_raw = df_pkl_raw.dropna(subset=['Ma_Vat_Tu'])
                
                df_pkl = apply_hsqd(df_pkl_raw, 'SL_PKL', 'UOM_PKL')
                df_pkl = df_pkl.groupby('Ma_Vat_Tu', as_index=False)['SL_PKL'].sum()

                # Xử lý ERP Xuất
                df_erp = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_CL_XUAT'])
                if not raw_erp.empty:
                    df_erp = raw_erp[[raw_erp.columns[0], raw_erp.columns[1]]].dropna().copy()
                    df_erp.columns = ['Ma_Vat_Tu', 'SL_CL_XUAT']
                    df_erp['Ma_Vat_Tu'] = df_erp['Ma_Vat_Tu'].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters))
                    df_erp['SL_CL_XUAT'] = pd.to_numeric(df_erp['SL_CL_XUAT'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    df_erp['UOM_ERP'] = None
                    df_erp = apply_hsqd(df_erp, 'SL_CL_XUAT', 'UOM_ERP').drop(columns=['UOM_ERP'])

                # Xử lý E54
                df_e54 = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_E54'])
                if not raw_e54.empty:
                    c_desc = get_col(raw_e54, ['Mô tả'], 1)
                    c_qty = get_col(raw_e54, ['Số lượng (1)'], 3)
                    tmp = raw_e54[[c_desc, c_qty]].dropna().copy()
                    codes = []
                    for _, r in tmp.iterrows():
                        d = str(r[c_desc]).upper()
                        f = "UNK"
                        for c in masters:
                            if c in d: f = c; break
                        codes.append(f)
                    tmp['Ma_Vat_Tu'] = codes
                    df_e54 = tmp[tmp['Ma_Vat_Tu'] != 'UNK'].copy()
                    df_e54['SL_E54'] = pd.to_numeric(df_e54[c_qty].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    df_e54 = df_e54.groupby('Ma_Vat_Tu', as_index=False)['SL_E54'].sum()

                # Merge
                mg = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
                mg = pd.merge(mg, df_erp, on='Ma_Vat_Tu', how='outer')
                mg = pd.merge(mg, df_e54, on='Ma_Vat_Tu', how='outer').fillna(0)

                def check_out(r):
                    if r['SL_INV'] == 0: return "❌ LỖI: KHÔNG CÓ TRÊN INVOICE"
                    if abs(r['SL_INV'] - r['SL_PKL']) > 0.1: return "🔴 LỆCH ĐÓNG GÓI (PKL)"
                    if 'SL_CL_XUAT' in r and abs(r['SL_INV'] - r['SL_CL_XUAT']) > 0.1: return "🟠 LỆCH KHO XUẤT CHING LUH"
                    if 'SL_E54' in r and abs(r['SL_INV'] - r['SL_E54']) > 0.1: return "🔴 LỆCH TỜ KHAI E54"
                    return "🟢 SẴN SÀNG XUẤT"

                mg['TRẠNG THÁI XUẤT'] = mg.apply(check_out, axis=1)
                
                st.session_state.chingluh_export_data = mg[['Ma_Vat_Tu', 'UOM_INV', 'SL_INV', 'DonGia', 'TriGia', 'TRẠNG THÁI XUẤT']].copy()

                st.success("✅ Đã xử lý phân hệ Xuất. Dữ liệu lô hàng đã được đẩy lên Không gian Trung chuyển!")
                
                # BẢNG ĐỐI CHIẾU LIVE EDITOR CHO PHÂN HỆ 1
                tab1, tab2, tab3 = st.tabs(["📊 BẢNG ĐỐI CHIẾU (LIVE EDITOR)", "⚖️ TRỌNG LƯỢNG (N.W/G.W)", "📈 DASHBOARD THỐNG KÊ LỖI"])
                with tab1:
                    edited_df_out = st.data_editor(mg.style.applymap(lambda x: 'background-color:#d1fae5' if '🟢' in str(x) else 'background-color:#fee2e2' if '🔴' in str(x) else '', subset=['TRẠNG THÁI XUẤT']), use_container_width=True, hide_index=True)
                    excel_buf_out = io.BytesIO()
                    with pd.ExcelWriter(excel_buf_out, engine='xlsxwriter') as writer:
                        edited_df_out.to_excel(writer, index=False, sheet_name='E54_Export_Checked')
                    st.download_button("📥 TẢI FILE EXCEL XUẤT KHẨU", data=excel_buf_out.getvalue(), file_name="E54_Export_Output.xlsx", type="primary")
                with tab2:
                    if 'NW' in df_pkl_raw.columns or 'GW' in df_pkl_raw.columns:
                        w1, w2 = st.columns(2)
                        with w1: st.metric("Tổng Trọng Lượng Tịnh (N.W)", f"{df_pkl_raw['NW'].sum():,.2f} KGM")
                        with w2: st.metric("Tổng Trọng Lượng Cộp (G.W)", f"{df_pkl_raw['GW'].sum():,.2f} KGM")
                    else: st.warning("Packing List không có cột Trọng lượng.")
                with tab3:
                    st.plotly_chart(px.pie(mg['TRẠNG THÁI XUẤT'].value_counts().reset_index(), values='count', names='TRẠNG THÁI XUẤT', title='Tỷ lệ khớp số liệu xuất'), use_container_width=True)

# ==============================================================================
# PHÂN HỆ 2: FU-LUH (NHẬP - E23)
# ==============================================================================
elif menu == "2️⃣ BÊN NHẬP: FU-LUH (E23)":
    st.markdown("<h2 style='color: #ea580c;'>📥 PHÂN HỆ NHẬP HÀNG (FU-LUH - E23)</h2>", unsafe_allow_html=True)
    
    inv_date = st.session_state.invoice_date
    deadline = inv_date + timedelta(days=15)
    days_left = (deadline - datetime.today().date()).days
    
    st.markdown("#### ⏱️ THEO DÕI THỜI HẠN THÔNG QUAN TẠI CHỖ (15 NGÀY)")
    if days_left < 0: st.error(f"🚨 CẢNH BÁO: Đã trễ hạn hải quan {abs(days_left)} ngày! (Hạn chót: {deadline.strftime('%d/%m/%Y')})")
    elif days_left <= 3: st.warning(f"⚠️ CHÚ Ý: Gấp rút mở E23. Chỉ còn {days_left} ngày (Hạn chót: {deadline.strftime('%d/%m/%Y')})")
    else: st.success(f"✅ Thời gian an toàn. Còn {days_left} ngày (Hạn chót: {deadline.strftime('%d/%m/%Y')})")

    st.markdown("---")
    
    transit_data = st.session_state.chingluh_export_data
    if transit_data.empty:
        st.warning("📭 Chưa có Dữ liệu hàng đi đường từ Ching Luh chuyển sang. Vui lòng quay lại Phân hệ 1 để duyệt xuất hàng trước.")
    else:
        st.info(f"🚚 Nhận được dữ liệu lô hàng gồm {len(transit_data)} vật tư đang đi đường từ Ching Luh.")
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 💻 Kho Nhập (Fu-Luh)")
            f_erp_in = st.file_uploader("SAP/ERP NHẬP KHO THỰC TẾ", type=["xlsx", "csv"])
        with c2:
            st.markdown("##### 🏛️ Hải Quan E23")
            f_e23 = st.file_uploader("TỜ KHAI NHẬP E23", type=["xlsx", "csv"])

        if st.button("🔍 ĐỐI CHIẾU NHẬP KHO & KHAI BÁO E23", type="primary", use_container_width=True):
            with st.spinner("Đang đối chiếu dữ liệu thực nhận với hàng đi đường..."):
                r_erp_in = extract_and_clean_data(f_erp_in, skip_erp) if f_erp_in else pd.DataFrame()
                r_e23 = extract_and_clean_data(f_e23, skip_ecus) if f_e23 else pd.DataFrame()
                
                masters = transit_data['Ma_Vat_Tu'].tolist()
                
                df_erp_in = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_FL_NHAP'])
                if not r_erp_in.empty:
                    df_erp_in = r_erp_in[[r_erp_in.columns[0], r_erp_in.columns[1]]].dropna().copy()
                    df_erp_in.columns = ['Ma_Vat_Tu', 'SL_FL_NHAP']
                    df_erp_in['Ma_Vat_Tu'] = df_erp_in['Ma_Vat_Tu'].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters))
                    df_erp_in['SL_FL_NHAP'] = pd.to_numeric(df_erp_in['SL_FL_NHAP'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    df_erp_in['UOM_ERP'] = None
                    df_erp_in = apply_hsqd(df_erp_in, 'SL_FL_NHAP', 'UOM_ERP').drop(columns=['UOM_ERP'])

                df_e23 = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_E23', 'HS_Code'])
                if not r_e23.empty:
                    c_desc = get_col(r_e23, ['Mô tả'], 1)
                    c_qty = get_col(r_e23, ['Số lượng (1)'], 3)
                    c_hs = get_col(r_e23, ['Mã số'], 0)
                    tmp = r_e23[[c_desc, c_qty, c_hs]].dropna().copy()
                    codes = []
                    for _, r in tmp.iterrows():
                        d = str(r[c_desc]).upper()
                        f = "UNK"
                        for c in masters:
                            if c in d: f = c; break
                        codes.append(f)
                    tmp['Ma_Vat_Tu'] = codes
                    df_e23 = tmp[tmp['Ma_Vat_Tu'] != 'UNK'].copy()
                    df_e23['SL_E23'] = pd.to_numeric(df_e23[c_qty].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    df_e23.rename(columns={c_hs: 'HS_Code'}, inplace=True)
                    df_e23 = df_e23.groupby(['Ma_Vat_Tu', 'HS_Code'], as_index=False)['SL_E23'].sum()

                mg = pd.merge(transit_data, df_erp_in, on='Ma_Vat_Tu', how='left')
                mg = pd.merge(mg, df_e23, on='Ma_Vat_Tu', how='left').fillna(0)

                mg['Thue_Du_Kien'] = 0.0
                if not st.session_state.master_thue.empty and not df_e23.empty:
                    thue_df = st.session_state.master_thue
                    thue_dict = dict(zip(thue_df['HS_Code'], thue_df['Thue_Suat']))
                    mg['Thue_Suat'] = mg['HS_Code'].astype(str).str.replace('.','').map(thue_dict).fillna(0)
                    mg['Thue_Du_Kien'] = mg['TriGia'] * pd.to_numeric(mg['Thue_Suat'], errors='coerce').fillna(0) / 100

                def check_in(r):
                    s_erp = abs(r['SL_INV'] - r['SL_FL_NHAP']) if 'SL_FL_NHAP' in r else 0
                    s_e23 = abs(r['SL_INV'] - r['SL_E23']) if 'SL_E23' in r else 0
                    if s_erp == 0 and s_e23 == 0: return "🟢 KHỚP (NHẬN ĐỦ, KHAI ĐÚNG)"
                    tol = 0.2 if r['UOM_INV'] in ['KGM', 'MTK'] else 0.01
                    if s_erp <= tol and s_e23 <= tol: return "🟡 KHỚP (DUNG SAI)"
                    if s_erp > tol: return "🚨 LỖI THẤT THOÁT: KHO FU-LUH NHẬP THIẾU"
                    return "🔴 LỆCH TỜ KHAI E23"

                mg['ĐÁNH GIÁ NHẬP & KHAI BÁO'] = mg.apply(check_in, axis=1)
                mg = mg.sort_values(by='ĐÁNH GIÁ NHẬP & KHAI BÁO', ascending=False)

                st.markdown("### 📊 KẾT QUẢ ĐỐI CHIẾU NHẬP HÀNG E23")
                edited_mg = st.data_editor(
                    mg.style.applymap(lambda x: 'background-color:#d1fae5' if '🟢' in str(x) or '🟡' in str(x) else 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🚨' in str(x) or '🔴' in str(x) else '', subset=['ĐÁNH GIÁ NHẬP & KHAI BÁO']), 
                    use_container_width=True, hide_index=True
                )
                
                tong_thue = mg['Thue_Du_Kien'].sum()
                if tong_thue > 0: st.info(f"💸 **Dự toán Thuế Nhập Khẩu:** Khoảng ${tong_thue:,.2f} USD dựa trên Biểu thuế.")

                st.markdown("<br>", unsafe_allow_html=True)
                excel_buf = io.BytesIO()
                with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
                    edited_mg.to_excel(writer, index=False, sheet_name='E23_Final_Clearance')
                st.download_button("📥 XUẤT BÁO CÁO QUYẾT TOÁN LÔ HÀNG (EXCEL)", data=excel_buf.getvalue(), file_name="HoSo_QuyetToan_E54_E23.xlsx", type="primary", use_container_width=True)
