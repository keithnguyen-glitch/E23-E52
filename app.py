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
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. TỪ ĐIỂN ĐA NGÔN NGỮ (VI - EN - ZH)
# ==========================================
LANG_DICT = {
    "🇻🇳 Tiếng Việt": {
        "gate_title": "🌐 CLG SCM EXIM _ IMPORT TEAM E54 & E23",
        "gate_sub": "Đăng nhập để sử dụng công cụ đối soát chứng từ",
        "pwd_label": "Nhập mã an ninh:",
        "pwd_err": "❌ Mã truy cập không hợp lệ!",
        "main_title": "🏢 HỆ THỐNG ĐỐI SOÁT CHỨNG TỪ E54 - E23",
        "main_sub": "Tối ưu hóa tác vụ: Tách mã R ➔ Đối soát ma trận số lượng ➔ Sinh chuỗi Hải quan tự động.",
        "config_side": "### ⚖️ THAM SỐ DUNG SAI",
        "tol_w": "Hàng Vải/Cân ký (KGM, MTK):",
        "tol_c": "Hàng Đếm chiếc (PCE, PRS):",
        "config_skip": "### ⚙️ DÒNG THỪA (SKIPROWS)",
        "skip_inv": "Dòng thừa Invoice/PKL:",
        "skip_cd": "Dòng thừa Chỉ Định:",
        "skip_erp": "Dòng thừa SAP (ZMM12):",
        "skip_ecus": "Dòng thừa Tờ Khai Hải Quan:",
        "phase0_title": "#### ⚡ BƯỚC 1: TRÍCH XUẤT MÃ LIỆU R TIỀN TRẠM",
        "phase0_sub": "Tải chứng từ thô từ kho gửi để lấy danh sách mã liệu đi truy vấn SAP ZMM12.",
        "pre_inv_lbl": "1. Invoice kho gửi (Tệp thô):",
        "pre_pkl_lbl": "2. Packing List kho gửi (Tệp thô):",
        "phase0_succ": "📋 Tự động phát hiện {} Mã liệu R độc nhất.",
        "phase0_copy": "👉 Copy danh sách mã dưới đây dán vào mục Multi-Selection trên SAP (ZMM12):",
        "phase1_title": "#### 📦 BƯỚC 2: ĐỐI SOÁT GIAO NHẬN THỰC TẾ & SỔ KHO SAP",
        "phase1_sub": "Hệ thống tự động gộp bảng và khớp nối số lượng giữa chứng từ gốc và dữ liệu tồn kho.",
        "f_inv_lbl": "1. INVOICE KHAI BÁO:",
        "f_pkl_lbl": "2. PACKING LIST KHAI BÁO:",
        "f_cd_lbl": "3. CHỈ ĐỊNH GIAO HÀNG:",
        "f_sap_lbl": "4. BÁO CÁO SAP ZMM12:",
        "tbl1_title": "##### 📊 KẾT QUẢ ĐỐI SOÁT GIAO NHẬN THỰC TẾ",
        "phase2_title": "#### 🏛️ BƯỚC 3: ĐỒNG BỘ TỜ KHAI HẢI QUAN & TRÍCH XUẤT",
        "phase2_sub": "Khớp nối chiều cuối cùng với file Hải quan và tự động sinh chuỗi khai báo.",
        "f_hq_lbl": "1. FILE TỜ KHAI HẢI QUAN (TAB HANG):",
        "tbl2_title": "##### 📊 BẢNG TỔNG HỢP KIỂM CHÉO LIÊN THÔNG",
        "toggle_err": "🚨 Bật: Chỉ hiển thị các dòng bị lệch số lượng",
        "tax_msg": "💸 Dự toán Thuế lô hàng Nhập Khẩu: Khoảng ${:,.2f} USD.",
        "btn_xlsx": "📊 TẢI EXCEL KẾT QUẢ ĐỐI CHIẾU",
        "btn_docx": "📝 TẢI BIÊN BẢN WORD XÁC NHẬN",
        "err_miss": "⚠️ Yêu cầu tối thiểu: Phải có Invoice khai báo và SAP ZMM12."
    },
    "🇺🇸 English": {
        "gate_title": "🌐 EXIM INTERNAL PORTAL",
        "gate_sub": "Log in to access the document reconciliation tool",
        "pwd_label": "Enter Security Code:",
        "pwd_err": "❌ Invalid Access Code!",
        "main_title": "🏢 E23 - E54 DOCUMENT CROSS-CHECK PLATFORM",
        "main_sub": "Task Optimization: Extract R-code ➔ Matrix Reconciliation ➔ Auto Customs String.",
        "config_side": "### ⚖️ TOLERANCE",
        "tol_w": "Fabric/Weight (KGM, MTK):",
        "tol_c": "Countables (PCE, PRS):",
        "config_skip": "### ⚙️ SKIPROWS",
        "skip_inv": "Invoice/PKL Rows:",
        "skip_cd": "Shipping Inst. Rows:",
        "skip_erp": "SAP (ZMM12) Rows:",
        "skip_ecus": "Customs Form Rows:",
        "phase0_title": "#### ⚡ STEP 1: PRE-ROUTE R-CODE EXTRACTION",
        "phase0_sub": "Upload raw docs to extract material codes for SAP ZMM12 query.",
        "pre_inv_lbl": "1. Raw Invoice:",
        "pre_pkl_lbl": "2. Raw Packing List:",
        "phase0_succ": "📋 Auto-detected {} unique Material R codes.",
        "phase0_copy": "👉 Copy the list below for SAP ZMM12 Multi-Selection:",
        "phase1_title": "#### 📦 STEP 2: ACTUAL DELIVERY VS SAP RECONCILIATION",
        "phase1_sub": "Auto-merge and cross-match quantities across commercial docs and system inventory.",
        "f_inv_lbl": "1. OFFICIAL INVOICE:",
        "f_pkl_lbl": "2. OFFICIAL PACKING LIST:",
        "f_cd_lbl": "3. SHIPPING INSTRUCTION:",
        "f_sap_lbl": "4. SAP ZMM12 REPORT:",
        "tbl1_title": "##### 📊 DELIVERY RECONCILIATION RESULTS",
        "phase2_title": "#### 🏛️ STEP 3: CUSTOMS SYNC & REPORT EXPORT",
        "phase2_sub": "Final validation against Customs data and auto-string generation.",
        "f_hq_lbl": "1. CUSTOMS DECLARATION (HANG TAB):",
        "tbl2_title": "##### 📊 CONSOLIDATED AUDIT MATRIX",
        "toggle_err": "🚨 Filter: Show only discrepancies",
        "tax_msg": "💸 Estimated Inbound Customs Duty: Approx ${:,.2f} USD.",
        "btn_xlsx": "📊 DOWNLOAD EXCEL AUDIT REPORT",
        "btn_docx": "📝 DOWNLOAD WORD CONFIRMATION",
        "err_miss": "⚠️ Required: Official Invoice and SAP ZMM12."
    },
    "🇨🇳 中文": {
        "gate_title": "🌐 进出口内部控制门户",
        "gate_sub": "登录以访问单证核对工具",
        "pwd_label": "请输入安全密码:",
        "pwd_err": "❌ 访问密码错误!",
        "main_title": "🏢 E23 - E54 单证交叉核对平台",
        "main_sub": "任务优化：提取 R 码 ➔ 数量矩阵核对 ➔ 自动生成海关描述。",
        "config_side": "### ⚖️ 容差设置",
        "tol_w": "面料/称重类 (KGM, MTK):",
        "tol_c": "数量类 (PCE, PRS):",
        "config_skip": "### ⚙️ 冗余行 (SKIPROWS)",
        "skip_inv": "发票/装箱单:",
        "skip_cd": "出货通知书:",
        "skip_erp": "SAP (ZMM12):",
        "skip_ecus": "海关报关单:",
        "phase0_title": "#### ⚡ 步骤 1: 提取原始 R 码",
        "phase0_sub": "上传原始单据以提取物料代码，用于查询 SAP ZMM12。",
        "pre_inv_lbl": "1. 原始发票:",
        "pre_pkl_lbl": "2. 原始装箱单:",
        "phase0_succ": "📋 自动检测到 {} 个唯一的物料 R 码。",
        "phase0_copy": "👇 复制下方列表至 SAP ZMM12 多项选择框:",
        "phase1_title": "#### 📦 步骤 2: 实际交货与 SAP 核对",
        "phase1_sub": "在商业单证和系统库存之间自动合并和交叉匹配数量。",
        "f_inv_lbl": "1. 官方发票:",
        "f_pkl_lbl": "2. 官方装箱单:",
        "f_cd_lbl": "3. 出货通知书:",
        "f_sap_lbl": "4. SAP ZMM12 报表:",
        "tbl1_title": "##### 📊 实际交货核对结果",
        "phase2_title": "#### 🏛️ 步骤 3: 海关同步与报表导出",
        "phase2_sub": "与海关数据进行最终验证并自动生成字符串。",
        "f_hq_lbl": "1. 海关报关单 (HANG 表):",
        "tbl2_title": "##### 📊 综合审计矩阵",
        "toggle_err": "🚨 仅显示差异项",
        "tax_msg": "💸 进口关税预估：约 ${:,.2f} USD。",
        "btn_xlsx": "📊 下载 Excel 审计报告",
        "btn_docx": "📝 下载 Word 确认函",
        "err_miss": "⚠️ 必填项：官方发票和 SAP ZMM12。"
    }
}

# ==========================================
# 2. CẤU HÌNH UI SIÊU LỚN (BIG UI) & BẢO MẬT
# ==========================================
st.set_page_config(layout="wide", page_title="EXIM Reconciliation", page_icon="🌐", initial_sidebar_state="expanded")

# Tùy chỉnh CSS để làm UI to hơn, rộng hơn, dễ nhìn hơn (Tăng font-size)
big_ui_style = """
    <style>
    /* Xóa các thành phần mặc định của Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Phóng to font chữ toàn cục */
    html, body, [class*="st-"] {
        font-size: 1.15rem !important; 
    }
    
    /* Phóng to các Header */
    h1 {font-size: 2.5rem !important; font-weight: 800 !important; color: #1E3A8A;}
    h3 {font-size: 1.8rem !important; font-weight: 700 !important;}
    h4 {font-size: 1.5rem !important; font-weight: 700 !important;}
    h5 {font-size: 1.3rem !important; font-weight: 600 !important; color: #b45309;}
    
    /* Làm nổi bật File Uploader */
    .stFileUploader label {
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    
    /* Bảng Dataframe to và rõ ràng hơn */
    [data-testid="stDataFrame"] {
        font-size: 1.1rem;
    }
    
    /* Style nền sáng sủa, sạch sẽ */
    .block-container {padding-top: 1rem; padding-bottom: 2rem; background-color: #ffffff;}
    [data-testid="stSidebar"] {background-color: #f1f5f9 !important;}
    .phase-box {border: 2px solid #e2e8f0; padding: 25px; border-radius: 12px; background-color: #f8fafc; margin-bottom: 30px;}
    .stButton>button {border-radius: 8px; font-weight: 800; font-size: 1.2rem; padding: 0.75rem 0rem;}
    </style>
"""
st.markdown(big_ui_style, unsafe_allow_html=True)

# Lập ngôn ngữ hiển thị động
with st.sidebar:
    sel_lang = st.selectbox("🌐 LANGUAGE / NGÔN NGỮ:", list(LANG_DICT.keys()))
    t = LANG_DICT[sel_lang]

# Đăng nhập
def check_password():
    if "auth" not in st.session_state: st.session_state["auth"] = False
    if st.session_state["auth"]: return True
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown(f"<h1 style='text-align: center; padding-top: 50px;'>{t['gate_title']}</h1>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; font-size:1.1em; color:#555;'>{t['gate_sub']}</p>", unsafe_allow_html=True)
        pwd = st.text_input(t['pwd_label'], type="password")
        if pwd:
            if pwd == st.secrets.get("app_password", "ChingLuh@2026"):
                st.session_state["auth"] = True
                st.rerun()
            else: st.error(t['pwd_err'])
    return False

if not check_password(): st.stop()

@st.cache_resource
def load_ocr_reader(): return easyocr.Reader(['vi', 'en'], gpu=False)
reader = load_ocr_reader()

if 'data_p1' not in st.session_state: st.session_state.data_p1 = pd.DataFrame()

# ==========================================
# 3. LÕI THUẬT TOÁN ĐỌC VÀ KHỬ NHIỄU DỮ LIỆU
# ==========================================
class DataEngine:
    @staticmethod
    def clean_numeric(series):
        if series is None: return 0
        cleaned = series.astype(str).str.replace(',', '', regex=False)
        cleaned = cleaned.str.replace(r'[^\d\.]', '', regex=True)
        return pd.to_numeric(cleaned, errors='coerce').fillna(0)

    @staticmethod
    def purify_code(val):
        if pd.isna(val): return ""
        s = str(val).strip().split('.')[0]
        return re.sub(r'[^A-Z0-9]', '', s.upper())

    @staticmethod
    def fz_match(code, masters, master_purified_dict):
        p_code = DataEngine.purify_code(code)
        if p_code in master_purified_dict: return master_purified_dict[p_code]
        if code in masters: return code
        m = difflib.get_close_matches(code, masters, n=1, cutoff=0.85)
        return m[0] if m else code

    @staticmethod
    def get_col(df, kws, fb_idx=0):
        for c in df.columns:
            for kw in kws:
                if kw.lower() in str(c).lower(): return c
        return df.columns[fb_idx] if len(df.columns) > fb_idx else None

    @staticmethod
    def read_smart(uploaded_file, target_keywords):
        if uploaded_file is None: return pd.DataFrame()
        ext = uploaded_file.name.split('.')[-1].lower()
        try:
            if ext in ['xlsx', 'xls', 'csv']:
                df_raw = pd.read_csv(uploaded_file, skiprows=0, on_bad_lines='skip') if ext == 'csv' else pd.read_excel(uploaded_file, skiprows=0)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                header_row_idx = None
                for i, row in df_raw.head(30).iterrows():
                    combined_str = " ".join(row.dropna().astype(str).lower())
                    if any(kw.lower() in combined_str for kw in target_keywords):
                        header_row_idx = i; break
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
        except Exception as e: st.error(f"Error parsing: {e}")
        return pd.DataFrame()

    @staticmethod
    def parse_ecus(raw_df, master_codes, prefix):
        if raw_df.empty: return pd.DataFrame(columns=['Ma_Vat_Tu', f'SL_{prefix}', 'HS_Code'])
        d_col = DataEngine.get_col(raw_df, ['Mô tả', 'Tên hàng'], 1)
        q_col = DataEngine.get_col(raw_df, ['Số lượng', 'Lượng tính thuế'], 3)
        h_col = DataEngine.get_col(raw_df, ['Mã số', 'HS'], 0)
        tmp = raw_df[[d_col, q_col, h_col]].dropna(subset=[d_col]).copy()
        codes = []
        for _, r in tmp.iterrows():
            desc = str(r[d_col]).upper()
            found = "UNK"
            for c in master_codes:
                if p_code := DataEngine.purify_code(c):
                    if p_code in re.sub(r'[^A-Z0-9]', '', desc):
                        found = c; break
            codes.append(found)
        tmp['Ma_Vat_Tu'] = codes
        res = tmp[tmp['Ma_Vat_Tu'] != 'UNK'].copy()
        res[f'SL_{prefix}'] = DataEngine.clean_numeric(res[q_col])
        res.rename(columns={h_col: 'HS_Code'}, inplace=True)
        return res.groupby(['Ma_Vat_Tu', 'HS_Code'], as_index=False)[f'SL_{prefix}'].sum()

class ExportEngine:
    @staticmethod
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Audit_Result')
            wb = writer.book
            ws = writer.sheets['Audit_Result']
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            fmt_red = wb.add_format({'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'border': 1})
            fmt_green = wb.add_format({'bg_color': '#D1FAE5', 'font_color': '#065F46', 'border': 1})
            
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)
                max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 5 # Tăng thêm padding cho cột
                ws.set_column(col_idx, col_idx, max_len)
            
            status_col_idx = len(df.columns) - 1
            for row_idx, status in enumerate(df.iloc[:, status_col_idx]):
                fmt = fmt_green if "🟢" in str(status) else fmt_red
                ws.write(row_idx + 1, status_col_idx, str(status), fmt)
        return output.getvalue()

    @staticmethod
    def to_word(df, txt_dict):
        doc = Document()
        doc.add_heading(txt_dict["btn_docx"].replace("📝 ", ""), 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Date/Time: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        status_col_idx = len(df.columns) - 1
        err_df = df[~df.iloc[:, status_col_idx].str.contains("🟢", regex=True)]
        if err_df.empty:
            doc.add_paragraph("All data matched successfully. No discrepancies found.")
        else:
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Material', 'Invoice Qty', 'Status'
            for _, r in err_df.head(50).iterrows():
                row = table.add_row().cells
                row[0].text, row[1].text, row[2].text = str(r.iloc[0]), str(r.iloc[2]), str(r.iloc[-1])
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()

# ==========================================
# 4. SIDEBAR CONFIG
# ==========================================
with st.sidebar:
    st.markdown(t["config_side"])
    tol_weight = st.slider(t["tol_w"], min_value=0.0, max_value=1.0, value=0.2, step=0.05)
    tol_count = st.slider(t["tol_c"], min_value=0.0, max_value=0.1, value=0.01, step=0.01)
    
    st.markdown(t["config_skip"])
    skip_inv = st.number_input(t["skip_inv"], value=15)
    skip_cd  = st.number_input(t["skip_cd"], value=17)
    skip_erp = st.number_input(t["skip_erp"], value=0)
    skip_ecus = st.number_input(t["skip_ecus"], value=7)
    
    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["auth"] = False
        st.rerun()

# ==========================================
# 5. GIAO DIỆN CHÍNH (UI TO, RÕ RÀNG)
# ==========================================
st.markdown(f"<h1>{t['main_title']}</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='font-size: 1.15em; color: #475569;'>{t['main_sub']}</p>", unsafe_allow_html=True)
st.markdown("---")

# ----------------- PHASE 0 -----------------
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown(t["phase0_title"])
st.caption(t["phase0_sub"])
c0_1, c0_2 = st.columns(2)
with c0_1:
    f_pre_inv = st.file_uploader(t["pre_inv_lbl"], type=["xlsx","csv","pdf"])
with c0_2:
    f_pre_pkl = st.file_uploader(t["pre_pkl_lbl"], type=["xlsx","csv","pdf"])

if f_pre_inv or f_pre_pkl:
    df_temp_i = DataEngine.read_smart(f_pre_inv, ['Material code', 'Material', 'Mã'])
    df_temp_p = DataEngine.read_smart(f_pre_pkl, ['Material code', 'Material', 'Mã'])
    raw_codes = []
    for d_node in [df_temp_i, df_temp_p]:
        if not d_node.empty:
            col_mat = DataEngine.get_col(d_node, ['Material code', 'Material', 'Mã'], None)
            if col_mat: raw_codes.extend(d_node[col_mat].dropna().astype(str).tolist())
    unique_r_codes = sorted(list(set([DataEngine.purify_code(c) for c in raw_codes if c and str(c).lower() != "nan"])))
    if unique_r_codes:
        st.success(t["phase0_succ"].format(len(unique_r_codes)))
        st.text_area(t["phase0_copy"], value="\n".join(unique_r_codes), height=150)
st.markdown("</div>", unsafe_allow_html=True)

# ----------------- PHASE 1 -----------------
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown(t["phase1_title"])
st.caption(t["phase1_sub"])
c1_1, c1_2 = st.columns(2)
with c1_1:
    f_inv_real = st.file_uploader(t["f_inv_lbl"], type=["xlsx","csv"])
    f_pkl_real = st.file_uploader(t["f_pkl_lbl"], type=["xlsx","csv"])
with c1_2:
    f_cd_real = st.file_uploader(t["f_cd_lbl"], type=["xlsx","csv"])
    f_sap_zmm = st.file_uploader(t["f_sap_lbl"], type=["xlsx","csv"])

if f_inv_real and f_sap_zmm:
    r_inv = DataEngine.read_smart(f_inv_real, ['Material code', 'Material', 'Quantity Customs'])
    r_pkl = DataEngine.read_smart(f_pkl_real, ['Material', 'Quantity', "Q'TY"])
    r_cd  = DataEngine.read_smart(f_cd_real, ['Mã NL', 'Material', 'Số Lượng'])
    r_sap = DataEngine.read_smart(f_sap_zmm, ['PO Number', 'Material', 'PO Qty'])

    c_inv_mat = DataEngine.get_col(r_inv, ['Material code', 'Material', 'Mã'], 1)
    df_inv = r_inv[[c_inv_mat, DataEngine.get_col(r_inv, ['Quantity', 'Số lượng'], 6), DataEngine.get_col(r_inv, ['Unit', 'ĐVT'], 5), DataEngine.get_col(r_inv, ['Price', 'Đơn giá'], 7), DataEngine.get_col(r_inv, ['Amount', 'Thành tiền'], 8)]].dropna(subset=[c_inv_mat]).copy()
    df_inv.columns = ['Ma_Vat_Tu', 'SL_INV', 'UOM_INV', 'DonGia', 'TriGia']
    df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].apply(DataEngine.purify_code)
    for c in ['SL_INV', 'DonGia', 'TriGia']: df_inv[c] = DataEngine.clean_numeric(df_inv[c])
    df_inv = df_inv.groupby(['Ma_Vat_Tu', 'UOM_INV', 'DonGia'], as_index=False)[['SL_INV', 'TriGia']].sum()
    
    masters = df_inv['Ma_Vat_Tu'].tolist()
    master_purified_dict = {DataEngine.purify_code(m): m for m in masters}

    p_mat = DataEngine.get_col(r_pkl, ['Material', 'Mã'], 0)
    p_qty = DataEngine.get_col(r_pkl, ["Q'TY", 'Quantity'], 5)
    df_pkl = pd.DataFrame()
    df_pkl['Ma_Vat_Tu'] = r_pkl[p_mat].astype(str).str.strip().upper().apply(lambda x: DataEngine.fz_match(x, masters, master_purified_dict))
    df_pkl['SL_PKL'] = DataEngine.clean_numeric(r_pkl[p_qty])
    df_pkl = df_pkl.groupby('Ma_Vat_Tu', as_index=False)['SL_PKL'].sum()

    df_cd = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ChiDinh'])
    if not r_cd.empty:
        cd_mat = DataEngine.get_col(r_cd, ['Mã NL', 'Material'], 1)
        cd_qty = DataEngine.get_col(r_cd, ['Số Lượng', 'Quantity'], 4)
        df_cd = pd.DataFrame()
        df_cd['Ma_Vat_Tu'] = r_cd[cd_mat].astype(str).str.strip().upper().apply(lambda x: DataEngine.fz_match(x, masters, master_purified_dict))
        df_cd['SL_ChiDinh'] = DataEngine.clean_numeric(r_cd[cd_qty])
        df_cd = df_cd.groupby('Ma_Vat_Tu', as_index=False)['SL_ChiDinh'].sum()

    c_sap_mat = DataEngine.get_col(r_sap, ['Material'], 10)
    c_sap_qty = DataEngine.get_col(r_sap, ['PO NL QTY', 'In NL QTY', 'In Qty'], 13)
    c_sap_nl = DataEngine.get_col(r_sap, ['NL#'], 22)
    c_sap_desc = DataEngine.get_col(r_sap, ['CUSTOMS DESC'], 23)
    df_sap_raw = r_sap[[c_sap_mat, c_sap_qty, c_sap_nl, c_sap_desc]].dropna(subset=[c_sap_mat]).copy()
    df_sap_raw.columns = ['Ma_Vat_Tu', 'SL_SAP', 'Ma_NL_HQ', 'Customs_Desc']
    df_sap_raw['Ma_Vat_Tu'] = df_sap_raw['Ma_Vat_Tu'].apply(DataEngine.purify_code)
    df_sap_raw['SL_SAP'] = DataEngine.clean_numeric(df_sap_raw['SL_SAP'])
    df_sap = df_sap_raw.groupby(['Ma_Vat_Tu', 'Ma_NL_HQ', 'Customs_Desc'], as_index=False)['SL_SAP'].sum()

    mg_p1 = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
    mg_p1 = pd.merge(mg_p1, df_cd, on='Ma_Vat_Tu', how='outer')
    mg_p1 = pd.merge(mg_p1, df_sap, on='Ma_Vat_Tu', how='outer').fillna(0)

    mg_p1['Lệch_PKL'] = mg_p1['SL_PKL'] - mg_p1['SL_INV']
    mg_p1['Lệch_SAP'] = mg_p1['SL_SAP'] - mg_p1['SL_INV']

    def eval_p1(r):
        if r['SL_INV'] == 0: return "❌ LỖI: THIẾU TRÊN INVOICE"
        tol = tol_weight if r['UOM_INV'] in ['KGM', 'MTK'] else tol_count
        if abs(r['Lệch_PKL']) > tol: return "🔴 LỆCH ĐÓNG GÓI PKL"
        if abs(r['Lệch_SAP']) > tol: return "🚨 LỖI LỆCH KHO SAP"
        return "🟢 KHỚP"

    mg_p1['KẾT LUẬN PHASE 1'] = mg_p1.apply(eval_p1, axis=1)
    st.markdown(t["tbl1_title"])
    st.dataframe(mg_p1, use_container_width=True, hide_index=True, height=500)
    st.session_state.data_p1 = mg_p1
st.markdown("</div>", unsafe_allow_html=True)

# ----------------- PHASE 2 -----------------
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown(t["phase2_title"])
st.caption(t["phase2_sub"])
f_ecus_hang = st.file_uploader(t["f_hq_lbl"], type=["xlsx","csv","xls"])

if f_ecus_hang and not st.session_state.data_p1.empty:
    mg_p2_goc = st.session_state.data_p1.copy()
    r_hq = DataEngine.read_smart(f_ecus_hang, ['Mã số hàng hóa', 'Mô tả hàng hóa', 'Số lượng (1)'])
    
    masters_p2 = mg_p2_goc['Ma_Vat_Tu'].tolist()
    df_hq = DataEngine.parse_ecus(r_hq, masters_p2, "HAI_QUAN")
    mg_final = pd.merge(mg_p2_goc, df_hq, on='Ma_Vat_Tu', how='outer').fillna(0)

    # ĐIỂM CHỐT: TỰ ĐỘNG SINH CHUỖI MÔ TẢ HẢI QUAN
    mg_final['CHUỖI MÔ TẢ HQ (FORM 30+)'] = mg_final['Ma_NL_HQ'].astype(str) + "#&" + mg_final['Customs_Desc'].astype(str)
    mg_final['Lệch_HQ'] = mg_final['SL_HAI_QUAN'] - mg_final['SL_INV']

    def eval_p2(r):
        if "❌" in str(r['KẾT LUẬN PHASE 1']) or "🔴" in str(r['KẾT LUẬN PHASE 1']) or "🚨" in str(r['KẾT LUẬN PHASE 1']):
            return r['KẾT LUẬN PHASE 1']
        tol = tol_weight if r['UOM_INV'] in ['KGM', 'MTK'] else tol_count
        if abs(r['Lệch_HQ']) > tol: return "🔴 LỆCH TỜ KHAI HQ"
        return "🟢 KHỚP HOÀN TOÀN"

    mg_final['TRẠNG THÁI CUỐI CÙNG'] = mg_final.apply(eval_p2, axis=1)
    mg_final = mg_final.sort_values(by='TRẠNG THÁI CUỐI CÙNG', ascending=False)

    st.markdown(t["tbl2_title"])
    err_filter = st.toggle(t["toggle_err"], value=False)
    board_display = mg_final[~mg_final['TRẠNG THÁI CUỐI CÙNG'].str.contains("🟢", regex=False)] if err_filter else mg_final
    
    # Bảng dữ liệu to, cao, rõ nét
    st.data_editor(
        board_display.style.applymap(lambda x: 'background-color:#d1fae5; color:#065f46; font-weight:bold' if '🟢' in str(x) else 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🔴' in str(x) or '❌' in str(x) or '🚨' in str(x) else '', subset=['TRẠNG THÁI CUỐI CÙNG']),
        use_container_width=True, hide_index=True, height=600
    )

    d1, d2 = st.columns(2)
    with d1: st.download_button(t["btn_xlsx"], ExportEngine.to_excel(mg_final), "Audit_Report_Final.xlsx", "primary", use_container_width=True)
    with d2: st.download_button(t["btn_docx"], ExportEngine.to_word(mg_final, t), "Audit_Memo_Final.docx", use_container_width=True)

elif f_ecus_hang and st.session_state.data_p1.empty:
    st.error(t["err_miss"])
st.markdown("</div>", unsafe_allow_html=True)
