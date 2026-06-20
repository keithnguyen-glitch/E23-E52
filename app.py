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
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. TỪ ĐIỂN DỊCH THUẬT ĐA QUỐC GIA (VI - EN - ZH)
# ==========================================
LANG_DICT = {
    "🇻🇳 Tiếng Việt": {
        "gate_title": "🌐 CỔNG ĐIỀU HÀNH CHUỖI CUNG ỨNG",
        "gate_sub": "Hệ thống liên thông dữ liệu Giao - Nhận Ching Luh & Fu-Luh",
        "pwd_label": "Nhập mã truy cập an ninh:",
        "pwd_err": "❌ Mã truy cập không hợp lệ!",
        "main_title": "💎 TRUNG TÂM KIỂM ĐỊNH LIÊN THÔNG CHỨNG TỪ XNK",
        "main_sub": "Hỗ trợ tự động hóa khép kín: Tách mã liệu R ➔ Quét sổ kho SAP ZMM12 ➔ Ráp chuỗi Hải quan #& ➔ Xuất form tờ khai đầu 30.",
        "config_side": "### ⚙️ CẤU HÌNH HỆ THỐNG",
        "tol_w": "Dung sai hàng Vải/Cân ký (KGM, MTK):",
        "tol_c": "Dung sai hàng Đếm chiếc (PCE, PRS, PR):",
        "skip_inv": "Dòng thừa Invoice/PKL:",
        "skip_cd": "Dòng thừa Chỉ Định:",
        "skip_erp": "Dòng thừa SAP (ZMM12):",
        "skip_ecus": "Dòng thừa Tờ Khai (Tab HANG):",
        "phase0_title": "### ⚡ PHASE 0: TRÍCH XUẤT MÃ LIỆU R TIỀN TRẠM (SAP QUERY HELPER)",
        "phase0_sub": "Nhận chứng từ thô từ kho gửi ➔ Thảy vào đây để lấy danh sách mã sạch đem đi truy vấn SAP ZMM12.",
        "pre_inv_lbl": "Tải lên Invoice nhận từ kho gửi:",
        "pre_pkl_lbl": "Tải lên Packing List nhận từ kho gửi:",
        "phase0_succ": "🎉 Hệ thống đã gom và lọc sạch trùng lặp! Trích xuất được {} Mã liệu R độc nhất.",
        "phase0_copy": "👇 Bấm vào ô dưới, nhấn Ctrl+A rồi Ctrl+C để copy danh sách đem dán thẳng vào SAP ZMM12:",
        "phase1_title": "### 📥 PHASE 1 & 2: ĐỐI CHIẾU DIỆN RỘNG VÀ TỰ ĐỘNG GENERATE CHUỖI HẢI QUAN",
        "phase1_sub": "Kéo thả tất cả các file sau khi đã lấy được báo cáo SAP ZMM12 và Tờ khai thô về máy.",
        "grp_comm": "📄 CHỨNG TỪ THƯƠNG MẠI GỐC",
        "grp_sap": "💻 SỔ SÁCH KHO NỘI BỘ (SAP)",
        "grp_ecus": "🏛️ TỜ KHAI HẢI QUAN (ECUS)",
        "f_inv_lbl": "Invoice Khai Báo Chính Thức:",
        "f_pkl_lbl": "Packing List Khai Báo Chính Thức:",
        "f_cd_lbl": "Chỉ Định Giao Hàng (Nếu có):",
        "f_sap_lbl": "Báo Cáo SAP ZMM12 vừa tải về:",
        "f_hq_lbl": "Tờ khai Hải quan (Tab HANG):",
        "btn_run": "🚀 KÍCH HOẠT SIÊU HỆ THỐNG ĐỐI CHIẾU TOÀN DIỆN",
        "err_miss": "⚠️ Yêu cầu tối thiểu phải có Invoice chính thức và báo cáo SAP ZMM12.",
        "tbl_title": "### 📊 BẢNG KẾT QUẢ ĐỐI CHIẾU LIÊN THÔNG TOÀN CỤC",
        "toggle_err": "🚨 BẬT BỘ LỌC LỖI (Ẩn toàn bộ dòng đúng để tập trung sửa dòng sai)",
        "tax_msg": "💸 Dự toán Thuế lô hàng Nhập Khẩu: Khoảng ${:,.2f} USD dựa trên Biểu thuế.",
        "btn_xlsx": "📊 TẢI EXCEL ĐỐI CHIẾU CHÉO (Có Summary & Khối màu)",
        "btn_docx": "📝 TẢI BIÊN BẢN WORD KÝ DUYỆT TRÌNH GIÁM ĐỐC",
        "master_title": "⚙️ NẠP DỮ LIỆU NỀN QUY ĐỔI & BIỂU THUẾ TRUNG TÂM"
    },
    "🇺🇸 English": {
        "gate_title": "🌐 SUPPLY CHAIN COMMAND CENTER",
        "gate_sub": "Interlinked Data System for Ching Luh & Fu-Luh Inbound/Outbound",
        "pwd_label": "Enter Security Access Code:",
        "pwd_err": "❌ Invalid Access Code!",
        "main_title": "💎 GLOBAL EXIM DATA CROSS-CHECK PLATFORM",
        "main_sub": "Closed-loop Automation: Extract Material Code R ➔ Audit SAP ZMM12 Stock ➔ Auto-assemble Customs #& ➔ Export Form 30 Declaration.",
        "config_side": "### ⚙️ SYSTEM CONFIGURATION",
        "tol_w": "Tolerance for Fabric/Weight (KGM, MTK):",
        "tol_c": "Tolerance for Countable Items (PCE, PRS, PR):",
        "skip_inv": "Header Rows Invoice/PKL:",
        "skip_cd": "Header Rows Shipping Instruction:",
        "skip_erp": "Header Rows SAP (ZMM12):",
        "skip_ecus": "Header Rows Customs Form (HANG):",
        "phase0_title": "### ⚡ PHASE 0: PRE-ROUTE MATERIAL CODE EXTRACTION (SAP QUERY HELPER)",
        "phase0_sub": "Upload raw documents from warehouse ➔ Extract sanitized code list to query SAP ZMM12.",
        "pre_inv_lbl": "Upload Invoice from Warehouse:",
        "pre_pkl_lbl": "Upload Packing List from Warehouse:",
        "phase0_succ": "🎉 Duplicates removed successfully! Extracted {} unique Material R codes.",
        "phase0_copy": "👇 Click below, press Ctrl+A then Ctrl+C to copy and paste directly into SAP ZMM12 Multi-Selection:",
        "phase1_title": "### 📥 PHASE 1 & 2: ENTERPRISE AUDITING & CUSTOMS STRING AUTO-GENERATION",
        "phase1_sub": "Drag and drop all files once you have downloaded the SAP ZMM12 report and raw Customs files.",
        "grp_comm": "📄 ORIGINAL COMMERCIAL DOCUMENTS",
        "grp_sap": "💻 INTERNAL SYSTEM RECORDS (SAP)",
        "grp_ecus": "🏛️ CUSTOMS DECLARATION (ECUS)",
        "f_inv_lbl": "Official Customs Invoice:",
        "f_pkl_lbl": "Official Packing List:",
        "f_cd_lbl": "Shipping Instruction (Optional):",
        "f_sap_lbl": "Downloaded SAP ZMM12 Report:",
        "f_hq_lbl": "Customs Declaration (Tab HANG):",
        "btn_run": "🚀 ACTIVATE COMPREHENSIVE RECONCILIATION ENGINE",
        "err_miss": "⚠️ Minimum requirement: Official Invoice and SAP ZMM12 report must be uploaded.",
        "tbl_title": "### 📊 GLOBAL COMPLIANCE RECONCILIATION MATRIX",
        "toggle_err": "🚨 ENABLE ERROR FILTER (Hide matched rows to focus on discrepancy fixes)",
        "tax_msg": "💸 Estimated Inbound Customs Duty: Approx ${:,.2f} USD based on Tariff Table.",
        "btn_xlsx": "📊 DOWNLOAD COLOR-CODED EXCEL AUDIT REPORT",
        "btn_docx": "📝 DOWNLOAD WORD MEMORANDUM FOR DIRECTOR SIGN OFF",
        "master_title": "⚙️ CENTRAL MASTER DATA CONVERSION & TARIFF MANAGEMENT"
    },
    "🇨🇳 中文": {
        "gate_title": "🌐 供应链数据管理中心",
        "gate_sub": "清禄与福禄工厂出入库数据互联系统",
        "pwd_label": "请输入安全访问密码:",
        "pwd_err": "❌ 访问密码错误!",
        "main_title": "💎 全球进出口报关数据智能核对平台",
        "main_sub": "闭环自动化：提取物料R码 ➔ 审计 SAP ZMM12 库存 ➔ 自动拼装海关 #& 字符串 ➔ 导出海关表30报关单。",
        "config_side": "### ⚙️ 系统参数配置",
        "tol_w": "面料/称重类货物容差 (KGM, MTK):",
        "tol_c": "数量类货物容差 (PCE, PRS, PR):",
        "skip_inv": "发票/装箱单标题冗余行:",
        "skip_cd": "出货通知书标题冗余行:",
        "skip_erp": "SAP (ZMM12) 报表冗余行:",
        "skip_ecus": "海关报关单 (HANG) 标题冗余行:",
        "phase0_title": "### ⚡ PHASE 0: 发货仓原始单据物料号提取 (SAP 查询助手)",
        "phase0_sub": "接收发货仓原始单据 ➔ 放入此模块提取去重物料号以查询 SAP ZMM12。",
        "pre_inv_lbl": "上传发货仓原始发票:",
        "pre_pkl_lbl": "上传发货仓原始装箱单:",
        "phase0_succ": "🎉 成功去重！精确提取到 {} 个唯一物料 R 码。",
        "phase0_copy": "👇 点击下方文本框，按 Ctrl+A 再 Ctrl+C 复制列表，直接粘贴至 SAP ZMM12 多项选择框：",
        "phase1_title": "### 📥 PHASE 1 & 2: 全局数据穿透核对与海关申报品名自动生成",
        "phase1_sub": "获取 SAP ZMM12 报表与海关原始文件后，将所有文件拖放至下方进行比对。",
        "grp_comm": "📄 原始商业单证",
        "grp_sap": "💻 内部仓储系统记录 (SAP)",
        "grp_ecus": "🏛️ 海关报关单记录 (ECUS)",
        "f_inv_lbl": "官方申报发票 (Invoice):",
        "f_pkl_lbl": "官方申报装箱单 (Packing List):",
        "f_cd_lbl": "出货通知书 (选填):",
        "f_sap_lbl": "导出的 SAP ZMM12 库存报表:",
        "f_hq_lbl": "海关报关单文件 (Tab HANG):",
        "btn_run": "🚀 启动全局供应链全维核对引擎",
        "err_miss": "⚠️ 核心控制点要求：必须至少上传官方发票与 SAP ZMM12 报表。",
        "tbl_title": "### 📊 全局合规对账数据矩阵",
        "toggle_err": "🚨 开启错误过滤 (隐藏完全匹配行，专注修改差异行)",
        "tax_msg": "💸 进口关税预估：基于税率表预计约为 ${:,.2f} USD。",
        "btn_xlsx": "📊 下载带颜色格式的 Excel 审计报告",
        "btn_docx": "📝 下载呈批阅Word备忘录 (供总经理签字)",
        "master_title": "⚙️ 中央主数据转换与税率综合管理"
    }
}

# ==========================================
# 2. CONFIG SIDEBAR & LIGHT/DARK THEME ENGINE
# ==========================================
# Thêm điều khiển ngôn ngữ và theme lên thanh Sidebar đầu tiên
with st.sidebar:
    st.markdown("### 🌐 LANGUAGE & THEME")
    sel_lang = st.selectbox("Chọn ngôn ngữ / Select Language / 选择语言:", list(LANG_DICT.keys()))
    t = LANG_DICT[sel_lang] # Object ngôn ngữ hiện tại
    
    # Tính năng Cải tiến: Thiết lập Light / Dark Mode dynamic qua CSS injection
    theme_selection = st.radio("Chế độ hiển thị / Mode:", ["☀️ Light Mode", "🌙 Dark Mode"])

# Cấu hình màu sắc CSS động theo lựa chọn của người dùng
if theme_selection == "🌙 Dark Mode":
    bg_color, text_color, card_bg, border_color, sidebar_bg = "#0f172a", "#f8fafc", "#1e293b", "#334155", "#1e293b"
else:
    bg_color, text_color, card_bg, border_color, sidebar_bg = "#ffffff", "#0f172a", "#f8fafc", "#e2e8f0", "#f8fafc"

hide_st_style = f"""
            <style>
            #MainMenu {{visibility: hidden;}}
            footer {{visibility: hidden;}}
            header {{visibility: hidden;}}
            .stDeployButton {{display:none;}}
            .block-container {{padding-top: 1rem; padding-bottom: 2rem; background-color: {bg_color}; color: {text_color};}}
            
            /* Theme overrides */
            .stApp {{background-color: {bg_color}; color: {text_color};}}
            [data-testid="stSidebar"] {{background-color: {sidebar_bg} !important;}}
            
            .metric-card {{background: {card_bg}; padding: 20px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); border: 1px solid {border_color}; border-left: 5px solid #2563eb; margin-bottom: 15px; color: {text_color};}}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# Lớp mật khẩu gate bảo vệ
def check_password():
    if "auth" not in st.session_state: st.session_state["auth"] = False
    if not st.session_state["auth"]:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.markdown(f"<h2 style='text-align: center; color: #1E3A8A; padding-top: 50px;'>{t['gate_title']}</h2>", unsafe_allow_html=True)
            st.markdown(f"<p style='text-align: center;'>{t['gate_sub']}</p>", unsafe_allow_html=True)
            pwd = st.text_input(t['pwd_label'], type="password")
            if pwd:
                if pwd == st.secrets.get("app_password", "ChingLuh@2026"):
                    st.session_state["auth"] = True
                    st.rerun()
                else: st.error(t['pwd_err'])
        return False
    return True

if not check_password(): st.stop()

@st.cache_resource
def load_ocr_reader(): return easyocr.Reader(['vi', 'en'], gpu=False)
reader = load_ocr_reader()

# ==========================================
# 3. LÕI THUẬT TOÁN ĐỌC VÀ LÀM SẠCH DỮ LIỆU
# ==========================================
def clean_numeric_string(series):
    if series is None: return 0
    cleaned = series.astype(str).str.replace(',', '', regex=False)
    cleaned = cleaned.str.replace(r'[^\d\.]', '', regex=True)
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

def purify_code(s):
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
                                cr = [str(c).replace('\n', ' ') if c else '' for c in r]
                                if any(cr): all_text_rows.append(cr)
            if all_text_rows: return pd.DataFrame(all_text_rows[1:], columns=[str(c) for c in all_text_rows[0]])
    except Exception as e: st.error(f"Error {uploaded_file.name}: {e}")
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
            if c in desc: found = c; break
        codes.append(found)
    tmp['Ma_Vat_Tu'] = codes
    res = tmp[tmp['Ma_Vat_Tu'] != 'UNK'].copy()
    res[f'SL_{prefix}'] = clean_numeric_string(res[q_col])
    res.rename(columns={h_col: 'HS_Code'}, inplace=True)
    return res.groupby(['Ma_Vat_Tu', 'HS_Code'], as_index=False)[f'SL_{prefix}'].sum()

def apply_hsqd(df, qty_col, uom_col):
    hsqd = st.session_state.master_hsqd
    if not hsqd.empty and not df.empty:
        df = pd.merge(df, hsqd, on='Ma_Vat_Tu', how='left')
        df['He_So_QD'] = df['He_So_QD'].fillna(1.0)
        df[qty_col] = df[qty_col] * df['He_So_QD']
        if uom_col in df.columns: df[uom_col] = df['DVT_Bao_Quan'].combine_first(df[uom_col])
        return df.drop(columns=['He_So_QD', 'DVT_Bao_Quan'], errors='ignore')
    return df

# =====================================================================
# 4. EXPORT ENGINE CƯỜNG HÓA 
# =====================================================================
class ExportEngine:
    @staticmethod
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Cross_Check_Result')
            wb = writer.book
            ws = writer.sheets['Cross_Check_Result']
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
            ws.set_column(status_col_idx, status_col_idx, 40)
            for row_idx, status in enumerate(df.iloc[:, status_col_idx]):
                fmt = fmt_green if "🟢" in str(status) or "🟡" in str(status) else fmt_red
                ws.write(row_idx + 1, status_col_idx, str(status), fmt)
        return output.getvalue()

    @staticmethod
    def to_word(df, phrase_dict):
        doc = Document()
        doc.add_heading('BIÊN BẢN ĐỐI CHIẾU SỐ LIỆU CHUỖI CUNG ỨNG', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Date/Time: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        
        doc.add_heading('Discrepancy items flagged by system:', level=1)
        status_col_idx = len(df.columns) - 1
        err_df = df[~df.iloc[:, status_col_idx].str.contains("🟢|🟡", regex=True)]
        if err_df.empty:
            doc.add_paragraph("All matched successfully.")
        else:
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Material', 'Invoice Qty', 'Status'
            for _, r in err_df.head(25).iterrows():
                row = table.add_row().cells
                row[0].text, row[1].text, row[2].text = str(r.iloc[0]), str(r.iloc[2]), str(r.iloc[-1])
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()

# ==========================================
# 5. SIDEBAR DYNAMIC INPUTS
# ==========================================
with st.sidebar:
    st.markdown(t["config_side"])
    tol_weight = st.slider(t["tol_w"], min_value=0.0, max_value=1.0, value=0.2, step=0.05)
    tol_count = st.slider(t["tol_c"], min_value=0.0, max_value=0.1, value=0.01, step=0.01)
    st.markdown("---")
    skip_inv = st.number_input(t["skip_inv"], value=15)
    skip_cd  = st.number_input(t["skip_cd"], value=17)
    skip_erp = st.number_input(t["skip_erp"], value=0)
    skip_ecus = st.number_input(t["skip_ecus"], value=7)

# ==========================================
# 6. THÂN GIAO DIỆN CHÍNH (MAIN DATA APP)
# ==========================================
st.markdown(f"<h1 style='color: #1E3A8A;'>{t['main_title']}</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='font-size: 1.05em;'>{t['main_sub']}</p>", unsafe_allow_html=True)
st.markdown("---")

with st.expander(t["master_title"]):
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        f_hsqd = st.file_uploader("Upload HSQĐ:", type=["xlsx","csv"])
        if f_hsqd:
            df = universal_file_extractor(f_hsqd, 1)
            if not df.empty:
                df = df[[get_col(df, ['Mã'], 1), get_col(df, ['Hệ số'], 6), get_col(df, ['Đơn vị'], 8)]].dropna()
                df.columns = ['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan']
                df['Ma_Vat_Tu'] = df['Ma_Vat_Tu'].apply(purify_code)
                df['He_So_QD'] = pd.to_numeric(df['He_So_QD'], errors='coerce').fillna(1.0)
                st.session_state.master_hsqd = df
                st.success("✅ Synced.")
    with m_col2:
        f_thue = st.file_uploader("Upload Tariff:", type=["xlsx","csv"])
        if f_thue:
            df = universal_file_extractor(f_thue, 2)
            if not df.empty:
                df = df[[get_col(df, ['Mã hàng', 'HS'], 1), get_col(df, ['Thuế suất'], 5)]].dropna()
                df.columns = ['HS_Code', 'Thue_Suat']
                df['HS_Code'] = df['HS_Code'].astype(str).str.replace('.', '', regex=False).str.strip()
                st.session_state.master_thue = df
                st.success("✅ Synced.")

st.markdown(t["phase0_title"])
st.caption(t["phase0_sub"])
col_f1, col_f2 = st.columns(2)
with col_f1: f_pre_inv = st.file_uploader(t["pre_inv_lbl"], type=["xlsx","csv","pdf"])
with col_f2: f_pre_pkl = st.file_uploader(t["pre_pkl_lbl"], type=["xlsx","csv","pdf"])

if f_pre_inv or f_pre_pkl:
    with st.spinner("Processing..."):
        df_temp_i = universal_file_extractor(f_pre_inv, skip_inv)
        df_temp_p = universal_file_extractor(f_pre_pkl, skip_inv)
        raw_codes = []
        for df_node in [df_temp_i, df_temp_p]:
            if not df_node.empty:
                col_mat = get_col(df_node, ['Material code', 'Material', 'Mã'], None)
                if col_mat: raw_codes.extend(df_node[col_mat].dropna().astype(str).str.strip().str.upper().tolist())
        unique_r_codes = sorted(list(set([c for c in raw_codes if c and c != "NAN" and len(c) > 3])))
        if unique_r_codes:
            st.success(t["phase0_succ"].format(len(unique_r_codes)))
            st.text_area(t["phase0_copy"], value="\n".join(unique_r_codes), height=100)

st.markdown("---")
st.markdown(t["phase1_title"])
st.caption(t["phase1_sub"])

col1, col2, col3 = st.columns(3)
with col1:
    st.info(t["grp_comm"])
    f_inv_real = st.file_uploader(t["f_inv_lbl"], type=["xlsx","csv"])
    f_pkl_real = st.file_uploader(t["f_pkl_lbl"], type=["xlsx","csv"])
    f_cd_real  = st.file_uploader(t["f_cd_lbl"], type=["xlsx","csv"])
with col2:
    st.warning(t["grp_sap"])
    f_sap_zmm = st.file_uploader(t["f_sap_lbl"], type=["xlsx","csv"])
with col3:
    st.success(t["grp_ecus"])
    f_ecus_hang = st.file_uploader(t["f_hq_lbl"], type=["xlsx","csv","xls"])

if st.button(t["btn_run"], type="primary", use_container_width=True):
    if not (f_inv_real and f_sap_zmm): st.error(t["err_miss"])
    else:
        with st.spinner("Processing reconciliation..."):
            r_inv = universal_file_extractor(f_inv_real, skip_inv)
            r_pkl = universal_file_extractor(f_pkl_real, skip_inv)
            r_cd  = universal_file_extractor(f_cd_real, skip_cd)
            r_sap = universal_file_extractor(f_sap_zmm, skip_erp)
            r_hq  = universal_file_extractor(f_ecus_hang, skip_ecus)

            c_inv_mat = get_col(r_inv, ['Material code', 'Material', 'Mã'], 1)
            df_inv = r_inv[[c_inv_mat, get_col(r_inv, ['Quantity Customs', 'Quantity', 'Số lượng'], 6), get_col(r_inv, ['Unit Customs', 'Unit', 'ĐVT'], 5), get_col(r_inv, ['Unit Price', 'Price'], 7), get_col(r_inv, ['Amount'], 8)]].dropna(subset=[c_inv_mat]).copy()
            df_inv.columns = ['Ma_Vat_Tu', 'SL_INV', 'UOM_INV', 'DonGia', 'TriGia']
            df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].astype(str).str.strip().str.upper()
            for c in ['SL_INV', 'DonGia', 'TriGia']: df_inv[c] = clean_numeric_string(df_inv[c])
            df_inv = df_inv.groupby(['Ma_Vat_Tu', 'UOM_INV', 'DonGia'], as_index=False)[['SL_INV', 'TriGia']].sum()
            
            masters = df_inv['Ma_Vat_Tu'].tolist()
            master_purified_dict = {purify_code(m): m for m in masters}

            p_mat = get_col(r_pkl, ['Material', 'Mã'], 0)
            p_qty = get_col(r_pkl, ["Q'TY", 'Quantity'], 5)
            df_pkl_raw = pd.DataFrame()
            df_pkl_raw['Ma_Vat_Tu'] = r_pkl[p_mat].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
            df_pkl_raw['SL_PKL'] = clean_numeric_string(r_pkl[p_qty])
            df_pkl_raw['UOM_PKL'] = None
            df_pkl = apply_hsqd(df_pkl_raw, 'SL_PKL', 'UOM_PKL').drop(columns=['UOM_PKL'], errors='ignore')
            df_pkl = df_pkl.groupby('Ma_Vat_Tu', as_index=False)['SL_PKL'].sum()

            df_cd = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ChiDinh'])
            if not r_cd.empty:
                cd_mat = get_col(r_cd, ['Mã NL', 'Material'], 1)
                cd_qty = get_col(r_cd, ['Số Lượng', 'Quantity'], 4)
                df_cd = pd.DataFrame()
                df_cd['Ma_Vat_Tu'] = r_cd[cd_mat].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
                df_cd['SL_ChiDinh'] = clean_numeric_string(r_cd[cd_qty])
                df_cd = df_cd.groupby('Ma_Vat_Tu', as_index=False)['SL_ChiDinh'].sum()

            df_sap = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_SAP_ZMM', 'Ma_NL_HQ', 'Customs_Desc'])
            if not r_sap.empty:
                c_sap_mat = get_col(r_sap, ['Material'], 10)
                c_sap_qty = get_col(r_sap, ['PO NL QTY', 'In NL QTY', 'In Qty'], 13)
                c_sap_nl = get_col(r_sap, ['NL#'], 22)
                c_sap_desc = get_col(r_sap, ['CUSTOMS DESC'], 23)
                df_sap_raw = r_sap[[c_sap_mat, c_sap_qty, c_sap_nl, c_sap_desc]].dropna(subset=[c_sap_mat]).copy()
                df_sap_raw.columns = ['Ma_Vat_Tu', 'SL_SAP_ZMM', 'Ma_NL_HQ', 'Customs_Desc']
                df_sap_raw['Ma_Vat_Tu'] = df_sap_raw['Ma_Vat_Tu'].astype(str).str.strip().str.upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
                df_sap_raw['SL_SAP_ZMM'] = clean_numeric_string(df_sap_raw['SL_SAP_ZMM'])
                df_sap = df_sap_raw.groupby(['Ma_Vat_Tu', 'Ma_NL_HQ', 'Customs_Desc'], as_index=False)['SL_SAP_ZMM'].sum()

            df_hq = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_HAI_QUAN'])
            if not r_hq.empty:
                c_hq_mat = get_col(r_hq, ['Mã số'], 0)
                c_hq_qty = get_col(r_hq, ['Số lượng (1)'], 3)
                df_hq = r_hq[[c_hq_mat, c_hq_qty]].dropna().copy()
                df_hq.columns = ['Ma_Vat_Tu', 'SL_HAI_QUAN']
                df_hq['SL_HAI_QUAN'] = clean_numeric_string(df_hq['SL_HAI_QUAN'])
                df_hq = df_hq.groupby('Ma_Vat_Tu', as_index=False)['SL_HAI_QUAN'].sum()

            mg = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_cd, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_sap, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_hq, on='Ma_Vat_Tu', how='outer').fillna(0)

            mg['MÔ TẢ TOÀN DIỆN HẢI QUAN (FORM 30+)'] = mg['Ma_NL_HQ'].astype(str) + "#&" + mg['Customs_Desc'].astype(str)

            def evaluate_row(r):
                if r['SL_INV'] == 0: return "❌ ERROR: THIẾU TRÊN INVOICE GỐC"
                tol = tol_weight if r['UOM_INV'] in ['KGM', 'MTK'] else tol_count
                if abs(r['SL_INV'] - r['SL_PKL']) > tol: return "🔴 LỆCH SỐ LIỆU ĐÓNG GÓI PACKING LIST"
                if 'SL_SAP_ZMM' in r and abs(r['SL_INV'] - r['SL_SAP_ZMM']) > tol: return "🚨 LỖI THẤT THOÁT: VÊNH TỒN KHO NỘI BỘ SAP"
                if 'SL_HAI_QUAN' in r and r['SL_HAI_QUAN'] > 0 and abs(r['SL_INV'] - r['SL_HAI_QUAN']) > tol: return "🔴 LỆCH THỰC TẾ KHAI TỜ KHAI HQ"
                return "🟢 CHỨNG TỪ KHỚP HOÀN TOÀN"

            mg['KẾT LUẬN KIỂM CHÉO'] = mg.apply(evaluate_row, axis=1)
            mg = mg.sort_values(by='KẾT LUẬN KIỂM CHÉO', ascending=False)

            st.markdown("---")
            st.markdown(t["tbl_title"])
            err_filter = st.toggle(t["toggle_err"], value=False)
            final_board = mg[~mg['KẾT LUẬN KIỂM CHÉO'].str.contains("🟢", regex=False)] if err_filter else mg
            
            st.data_editor(
                final_board.style.applymap(lambda x: 'background-color:#d1fae5; color:#065f46' if '🟢' in str(x) else 'background-color:#7f1d1d; color:#fca5a5; font-weight:bold' if '🚨' in str(x) else 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🔴' in str(x) or '❌' in str(x) else '', subset=['KẾT LUẬN KIỂM CHÉO']),
                use_container_width=True, hide_index=True, height=400
            )

            tong_thue = mg['Thue_Du_Kien'].sum() if 'Thue_Du_Kien' in mg.columns else 0
            if tong_thue > 0: st.info(t["tax_msg"].format(tong_thue))

            st.markdown("#### 📥 EXPORT")
            d1, d2 = st.columns(2)
            with d1: st.download_button(t["btn_xlsx"], ExportEngine.to_excel(mg), "Exim_Audit_Report.xlsx", "primary", use_container_width=True)
            with d2: st.download_button(t["btn_docx"], ExportEngine.to_word(mg, t), "Exim_Memorandum.docx", use_container_width=True)
