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
# 1. TỪ ĐIỂN DỊCH THUẬT NỘI BỘ (VI - EN - ZH)
# ==========================================
LANG_DICT = {
    "🇻🇳 Tiếng Việt": {
        "gate_title": "🌐 CLG SCM EXIM VN INTERNALS",
        "gate_sub": "Hệ thống đối soát liên thông chứng từ phục vụ nhóm nghiệp vụ nhập kho",
        "pwd_label": "Nhập mã an ninh nội bộ:",
        "pwd_err": "❌ Mã truy cập không hợp lệ!",
        "main_title": "🏢 PHÂN HỆ KIỂM SOÁT CHỨNG TỪ GIAO NHẬN E23 - E54",
        "main_sub": "Bộ công cụ hỗ trợ tác chiến tuần tự dành cho nhóm nghiệp vụ nhập thuộc CLG SCM EXIM VN (Hỗ trợ đọc Excel, PDF)",
        "config_side": "### ⚖️ THIẾT LẬP THAM SỐ",
        "tol_w": "Dung sai hàng Vải/Cân ký (KGM, MTK):",
        "tol_c": "Dung sai hàng Đếm chiếc (PCE, PRS, PR):",
        "skip_inv": "Dòng thừa Invoice/PKL thô:",
        "skip_cd": "Dòng thừa Chỉ Định:",
        "skip_erp": "Dòng thừa SAP (ZMM12):",
        "skip_ecus": "Dòng thừa Tờ Khai (Tab HANG):",
        "phase0_title": "#### ⚡ PHASE 0: PHÒNG VỆ SỐ LIỆU TỒN KHO GỐC E21 & TRÍCH XUẤT MÃ LIỆU R TIỀN TRẠM",
        "phase0_sub": "Mục tiêu: Đóng băng dữ liệu Nhập khẩu E21 ban đầu làm mốc đối chiếu tối cao, song song trích xuất tập mã R sạch để truy vấn SAP ZMM12.",
        "pre_inv_lbl": "1. Tải lên Invoice kho gửi (Tệp thô):",
        "pre_pkl_lbl": "2. Tải lên Packing List kho gửi (Tệp thô):",
        "pre_e21_lbl": "3. BÁO CÁO TỜ KHAI NHẬP KHẨU E21 GỐC BAN ĐẦU (Quản trị thanh khoản):",
        "pre_hsqd_lbl": "4. Bảng hệ số quy đổi định mức (HSQĐ):",
        "pre_thue_lbl": "5. Bảng biểu thuế (Tra cứu mã HS):",
        "phase0_succ": "📋 Tự động phát hiện {} Mã liệu R độc nhất từ chứng từ thô.",
        "phase0_copy": "👉 Nhấn Ctrl+A rồi Ctrl+C để copy danh sách mã liệu R dán vào mục Multi-Selection trên SAP (ZMM12):",
        "phase1_title": "#### 📦 PHASE 1: ĐỐI SOÁT MA TRẬN GIAO NHẬN THỰC TẾ & KHẤP CHÉO SỔ KHO",
        "phase1_sub": "Hệ thống tự động chạy phép trừ số lượng đa chiều để phát hiện chính xác khối lượng hao hụt.",
        "grp_comm": "📄 CHỨNG TỪ THƯƠNG MẠI GỐC",
        "grp_sap": "💻 SỔ SÁCH KHO NỘI BỘ (SAP)",
        "f_inv_lbl": "1. INVOICE KHAI BÁO CHÍNH THỨC (Hồ sơ gốc):",
        "f_pkl_lbl": "2. PACKING LIST KHAI BÁO CHÍNH THỨC (Hồ sơ gốc):",
        "f_cd_lbl": "3. CHỈ ĐỊNH GIAO HÀNG (Lịch điều xe):",
        "f_sap_lbl": "4. BÁO CÁO KHO SAP ZMM12 VỪA TRUY VẤN VỀ:",
        "tbl1_title": "##### 📊 KẾT QUẢ PHÂN TÍCH CHÊNH LỆCH GIAO NHẬN THỰC TẾ & THANH KHOẢN GỐC",
        "phase2_title": "#### 🏛️ PHASE 2: ĐỒNG BỘ TỜ KHAI HẢI QUAN & TRÍCH XUẤT FILE REPORT",
        "phase2_sub": "Cross-check chiều cuối cùng với dữ liệu tờ khai Hải quan để xuất báo cáo trình ký.",
        "f_hq_lbl": "1. TẢI LÊN FILE DỮ LIỆU TỜ KHAI HẢI QUAN (TAB HANG):",
        "tbl2_title": "##### 📊 BẢNG VÀNG KIỂM ĐỊNH LIÊN THÔNG TOÀN DIỆN DIỆN RỘNG",
        "toggle_err": "🚨 Chỉ lọc hiển thị các dòng phát hiện chênh lệch lỗi số lượng",
        "tax_msg": "💸 Dự toán Thuế lô hàng Nhập Khẩu: Khoảng ${:,.2f} USD dựa trên Biểu thuế.",
        "btn_xlsx": "📊 TẢI EXCEL ĐỐI CHIẾU CHÉO (Có phân loại sheet lỗi)",
        "btn_docx": "📝 TẢI BIÊN BẢN WORD KÝ DUYỆT TRÌNH GIÁM ĐỐC",
        "toast_e21": "✅ ĐÃ ĐÓNG BĂNG BỂ CHỨA TỒN THANH KHOẢN GỐC E21!"
    },
    "🇺🇸 English": {
        "gate_title": "🌐 CLG SCM EXIM VN INTERNALS",
        "gate_sub": "Internal interlinked document reconciliation tool for inbound operations",
        "pwd_label": "Enter Internal Security Code:",
        "pwd_err": "❌ Invalid Access Code!",
        "main_title": "🏢 E23 - E54 DOCUMENT RECONCILIATION MODULE",
        "main_sub": "Sequential data pipeline customized for CLG SCM EXIM VN inbound team (Excel/PDF supported)",
        "config_side": "### ⚙️ PARAMETERS SETTING",
        "tol_w": "Tolerance for Fabric/Weight (KGM, MTK):",
        "tol_c": "Tolerance for Countable Items (PCE, PRS, PR):",
        "skip_inv": "Raw Invoice/PKL Header Rows:",
        "skip_cd": "Shipping Instruction Header Rows:",
        "skip_erp": "SAP ZMM12 Header Rows:",
        "skip_ecus": "Customs Form HANG Header Rows:",
        "phase0_title": "#### ⚡ PHASE 0: E21 INITIAL IMPORT BALANCE SHIELD & PRE-ROUTE MATERIAL CODE EXTRACTION",
        "phase0_sub": "Objective: Lock initial E21 Import data as the supreme baseline, while extracting clean R-code list to query SAP ZMM12.",
        "pre_inv_lbl": "1. Upload Raw Invoice from Warehouse:",
        "pre_pkl_lbl": "2. Upload Raw Packing List from Warehouse:",
        "pre_e21_lbl": "3. INITIAL E21 IMPORT DECLARATION REPORT (Liquidation Management):",
        "pre_hsqd_lbl": "4. UOM Conversion Factor Table (HSQĐ):",
        "pre_thue_lbl": "5. Tariff Duty Table (HS Code lookup):",
        "phase0_succ": "📋 Auto-detected {} unique Material R codes from raw documents.",
        "phase0_copy": "👉 Press Ctrl+A then Ctrl+C to copy code list and paste directly into SAP ZMM12 Multi-Selection window:",
        "phase1_title": "#### 📦 PHASE 1: ACTUAL INBOUND RECONCILIATION & STOCK CROSS-CHECK",
        "phase1_sub": "The system executes multi-dimensional subtraction to isolate exact variance and shortages.",
        "grp_comm": "📄 ORIGINAL COMMERCIAL DOCUMENTS",
        "grp_sap": "💻 INTERNAL SAP RECORDS",
        "f_inv_lbl": "1. OFFICIAL CUSTOMS INVOICE (Master):",
        "f_pkl_lbl": "2. OFFICIAL PACKING LIST (Master):",
        "f_cd_lbl": "3. SHIPPING INSTRUCTION (Trucking schedule):",
        "f_sap_lbl": "4. QUERIED SAP ZMM12 STOCK REPORT:",
        "tbl1_title": "##### 📊 ACTUAL INBOUND VARIANCE & E21 LIQUIDATION BALANCE ANALYSIS",
        "phase2_title": "#### 🏛️ PHASE 2: CUSTOMS DECLARATION SYNC & REPORT EXPORT",
        "phase2_sub": "Final validation against actual customs declaration dataset to generate sign-off reports.",
        "f_hq_lbl": "1. UPLOAD CUSTOMS DECLARATION DATASET (TAB HANG):",
        "tbl2_title": "##### 📊 GLOBAL INTERLINKED COMPLIANCE RECONCILIATION MATRIX",
        "toggle_err": "🚨 Filter errors only (Hide fully matched rows to focus on discrepancies)",
        "tax_msg": "💸 Estimated Inbound Customs Duty: Approx ${:,.2f} USD based on provided tariff.",
        "btn_xlsx": "📊 DOWNLOAD AUDIT EXCEL REPORT (With conditional sheets)",
        "btn_docx": "📝 DOWNLOAD WORD MEMORANDUM FOR DIRECTOR APPROVAL",
        "toast_e21": "✅ INITIAL E21 IMPORT BALANCE SHIELD SECURED!"
    },
    "🇨🇳 中文": {
        "gate_title": "🌐 CLG SCM EXIM VN 内部系统",
        "gate_sub": "面向进口业务组的进出口单证内部核对工具",
        "pwd_label": "请输入内部安全密码:",
        "pwd_err": "❌ 访问密码错误!",
        "main_title": "🏢 E23 - E54 单证合规一体化核对模块",
        "main_sub": "专为 CLG SCM EXIM VN 进口团队定制的流水线式数据核对工具（支持Excel/PDF）",
        "config_side": "### ⚙️ 系统参数配置",
        "tol_w": "面料/称重类货物容差 (KGM, MTK):",
        "tol_c": "数量类货物容差 (PCE, PRS, PR):",
        "skip_inv": "原始发票/装箱单标题冗余行:",
        "skip_cd": "出货通知书标题冗余行:",
        "skip_erp": "SAP ZMM12 报表冗余行:",
        "skip_ecus": "海关报关单 HANG 标题冗余行:",
        "phase0_title": "#### ⚡ PHASE 0: E21 原始进口余量锁定与发货仓物料号提取",
        "phase0_sub": "目标：冻结原始 E21 进口数据作为最高基准线，同时提取干净的 R 码列表以查询 SAP ZMM12。",
        "pre_inv_lbl": "1. 上传发货仓原始发票:",
        "pre_pkl_lbl": "2. 上传发货仓原始装箱单:",
        "pre_e21_lbl": "3. 原始 E21 进口报关明细报表 (核销清算管理):",
        "pre_hsqd_lbl": "4. 单位转换系数表 (HSQĐ):",
        "pre_thue_lbl": "5. 海关税率表 (HS 编码查询):",
        "phase0_succ": "📋 自动从原始单据中检测到 {} 个唯一的物料 R 码。",
        "phase0_copy": "👇 点击下方文本框，按 Ctrl+A 再 Ctrl+C 复制列表，直接粘贴至 SAP ZMM12 多项选择框：",
        "phase1_title": "#### 📦 PHASE 1: 实际到货核对与仓储账目交叉审计",
        "phase1_sub": "系统自动执行多维数量相减，以精确锁定货物损耗与短 heavy 数量。",
        "grp_comm": "📄 原始商业单证",
        "grp_sap": "💻 内部仓储记录 (SAP)",
        "f_inv_lbl": "1. 官方申报正本发票:",
        "f_pkl_lbl": "2. 官方申报正本装箱单:",
        "f_cd_lbl": "3. 出货通知书 (派车计划):",
        "f_sap_lbl": "4. 调取的 SAP ZMM12 库存报表:",
        "tbl1_title": "##### 📊 实际收发货差异与 E21 余量核销穿透分析",
        "phase2_title": "#### 🏛️ PHASE 2: 海关报关单同步与审计报告导出",
        "phase2_sub": "与实际海关报关单数据进行最终维度核对，以导出呈批阅报告。",
        "f_hq_lbl": "1. 上传海关报关单文件 (Tab HANG):",
        "tbl2_title": "##### 📊 全局供应链穿透式合规核对矩阵",
        "toggle_err": "🚨 开启错误过滤 (隐藏完全匹配行，专注差异行修复)",
        "tax_msg": "💸 进口关税预估：基于税率表预计约为 ${:,.2f} USD。",
        "btn_xlsx": "📊 下载带颜色格式及独立错误页的 Excel 报表",
        "btn_docx": "📝 下载呈批阅 Word 备忘录 (供总经理签字)",
        "toast_e21": "✅ 原始 E21 进口核销账目基准线已成功锁定！"
    }
}

# ==========================================
# 2. CẤU HÌNH LIGHT MODE CỐ ĐỊNH PHÂN GIẢI CAO
# ==========================================
bg_color, text_color, card_bg, border_color = "#ffffff", "#0f172a", "#f8fafc", "#e2e8f0"
sidebar_bg, sidebar_text = "#f1f5f9", "#0f172a"

hide_st_style = f"""
            <style>
            #MainMenu {{visibility: hidden;}}
            footer {{visibility: hidden;}}
            header {{visibility: hidden;}}
            .stDeployButton {{display:none;}}
            .block-container {{padding-top: 1rem; padding-bottom: 2rem; background-color: {bg_color}; color: {text_color};}}
            .stApp {{background-color: {bg_color}; color: {text_color};}}
            
            /* THIẾT LẬP TƯƠNG PHẢN ĐỘ KHỐI SIDEBAR */
            [data-testid="stSidebar"] {{background-color: {sidebar_bg} !important;}}
            [data-testid="stSidebar"] p, 
            [data-testid="stSidebar"] span, 
            [data-testid="stSidebar"] label, 
            [data-testid="stSidebar"] h3 {{
                color: {sidebar_text} !important;
                font-weight: 600 !important;
            }}
            
            .metric-card {{background: {card_bg}; padding: 15px; border-radius: 10px; border: 1px solid {border_color}; border-left: 5px solid #2563eb; margin-bottom: 15px; color: {text_color};}}
            .phase-box {{border: 1px solid {border_color}; padding: 20px; border-radius: 8px; background-color: {card_bg}; margin-bottom: 25px;}}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# Lập ngôn ngữ hiển thị từ Sidebar điều hướng
with st.sidebar:
    sel_lang = st.selectbox("🌐 LANGUAGE / NGÔN NGỮ:", list(LANG_DICT.keys()))
    t = LANG_DICT[sel_lang]

def check_password():
    if "auth" not in st.session_state: st.session_state["auth"] = False
    if not st.session_state["auth"]:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.markdown(f"<h3 style='text-align: center; color: #1E3A8A; padding-top: 50px;'>{t['gate_title']}</h3>", unsafe_allow_html=True)
            st.markdown(f"<p style='text-align: center; font-size:0.9em; color:#555;'>{t['gate_sub']}</p>", unsafe_allow_html=True)
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

if 'master_hsqd' not in st.session_state: st.session_state.master_hsqd = pd.DataFrame()
if 'master_thue' not in st.session_state: st.session_state.master_thue = pd.DataFrame()
if 'baseline_e21_pool' not in st.session_state: st.session_state.baseline_e21_pool = pd.DataFrame()
if 'data_phase1' not in st.session_state: st.session_state.data_phase1 = pd.DataFrame()

# ==========================================
# 3. ĐỘNG CƠ KHỬ NHIỄU & PHÒNG VỆ FORMAT FILE Excel
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
        if pd.isna(val): return ""
        s = str(val).strip().split('.')[0] # CHẶN LỖI ĐUÔI .0 THẬP PHÂN CỦA Excel
        return re.sub(r'[^A-Z0-9]', '', s.upper())

class SmartFileReader:
    @staticmethod
    def read_smart(uploaded_file, target_keywords):
        if uploaded_file is None: return pd.DataFrame()
        ext = uploaded_file.name.split('.')[-1].lower()
        try:
            if ext in ['xlsx', 'xls', 'csv']:
                df_raw = pd.read_csv(uploaded_file, skiprows=0, on_bad_lines='skip') if ext == 'csv' else pd.read_excel(uploaded_file, skiprows=0)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                
                header_row_idx = None
                for i, row in df_raw.head(25).iterrows():
                    combined_row_str = " ".join(row.dropna().astype(str).lower())
                    if any(kw.lower() in combined_row_str for kw in target_keywords):
                        header_row_idx = i
                        break
                
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
        except Exception as e: st.error(f"Error parsing file: {e}")
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

# ==========================================
# 6. THÂN DIỄN TIẾN TUẦN TỰ WORKFLOW
# ==========================================
st.markdown(f"<h3 style='color: #1E3A8A; margin-top:0px;'>{t['main_title']}</h3>", unsafe_allow_html=True)
st.markdown(f"<p style='font-size: 0.9em; color: #555;'>{t['main_sub']}</p>", unsafe_allow_html=True)
st.markdown("---")

# PHASE 0
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown(t["phase0_title"])
st.caption(t["phase0_sub"])
c0_1, c0_2, c0_3 = st.columns(3)
with c0_1:
    f_pre_inv = st.file_uploader(t["pre_inv_lbl"], type=["xlsx","csv","pdf"])
    f_pre_pkl = st.file_uploader(t["pre_pkl_lbl"], type=["xlsx","csv","pdf"])
with c0_2:
    f_ref_e21 = st.file_uploader(t["pre_e21_lbl"], type=["xlsx","csv"])
with c0_3:
    f_hsqd = st.file_uploader(t["pre_hsqd_lbl"], type=["xlsx","csv"])
    f_thue = st.file_uploader(t["pre_thue_lbl"], type=["xlsx","csv"])

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
        st.toast(t["toast_e21"])

if f_pre_inv or f_pre_pkl:
    df_temp_i = SmartFileReader.read_smart(f_pre_inv, ['Material code', 'Material', 'Mã'])
    df_temp_p = SmartFileReader.read_smart(f_pre_pkl, ['Material code', 'Material', 'Mã'])
    raw_codes = []
    for d_node in [df_temp_i, df_temp_p]:
        if not d_node.empty:
            col_mat = get_col(d_node, ['Material code', 'Material', 'Mã'], None)
            if col_mat: raw_codes.extend(df_node[col_mat].dropna().astype(str).tolist())
    unique_r_codes = sorted(list(set([DataSanitizer.purify_material_code(c) for c in raw_codes if c and str(c).lower() != "nan"])))
    if unique_r_codes:
        st.success(t["phase0_succ"].format(len(unique_r_codes)))
        st.text_area(t["phase0_copy"], value="\n".join(unique_r_codes), height=100)
st.markdown("</div>", unsafe_allow_html=True)

# PHASE 1
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown(t["phase1_title"])
st.caption(t["phase1_sub"])
c1_1, c1_2, c1_3 = st.columns(3)
with c1_1:
    f_inv_real = st.file_uploader(t["f_inv_lbl"], type=["xlsx","csv"])
    f_pkl_real = st.file_uploader(t["f_pkl_lbl"], type=["xlsx","csv"])
with c1_2:
    f_cd_real = st.file_uploader(t["f_cd_lbl"], type=["xlsx","csv"])
with c1_3:
    f_sap_zmm = st.file_uploader(t["f_sap_lbl"], type=["xlsx","csv"])

if f_inv_real and f_sap_zmm:
    r_inv = SmartFileReader.read_smart(f_inv_real, ['Material code', 'Material', 'Quantity Customs'])
    r_pkl = SmartFileReader.read_smart(f_pkl_real, ['Material', 'Quantity', "Q'TY"])
    r_cd  = SmartFileReader.read_smart(f_cd_real, ['Mã NL', 'Material', 'Số Lượng'])
    r_sap = SmartFileReader.read_smart(f_sap_zmm, ['PO Number', 'Material', 'PO Qty'])

    c_inv_mat = get_col(r_inv, ['Material code', 'Material', 'Mã'], 1)
    df_inv = r_inv[[c_inv_mat, get_col(r_inv, ['Quantity', 'Số lượng'], 6), get_col(r_inv, ['Unit', 'ĐVT'], 5), get_col(r_inv, ['Price', 'Đơn giá'], 7), get_col(r_inv, ['Amount', 'Thành tiền'], 8)]].dropna(subset=[c_inv_mat]).copy()
    df_inv.columns = ['Ma_Vat_Tu', 'SL_INV', 'UOM_INV', 'DonGia', 'TriGia']
    df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].apply(DataSanitizer.purify_material_code)
    for c in ['SL_INV', 'DonGia', 'TriGia']: df_inv[c] = DataSanitizer.clean_numeric(df_inv[c])
    df_inv = df_inv.groupby(['Ma_Vat_Tu', 'UOM_INV', 'DonGia'], as_index=False)[['SL_INV', 'TriGia']].sum()
    
    masters = df_inv['Ma_Vat_Tu'].tolist()
    master_purified_dict = {purify_code(m): m for m in masters}
    df_inv['Math_TriGia_Check'] = np.where(abs((df_inv['SL_INV'] * df_inv['DonGia']) - df_inv['TriGia']) > 0.05, "🚨 BAD AMOUNT", "🟢 OK")

    p_mat = get_col(r_pkl, ['Material', 'Mã'], 0)
    p_qty = get_col(r_pkl, ["Q'TY", 'Quantity'], 5)
    df_pkl = pd.DataFrame()
    df_pkl['Ma_Vat_Tu'] = r_pkl[p_mat].astype(str).str.strip().upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
    df_pkl['SL_PKL'] = DataSanitizer.clean_numeric(r_pkl[p_qty])
    df_pkl = df_pkl.groupby('Ma_Vat_Tu', as_index=False)['SL_PKL'].sum()

    df_cd = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_ChiDinh'])
    if not r_cd.empty:
        cd_mat = get_col(r_cd, ['Mã NL', 'Material'], 1)
        cd_qty = get_col(r_cd, ['Số Lượng', 'Quantity'], 4)
        df_cd = pd.DataFrame()
        df_cd['Ma_Vat_Tu'] = r_cd[cd_mat].astype(str).str.strip().upper().apply(lambda x: fz_match(x, masters, master_purified_dict))
        df_cd['SL_ChiDinh'] = DataSanitizer.clean_numeric(r_cd[cd_qty])
        df_cd = df_cd.groupby('Ma_Vat_Tu', as_index=False)['SL_ChiDinh'].sum()

    c_sap_mat = get_col(r_sap, ['Material'], 10)
    c_sap_qty = get_col(r_sap, ['PO NL QTY', 'In NL QTY', 'In Qty'], 13)
    c_sap_nl = get_col(r_sap, ['NL#'], 22)
    c_sap_desc = get_col(r_sap, ['CUSTOMS DESC'], 23)
    df_sap_raw = r_sap[[c_sap_mat, c_sap_qty, c_sap_nl, c_sap_desc]].dropna(subset=[c_sap_mat]).copy()
    df_sap_raw.columns = ['Ma_Vat_Tu', 'SL_SAP_ZMM', 'Ma_NL_HQ', 'Customs_Desc']
    df_sap_raw['Ma_Vat_Tu'] = df_sap_raw['Ma_Vat_Tu'].apply(DataSanitizer.purify_material_code)
    df_sap_raw['SL_SAP_ZMM'] = DataSanitizer.clean_numeric(df_sap_raw['SL_SAP_ZMM'])
    df_sap = df_sap_raw.groupby(['Ma_Vat_Tu', 'Ma_NL_HQ', 'Customs_Desc'], as_index=False)['SL_SAP_ZMM'].sum()

    mg_p1 = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
    mg_p1 = pd.merge(mg_p1, df_cd, on='Ma_Vat_Tu', how='outer')
    mg_p1 = pd.merge(mg_p1, df_sap, on='Ma_Vat_Tu', how='outer')
    if not st.session_state.baseline_e21_pool.empty:
        mg_p1 = pd.merge(mg_p1, st.session_state.baseline_e21_pool, on='Ma_Vat_Tu', how='left')
    mg_p1 = mg_p1.fillna(0)

    mg_p1['Chênh_Lệch_PKL'] = mg_p1['SL_PKL'] - mg_p1['SL_INV']
    mg_p1['Chênh_Lệch_SAP'] = mg_p1['SL_SAP_ZMM'] - mg_p1['SL_INV']
    if 'SL_E21_GOC' in mg_p1.columns:
        mg_p1['Dư_ThanhKhoản_E21'] = mg_p1['SL_E21_GOC'] - mg_p1['SL_INV']

    def eval_p1(r):
        if r['SL_INV'] == 0: return "❌ ERROR: MISSING ON INVOICE"
        if "🚨" in str(r['Math_TriGia_Check']): return "🚨 FATAL: AMOUNT MISMATCH"
        tol = tol_weight if r['UOM_INV'] in ['KGM', 'MTK'] else tol_count
        if abs(r['Chênh_Lệch_PKL']) > tol: return "🔴 LỆCH SỐ LIỆU ĐÓNG GÓI PACKING LIST"
        if abs(r['Chênh_Lệch_SAP']) > tol: return "🚨 LỖI THẤT THOÁT: VÊNH TỒN KHO NỘI BỘ SAP"
        if 'Dư_ThanhKhoản_E21' in r and r['Dư_ThanhKhoản_E21'] < 0: return "🚨 NGUY HIỂM: VƯỢT ĐỊNH MỨC THANH KHOẢN E21 GỐC"
        return "🟢 GIAO NHẬN THỰC TẾ KHỚP HOÀN TOÀN"

    mg_p1['KẾT LUẬN PHASE 1'] = mg_p1.apply(eval_p1, axis=1)
    st.markdown(t["tbl1_title"])
    st.dataframe(mg_p1, use_container_width=True, hide_index=True)
    st.session_state.data_p1 = mg_p1
st.markdown("</div>", unsafe_allow_html=True)

# PHASE 2
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown(t["phase2_title"])
st.caption(t["phase2_sub"])
f_ecus_hang = st.file_uploader(t["f_hq_lbl"], type=["xlsx","csv","xls"])

if f_ecus_hang and not st.session_state.data_p1.empty:
    mg_p2_goc = st.session_state.data_p1.copy()
    r_hq = SmartFileReader.read_smart(f_ecus_hang, ['Mã số hàng hóa', 'Mô tả hàng hóa', 'Số lượng (1)'])
    
    masters_p2 = mg_p2_goc['Ma_Vat_Tu'].tolist()
    df_hq = parse_ecus(r_hq, masters_p2, "HAI_QUAN")
    mg_final = pd.merge(mg_p2_goc, df_hq, on='Ma_Vat_Tu', how='outer').fillna(0)

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

    st.markdown(t["tbl2_title"])
    err_filter = st.toggle(t["toggle_err"], value=False)
    board_display = mg_final[~mg_final['TRẠNG THÁI CUỐI CÙNG'].str.contains("🟢", regex=False)] if err_filter else mg_final
    
    st.data_editor(
        board_display.style.applymap(lambda x: 'background-color:#d1fae5; color:#065f46' if '🟢' in str(x) else 'background-color:#7f1d1d; color:#fca5a5; font-weight:bold' if '🚨' in str(x) else 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🔴' in str(x) or '❌' in str(x) else '', subset=['TRẠNG THÁI CUỐI CÙNG']),
        use_container_width=True, hide_index=True
    )

    tong_thue = mg_final['Thue_Du_Kien'].sum() if 'Thue_Du_Kien' in mg_final.columns else 0
    if tong_thue > 0: st.info(t["tax_msg"].format(tong_thue))

    st.markdown(f"#### {t['grp_ecus']}")
    d1, d2 = st.columns(2)
    with d1: st.download_button(t["btn_xlsx"], ExportEngine.to_excel(mg_final), "Exim_Audit_Report.xlsx", "primary", use_container_width=True)
    with d2: st.download_button(t["btn_docx"], ExportEngine.to_word(mg_final, t), "Exim_Memorandum.docx", use_container_width=True)
elif f_ecus_hang and st.session_state.data_p1.empty:
    st.error("⚠️ Please complete Phase 1 before uploading Customs Declaration file.")
st.markdown("</div>", unsafe_allow_html=True)
