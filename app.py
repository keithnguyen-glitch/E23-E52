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
        "err_miss": "⚠️ Yêu cầu tối thiểu: Phải có Invoice khai báo và SAP ZMM12.",
        
        "dl_title": "#### 🗂️ NẠP HỒ DỮ LIỆU ĐỐI CHIẾU",
        "dl_sub": "Kéo thả 6 file thô vào đây. Hỗ trợ mọi định dạng: Excel, CSV, Text, PDF, Ảnh (JPG, PNG)...",
        "f_hd03_lbl": "3. SỔ HD 03 (HẢI QUAN)",
        "f_zmm12_lbl": "4. SAP ZMM12 (Nội bộ)",
        "f_iop01_lbl": "5. SAP IOP01 (Quy đổi)",
        "f_mb52_lbl": "6. SAP MB52/MB51 (Tồn kho)",
        "eng_title": "#### 🧠 ENGINE KIỂM TOÁN & TẠO FORM ECUS",
        "spin_msg": "Đang chạy thuật toán kiểm kho và khấu trừ FIFO...",
        "succ_msg": "🎉 Cỗ máy đã thực thi xong thuật toán Tiền kiểm và Chẻ dòng ECUS!",
        "err_filter_lbl": "🚨 Chỉ hiển thị các dòng BỊ LỖI / CẢNH BÁO",
        "btn_ecus_xlsx": "📥 TẢI FILE EXCEL CHUẨN ECUS",
        "miss_file_msg": "💡 Vui lòng tải đủ 6 file hệ thống để cỗ máy tiến hành đối soát đa chiều và sinh Form ECUS.",
        "st_err_mb52": "🔴 LỖI: ÂM KHO MB52 (Tồn: {})",
        "st_warn_zmm": "🟡 CẢNH BÁO: ZMM12 THẤP HƠN NHU CẦU | ",
        "st_err_hd03_miss": "🔴 LỖI: KHÔNG TÌM THẤY TRONG SỔ HD 03",
        "st_ok": "🟢 HỢP LỆ (SẴN SÀNG KHAI)",
        "st_warn_chk": "CẦN CHECK LẠI SỔ KẾ TOÁN",
        "st_err_hd03_empty": "🔴 LỖI: SỔ HD 03 ĐÃ HẾT TỒN"
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
        "err_miss": "⚠️ Required: Official Invoice and SAP ZMM12.",
        
        "dl_title": "#### 🗂️ LOAD RECONCILIATION DATA",
        "dl_sub": "Drag and drop 6 raw files here. Supports all formats: Excel, CSV, Text, PDF, Image (JPG, PNG)...",
        "f_hd03_lbl": "3. HD 03 BOOK (CUSTOMS)",
        "f_zmm12_lbl": "4. SAP ZMM12 (Internal)",
        "f_iop01_lbl": "5. SAP IOP01 (Conversion)",
        "f_mb52_lbl": "6. SAP MB52/MB51 (Inventory)",
        "eng_title": "#### 🧠 AUDIT ENGINE & ECUS FORM GENERATION",
        "spin_msg": "Running inventory check and FIFO allocation algorithm...",
        "succ_msg": "🎉 Engine successfully executed Pre-check and ECUS line splitting!",
        "err_filter_lbl": "🚨 Show only ERROR / WARNING lines",
        "btn_ecus_xlsx": "📥 DOWNLOAD STANDARD ECUS EXCEL",
        "miss_file_msg": "💡 Please upload all 6 system files to proceed with multi-dimensional reconciliation and ECUS Form generation.",
        "st_err_mb52": "🔴 ERROR: NEGATIVE MB52 INV (Stock: {})",
        "st_warn_zmm": "🟡 WARNING: ZMM12 LOWER THAN DEMAND | ",
        "st_err_hd03_miss": "🔴 ERROR: NOT FOUND IN HD 03 BOOK",
        "st_ok": "🟢 VALID (READY TO DECLARE)",
        "st_warn_chk": "NEED TO RECHECK ACCOUNTING BOOK",
        "st_err_hd03_empty": "🔴 ERROR: HD 03 BOOK OUT OF STOCK"
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
        "err_miss": "⚠️ 必填项：官方发票和 SAP ZMM12。",
        
        "dl_title": "#### 🗂️ 加载核对数据",
        "dl_sub": "将6个原始文件拖放到此处。支持所有格式：Excel、CSV、Text、PDF、图片（JPG、PNG）...",
        "f_hd03_lbl": "3. HD 03 账册 (海关)",
        "f_zmm12_lbl": "4. SAP ZMM12 (内部)",
        "f_iop01_lbl": "5. SAP IOP01 (转换)",
        "f_mb52_lbl": "6. SAP MB52/MB51 (库存)",
        "eng_title": "#### 🧠 审计引擎 & 生成 ECUS 表单",
        "spin_msg": "正在运行库存检查和 FIFO 分配算法...",
        "succ_msg": "🎉 引擎成功执行预检和 ECUS 行拆分！",
        "err_filter_lbl": "🚨 仅显示 错误 / 警告 行",
        "btn_ecus_xlsx": "📥 下载标准 ECUS EXCEL",
        "miss_file_msg": "💡 请上传所有6个系统文件，以进行多维核对并生成 ECUS 表单。",
        "st_err_mb52": "🔴 错误: MB52 负库存 (库存: {})",
        "st_warn_zmm": "🟡 警告: ZMM12 低于需求 | ",
        "st_err_hd03_miss": "🔴 错误: 在 HD 03 账册中未找到",
        "st_ok": "🟢 有效 (准备申报)",
        "st_warn_chk": "需要重新检查会计账册",
        "st_err_hd03_empty": "🔴 错误: HD 03 账册缺货"
    }
}

# ==========================================
# 2. CẤU HÌNH UI SIÊU LỚN (BIG UI) & BẢO MẬT
# ==========================================
st.set_page_config(layout="wide", page_title="EXIM Reconciliation", page_icon="🌐", initial_sidebar_state="expanded")

big_ui_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    html, body, [class*="st-"] {
        font-size: 1.15rem !important; 
    }
    
    h1 {font-size: 2.5rem !important; font-weight: 800 !important; color: #1E3A8A;}
    h3 {font-size: 1.8rem !important; font-weight: 700 !important;}
    h4 {font-size: 1.5rem !important; font-weight: 700 !important;}
    h5 {font-size: 1.3rem !important; font-weight: 600 !important; color: #b45309;}
    
    .stFileUploader label {
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    
    [data-testid="stDataFrame"] {
        font-size: 1.1rem;
    }
    
    .block-container {padding-top: 1rem; padding-bottom: 2rem; background-color: #ffffff;}
    [data-testid="stSidebar"] {background-color: #f1f5f9 !important;}
    .phase-box {border: 2px solid #e2e8f0; padding: 25px; border-radius: 12px; background-color: #f8fafc; margin-bottom: 30px;}
    .stButton>button {border-radius: 8px; font-weight: 800; font-size: 1.2rem; padding: 0.75rem 0rem;}
    </style>
"""
st.markdown(big_ui_style, unsafe_allow_html=True)

with st.sidebar:
    sel_lang = st.selectbox("🌐 LANGUAGE / NGÔN NGỮ:", list(LANG_DICT.keys()))
    t = LANG_DICT[sel_lang]

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
        if isinstance(series, pd.DataFrame):
            series = series.iloc[:, 0]
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
        # 1. Tìm theo từ khóa
        for c in df.columns:
            for kw in kws:
                if kw.lower() in str(c).lower(): return c
        # 2. Nếu không thấy, lấy cột theo index dự phòng (nếu index đó tồn tại)
        if fb_idx is not None and fb_idx < len(df.columns):
            return df.columns[fb_idx]
        # 3. Kẹt quá thì lấy cột đầu tiên, tuyệt đối không trả về None để chống lỗi sập
        return df.columns[0] if len(df.columns) > 0 else None

    @staticmethod
    def read_smart(uploaded_file, target_keywords):
        if uploaded_file is None: return pd.DataFrame()
        ext = uploaded_file.name.split('.')[-1].lower()
        
        def deduplicate_cols(cols):
            seen = {}
            res = []
            for c in cols:
                c_str = str(c).strip()
                if c_str in seen:
                    seen[c_str] += 1
                    res.append(f"{c_str}_{seen[c_str]}")
                else:
                    seen[c_str] = 0
                    res.append(c_str)
            return res

        try:
            # 1. NHÓM EXCEL & TEXT
            if ext in ['xlsx', 'xls', 'csv', 'xlsb', 'xlsm', 'txt']:
                df_raw = None
                if ext in ['csv', 'txt']:
                    try:
                        df_raw = pd.read_csv(uploaded_file, on_bad_lines='skip', encoding='utf-8')
                    except Exception:
                        uploaded_file.seek(0)
                        df_raw = pd.read_excel(uploaded_file)
                else:
                    try:
                        df_raw = pd.read_excel(uploaded_file)
                    except Exception:
                        uploaded_file.seek(0)
                        try:
                            df_raw = pd.read_excel(uploaded_file, engine='openpyxl')
                        except:
                            uploaded_file.seek(0)
                            try:
                                df_raw = pd.read_excel(uploaded_file, engine='xlrd')
                            except:
                                uploaded_file.seek(0)
                                df_raw = pd.read_csv(uploaded_file, on_bad_lines='skip')
                
                if df_raw is None or df_raw.empty:
                    return pd.DataFrame()
                    
                df_raw.columns = deduplicate_cols(df_raw.columns)
                
                header_row_idx = None
                for i, row in df_raw.head(30).iterrows():
                    combined_str = " ".join([str(val).lower() for val in row.dropna()])
                    if any(kw.lower() in combined_str for kw in target_keywords):
                        header_row_idx = i; break
                        
                if header_row_idx is not None:
                    actual_headers = df_raw.iloc[header_row_idx].tolist()
                    df_clean = df_raw.iloc[header_row_idx + 1:].copy()
                    df_clean.columns = deduplicate_cols(actual_headers)
                    df_clean = df_clean.loc[:, ~df_clean.columns.str.contains('^nan|^Unnamed', case=False)].dropna(how='all')
                    
                    for col in df_clean.columns:
                        if any(kw.lower() in str(col).lower() for kw in target_keywords):
                            df_clean[col] = df_clean[col].ffill()
                            break
                    return df_clean
                return df_raw.dropna(how='all')
            
            # 2. NHÓM PDF
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
                        else:
                            text = page.extract_text()
                            if text:
                                for line in text.split('\n'): all_rows.append([line])
                if all_rows: 
                    return pd.DataFrame(all_rows[1:], columns=deduplicate_cols(all_rows[0]))
            
            # 3. NHÓM HÌNH ẢNH SCAN
            elif ext in ['png', 'jpg', 'jpeg']:
                img = Image.open(uploaded_file).convert('RGB')
                img.thumbnail((1600, 1600), Image.Resampling.LANCZOS)
                ocr_results = reader.readtext(np.array(img), detail=0)
                return pd.DataFrame(ocr_results, columns=["Du_Lieu_OCR"])

        except Exception as e: 
            st.error(f"🚨 Cảnh báo: File '{uploaded_file.name}' bị lỗi cấu trúc ngầm. Chi tiết: {e}")
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
                max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 5
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

# ----------------- BƯỚC 1: DATA LAKE (HỨNG 6 FILE CHỨNG TỪ & SAP) -----------------
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown(t["dl_title"])
st.caption(t["dl_sub"])

file_types = ["xlsx", "csv", "xls", "xlsb", "xlsm", "txt", "pdf", "jpg", "png", "jpeg"]

c1, c2, c3 = st.columns(3)
with c1: f_inv = st.file_uploader(t["f_inv_lbl"], type=file_types)
with c2: f_pkl = st.file_uploader(t["f_pkl_lbl"], type=file_types)
with c3: f_hd03 = st.file_uploader(t["f_hd03_lbl"], type=file_types)

c4, c5, c6 = st.columns(3)
with c4: f_zmm12 = st.file_uploader(t["f_zmm12_lbl"], type=file_types)
with c5: f_iop01 = st.file_uploader(t["f_iop01_lbl"], type=file_types)
with c6: f_mb52 = st.file_uploader(t["f_mb52_lbl"], type=file_types)
st.markdown("</div>", unsafe_allow_html=True)

# ----------------- BƯỚC 2 & 3: ENGINE TIỀN KIỂM & CHẺ DÒNG FIFO -----------------
if f_inv and f_pkl and f_hd03 and f_zmm12 and f_iop01 and f_mb52:
    st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
    st.markdown(t["eng_title"])
    
    with st.spinner(t["spin_msg"]):
        # 1. ĐỌC DỮ LIỆU
        r_inv = DataEngine.read_smart(f_inv, ['Material code', 'Material'])
        r_hd03 = DataEngine.read_smart(f_hd03, ['Mã nguyên liệu', 'Material', 'Mã NL'])
        r_iop01 = DataEngine.read_smart(f_iop01, ['Material', 'NLClass'])
        r_mb52 = DataEngine.read_smart(f_mb52, ['Material', 'Unrestricted'])
        r_zmm12 = DataEngine.read_smart(f_zmm12, ['PO Number', 'Material'])

        # 2. CHUẨN BỊ BỘ NHỚ DATA
        # Tồn kho MB52
        c_mb_mat = DataEngine.get_col(r_mb52, ['Material', 'Mã'], 1)
        c_mb_stock = DataEngine.get_col(r_mb52, ['Unrestricted', 'Tồn'], 3)
        df_mb = r_mb52[[c_mb_mat, c_mb_stock]].copy()
        df_mb.columns = ['Ma_R', 'Ton_Kho_Thuc']
        df_mb['Ma_R'] = df_mb['Ma_R'].apply(DataEngine.purify_code)
        df_mb['Ton_Kho_Thuc'] = DataEngine.clean_numeric(df_mb['Ton_Kho_Thuc'])
        dict_mb52 = df_mb.groupby('Ma_R')['Ton_Kho_Thuc'].sum().to_dict()

        # Quy đổi IOP01
        c_iop_mat = DataEngine.get_col(r_iop01, ['Material'], 1)
        c_iop_rate = DataEngine.get_col(r_iop01, ['NLRate', 'Tỷ lệ'], 7)
        c_iop_unit = DataEngine.get_col(r_iop01, ['NLUnit', 'ĐVT'], 6)
        df_iop = r_iop01[[c_iop_mat, c_iop_rate, c_iop_unit]].copy()
        df_iop.columns = ['Ma_R', 'NLRate', 'NLUnit']
        df_iop['Ma_R'] = df_iop['Ma_R'].apply(DataEngine.purify_code)
        df_iop['NLRate'] = DataEngine.clean_numeric(df_iop['NLRate'])
        df_iop = df_iop.drop_duplicates(subset=['Ma_R'])
        
        # Sổ HD 03 (Nguồn chẻ dòng)
        c_hd_mat = DataEngine.get_col(r_hd03, ['Mã nguyên liệu', 'Mã NL'], 2)
        c_hd_tk = DataEngine.get_col(r_hd03, ['Số tờ khai', 'TK'], 0)
        c_hd_date = DataEngine.get_col(r_hd03, ['Ngày tờ khai', 'Ngày'], 1)
        c_hd_ton = DataEngine.get_col(r_hd03, ['Lượng tồn', 'Lượng còn lại', 'Balance'], 4)
        c_hd_price = DataEngine.get_col(r_hd03, ['Đơn giá', 'Price'], 5)
        df_hd03 = r_hd03[[c_hd_mat, c_hd_tk, c_hd_date, c_hd_ton, c_hd_price]].copy()
        df_hd03.columns = ['Ma_R', 'So_TK_Goc', 'Ngay_TK', 'Ton_HD03', 'Don_Gia_HQ']
        df_hd03['Ma_R'] = df_hd03['Ma_R'].apply(DataEngine.purify_code)
        df_hd03['Ton_HD03'] = DataEngine.clean_numeric(df_hd03['Ton_HD03'])
        df_hd03['Don_Gia_HQ'] = DataEngine.clean_numeric(df_hd03['Don_Gia_HQ'])
        
        # Sổ ZMM12 (Để check chéo)
        c_zmm_mat = DataEngine.get_col(r_zmm12, ['Material'], 10)
        c_zmm_ton = DataEngine.get_col(r_zmm12, ['In Qty', 'Qty'], 13)
        df_zmm12 = r_zmm12[[c_zmm_mat, c_zmm_ton]].copy()
        df_zmm12.columns = ['Ma_R', 'Ton_ZMM12']
        df_zmm12['Ma_R'] = df_zmm12['Ma_R'].apply(DataEngine.purify_code)
        df_zmm12['Ton_ZMM12'] = DataEngine.clean_numeric(df_zmm12['Ton_ZMM12'])
        dict_zmm12 = df_zmm12.groupby('Ma_R')['Ton_ZMM12'].sum().to_dict()

        # Nhu cầu Invoice
        c_inv_mat = DataEngine.get_col(r_inv, ['Material code', 'Material'], 1)
        c_inv_qty = DataEngine.get_col(r_inv, ['Quantity', 'PO Qty'], 5)
        df_inv = r_inv[[c_inv_mat, c_inv_qty]].copy()
        df_inv.columns = ['Ma_R', 'Qty_Invoice']
        df_inv['Ma_R'] = df_inv['Ma_R'].apply(DataEngine.purify_code)
        df_inv['Qty_Invoice'] = DataEngine.clean_numeric(df_inv['Qty_Invoice'])
        inv_grouped = df_inv.groupby('Ma_R')['Qty_Invoice'].sum().reset_index()

        # 3. THUẬT TOÁN TIỀN KIỂM & CHẺ DÒNG FIFO
        ecus_output = []
        
        for _, inv_row in inv_grouped.iterrows():
            ma_r = inv_row['Ma_R']
            if not ma_r: continue
            
            qty_inv = inv_row['Qty_Invoice']
            
            # Tiền kiểm 1: Lấy Tỷ lệ quy đổi
            rate_info = df_iop[df_iop['Ma_R'] == ma_r]
            rate = rate_info['NLRate'].iloc[0] if not rate_info.empty else 1.0
            unit = rate_info['NLUnit'].iloc[0] if not rate_info.empty else "UNK"
            
            # Khối lượng Hải quan thực tế cần xử lý
            target_qty_hq = qty_inv * rate
            
            # Tiền kiểm 2: Kiểm tra Âm Kho (MB52)
            ton_mb52 = dict_mb52.get(ma_r, 0)
            if ton_mb52 < qty_inv:
                ecus_output.append({
                    'Mã Vật Tư': ma_r, 'ĐVT HQ': unit, 'Lượng Cần Khai': target_qty_hq,
                    'Số TK Gốc': '---', 'Đơn Giá HQ': 0, 'Trị Giá': 0,
                    'TRẠNG THÁI': t['st_err_mb52'].format(ton_mb52)
                })
                continue # Kho âm -> Không cho phép chẻ dòng, chặn luôn!
                
            # Tiền kiểm 3: Check Sổ Kế toán ZMM12
            ton_zmm = dict_zmm12.get(ma_r, 0)
            warning_zmm = ""
            if ton_zmm < target_qty_hq:
                warning_zmm = t['st_warn_zmm']

            # THỰC THI CHẺ DÒNG TRÊN SỔ HD 03
            hd03_sub = df_hd03[df_hd03['Ma_R'] == ma_r].copy()
            
            if hd03_sub.empty:
                ecus_output.append({
                    'Mã Vật Tư': ma_r, 'ĐVT HQ': unit, 'Lượng Cần Khai': target_qty_hq,
                    'Số TK Gốc': '---', 'Đơn Giá HQ': 0, 'Trị Giá': 0,
                    'TRẠNG THÁI': t['st_err_hd03_miss']
                })
                continue
                
            # Quét tuần tự từng dòng tờ khai trong HD03
            remaining_to_allocate = target_qty_hq
            for _, hd_row in hd03_sub.iterrows():
                if remaining_to_allocate <= 0: break
                
                ton_tk_goc = hd_row['Ton_HD03']
                if ton_tk_goc <= 0: continue
                
                # Bốc lượng: Lấy phần nhỏ hơn giữa nhu cầu và tồn của tờ khai
                take = min(remaining_to_allocate, ton_tk_goc)
                remaining_to_allocate -= take
                
                trang_thai = warning_zmm + t['st_ok'] if not warning_zmm else warning_zmm + t['st_warn_chk']
                
                ecus_output.append({
                    'Mã Vật Tư': ma_r, 
                    'ĐVT HQ': unit, 
                    'Lượng Cần Khai': round(take, 2),
                    'Số TK Gốc': hd_row['So_TK_Goc'], 
                    'Đơn Giá HQ': hd_row['Don_Gia_HQ'], 
                    'Trị Giá': round(take * hd_row['Don_Gia_HQ'], 2),
                    'TRẠNG THÁI': trang_thai
                })
            
            # Nếu vét sạch Sổ HD 03 mà vẫn thiếu lượng
            if remaining_to_allocate > 0:
                ecus_output.append({
                    'Mã Vật Tư': ma_r, 'ĐVT HQ': unit, 'Lượng Cần Khai': round(remaining_to_allocate, 2),
                    'Số TK Gốc': '---', 'Đơn Giá HQ': 0, 'Trị Giá': 0,
                    'TRẠNG THÁI': t['st_err_hd03_empty']
                })

        # 4. HIỂN THỊ KẾT QUẢ
        df_final = pd.DataFrame(ecus_output)
        
    st.success(t["succ_msg"])
    
    err_filter = st.toggle(t["err_filter_lbl"], value=False)
    board_display = df_final[~df_final['TRẠNG THÁI'].str.contains("🟢", regex=False)] if err_filter else df_final

    st.data_editor(
        board_display.style.map(
            lambda x: 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🔴' in str(x) 
            else ('background-color:#fef08a; color:#854d0e; font-weight:bold' if '🟡' in str(x) 
            else ('background-color:#d1fae5; color:#065f46; font-weight:bold' if '🟢' in str(x) else '')), 
            subset=['TRẠNG THÁI']
        ),
        use_container_width=True, hide_index=True, height=500
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='ECUS_UPLOAD')
    
    st.download_button(
        label=t["btn_ecus_xlsx"],
        data=output.getvalue(),
        file_name=f"ECUS_Data_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.info(t["miss_file_msg"])
