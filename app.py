import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import pdfplumber
import easyocr
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import difflib
from datetime import datetime
import warnings

warnings.filterwarnings("ignore")

# ==============================================================================
# SECTION 1 — i18n DICTIONARY (VI / EN / ZH)
# All UI strings live here. Engine messages are also localised.
# ==============================================================================
LANG_DICT: dict[str, dict] = {
    "🇻🇳 Tiếng Việt": {
        # Gate
        "gate_title": "🌐 CLG SCM EXIM _ IMPORT TEAM E54 & E23",
        "gate_sub": "Đăng nhập để sử dụng công cụ đối soát chứng từ",
        "pwd_label": "Nhập mã an ninh:",
        "pwd_err": "❌ Mã truy cập không hợp lệ!",
        # Main
        "main_title": "🏢 HỆ THỐNG ĐỐI SOÁT CHỨNG TỪ E54 - E23",
        "main_sub": "Tối ưu hóa tác vụ: Tách mã R ➔ Đối soát ma trận số lượng ➔ Sinh chuỗi Hải quan tự động.",
        # Sidebar
        "config_side": "### ⚖️ THAM SỐ DUNG SAI",
        "tol_w": "Hàng Vải/Cân ký (KGM, MTK):",
        "tol_c": "Hàng Đếm chiếc (PCE, PRS):",
        "config_skip": "### ⚙️ DÒNG THỪA (SKIPROWS)",
        "skip_inv": "Dòng thừa Invoice/PKL:",
        "skip_cd": "Dòng thừa Chỉ Định:",
        "skip_erp": "Dòng thừa SAP (ZMM12):",
        "skip_ecus": "Dòng thừa Tờ Khai Hải Quan:",
        # Data lake
        "dl_title": "#### 🗂️ NẠP HỒ DỮ LIỆU ĐỐI CHIẾU",
        "dl_sub": "Kéo thả 6 file thô vào đây. Hỗ trợ mọi định dạng: Excel, CSV, Text, PDF, Ảnh (JPG, PNG)...",
        "f_inv_lbl": "1. INVOICE (MC)",
        "f_pkl_lbl": "2. PACKING LIST (MC)",
        "f_hd03_lbl": "3. SỔ HD 03 (HẢI QUAN)",
        "f_zmm12_lbl": "4. SAP ZMM12 (Nội bộ)",
        "f_iop01_lbl": "5. SAP IOP01 (Quy đổi)",
        "f_mb52_lbl": "6. SAP MB52/MB51 (Tồn kho)",
        # Engine
        "eng_title": "#### 🧠 ENGINE KIỂM TOÁN & TẠO FORM ECUS",
        "spin_msg": "Đang chạy thuật toán kiểm kho và khấu trừ FIFO...",
        "succ_msg": "🎉 Cỗ máy đã thực thi xong thuật toán Tiền kiểm và Chẻ dòng ECUS!",
        "err_filter_lbl": "🚨 Chỉ hiển thị các dòng BỊ LỖI / CẢNH BÁO",
        "btn_ecus_xlsx": "📥 TẢI FILE EXCEL CHUẨN ECUS",
        "btn_docx": "📝 TẢI BIÊN BẢN WORD XÁC NHẬN",
        "miss_file_msg": "💡 Vui lòng tải đủ 6 file hệ thống để cỗ máy tiến hành đối soát đa chiều và sinh Form ECUS.",
        # Status messages (engine)
        "st_err_mb52": "🔴 LỖI: ÂM KHO MB52 (Tồn: {})",
        "st_warn_zmm": "🟡 CẢNH BÁO: ZMM12 THẤP HƠN NHU CẦU | ",
        "st_err_hd03_miss": "🔴 LỖI: KHÔNG TÌM THẤY TRONG SỔ HD 03",
        "st_ok": "🟢 HỢP LỆ (SẴN SÀNG KHAI)",
        "st_warn_chk": "CẦN CHECK LẠI SỔ KẾ TOÁN",
        "st_err_hd03_empty": "🔴 LỖI: SỔ HD 03 ĐÃ HẾT TỒN",
        # Metrics
        "metric_total": "Tổng dòng",
        "metric_ok": "✅ Hợp lệ",
        "metric_warn": "🟡 Cảnh báo",
        "metric_err": "🔴 Lỗi",
        "metric_value": "💰 Tổng trị giá (USD)",
        # Column headers (output table)
        "col_material": "Mã Vật Tư",
        "col_unit": "ĐVT HQ",
        "col_qty": "Lượng Cần Khai",
        "col_tk": "Số TK Gốc",
        "col_price": "Đơn Giá HQ",
        "col_value": "Trị Giá",
        "col_status": "TRẠNG THÁI",
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
        "dl_title": "#### 🗂️ LOAD RECONCILIATION DATA",
        "dl_sub": "Drag and drop 6 raw files here. Supports all formats: Excel, CSV, Text, PDF, Image (JPG, PNG)...",
        "f_inv_lbl": "1. INVOICE (MC)",
        "f_pkl_lbl": "2. PACKING LIST (MC)",
        "f_hd03_lbl": "3. HD 03 BOOK (CUSTOMS)",
        "f_zmm12_lbl": "4. SAP ZMM12 (Internal)",
        "f_iop01_lbl": "5. SAP IOP01 (Conversion)",
        "f_mb52_lbl": "6. SAP MB52/MB51 (Inventory)",
        "eng_title": "#### 🧠 AUDIT ENGINE & ECUS FORM GENERATION",
        "spin_msg": "Running inventory check and FIFO allocation algorithm...",
        "succ_msg": "🎉 Engine successfully executed Pre-check and ECUS line splitting!",
        "err_filter_lbl": "🚨 Show only ERROR / WARNING lines",
        "btn_ecus_xlsx": "📥 DOWNLOAD STANDARD ECUS EXCEL",
        "btn_docx": "📝 DOWNLOAD WORD CONFIRMATION",
        "miss_file_msg": "💡 Please upload all 6 system files to proceed with multi-dimensional reconciliation and ECUS Form generation.",
        "st_err_mb52": "🔴 ERROR: NEGATIVE MB52 STOCK (Stock: {})",
        "st_warn_zmm": "🟡 WARNING: ZMM12 LOWER THAN DEMAND | ",
        "st_err_hd03_miss": "🔴 ERROR: NOT FOUND IN HD 03 BOOK",
        "st_ok": "🟢 VALID (READY TO DECLARE)",
        "st_warn_chk": "NEED TO RECHECK ACCOUNTING BOOK",
        "st_err_hd03_empty": "🔴 ERROR: HD 03 BOOK OUT OF STOCK",
        "metric_total": "Total Rows",
        "metric_ok": "✅ Valid",
        "metric_warn": "🟡 Warnings",
        "metric_err": "🔴 Errors",
        "metric_value": "💰 Total Value (USD)",
        "col_material": "Material Code",
        "col_unit": "Customs UOM",
        "col_qty": "Declared Qty",
        "col_tk": "Source Declaration",
        "col_price": "Customs Unit Price",
        "col_value": "Total Value",
        "col_status": "STATUS",
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
        "dl_title": "#### 🗂️ 加载核对数据",
        "dl_sub": "将6个原始文件拖放到此处。支持所有格式：Excel、CSV、Text、PDF、图片（JPG、PNG）...",
        "f_inv_lbl": "1. 发票 (MC)",
        "f_pkl_lbl": "2. 装箱单 (MC)",
        "f_hd03_lbl": "3. HD 03 账册 (海关)",
        "f_zmm12_lbl": "4. SAP ZMM12 (内部)",
        "f_iop01_lbl": "5. SAP IOP01 (转换)",
        "f_mb52_lbl": "6. SAP MB52/MB51 (库存)",
        "eng_title": "#### 🧠 审计引擎 & 生成 ECUS 表单",
        "spin_msg": "正在运行库存检查和 FIFO 分配算法...",
        "succ_msg": "🎉 引擎成功执行预检和 ECUS 行拆分！",
        "err_filter_lbl": "🚨 仅显示 错误 / 警告 行",
        "btn_ecus_xlsx": "📥 下载标准 ECUS EXCEL",
        "btn_docx": "📝 下载 Word 确认函",
        "miss_file_msg": "💡 请上传所有6个系统文件，以进行多维核对并生成 ECUS 表单。",
        "st_err_mb52": "🔴 错误: MB52 负库存 (库存: {})",
        "st_warn_zmm": "🟡 警告: ZMM12 低于需求 | ",
        "st_err_hd03_miss": "🔴 错误: 在 HD 03 账册中未找到",
        "st_ok": "🟢 有效 (准备申报)",
        "st_warn_chk": "需要重新检查会计账册",
        "st_err_hd03_empty": "🔴 错误: HD 03 账册缺货",
        "metric_total": "总行数",
        "metric_ok": "✅ 有效",
        "metric_warn": "🟡 警告",
        "metric_err": "🔴 错误",
        "metric_value": "💰 总价值 (USD)",
        "col_material": "物料代码",
        "col_unit": "海关计量单位",
        "col_qty": "申报数量",
        "col_tk": "原始报关单",
        "col_price": "海关单价",
        "col_value": "总价值",
        "col_status": "状态",
    },
}

# ==============================================================================
# SECTION 2 — PAGE CONFIG & GLOBAL CSS
# ==============================================================================
st.set_page_config(
    layout="wide",
    page_title="EXIM Reconciliation",
    page_icon="🌐",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    #MainMenu, footer, header {visibility: hidden;}

    html, body, [class*="st-"] { font-size: 1.15rem !important; }

    h1  { font-size: 2.5rem  !important; font-weight: 800 !important; color: #1E3A8A; }
    h3  { font-size: 1.8rem  !important; font-weight: 700 !important; }
    h4  { font-size: 1.5rem  !important; font-weight: 700 !important; }
    h5  { font-size: 1.3rem  !important; font-weight: 600 !important; color: #b45309; }

    .stFileUploader label { font-weight: 700 !important; color: #0f172a !important; }
    [data-testid="stDataFrame"] { font-size: 1.1rem; }

    .block-container   { padding-top: 1rem; padding-bottom: 2rem; background-color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #f1f5f9 !important; }

    .phase-box {
        border: 2px solid #e2e8f0;
        padding: 25px;
        border-radius: 12px;
        background-color: #f8fafc;
        margin-bottom: 30px;
    }
    .stButton > button {
        border-radius: 8px;
        font-weight: 800;
        font-size: 1.2rem;
        padding: 0.75rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Language selector — must be resolved before any other widget
with st.sidebar:
    sel_lang = st.selectbox("🌐 LANGUAGE / NGÔN NGỮ:", list(LANG_DICT.keys()))

t: dict = LANG_DICT[sel_lang]

# ==============================================================================
# SECTION 3 — AUTHENTICATION GATE
# ==============================================================================
def check_password() -> bool:
    if st.session_state.get("auth"):
        return True
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown(
            f"<h1 style='text-align:center;padding-top:50px'>{t['gate_title']}</h1>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<p style='text-align:center;font-size:1.1em;color:#555'>{t['gate_sub']}</p>",
            unsafe_allow_html=True,
        )
        pwd = st.text_input(t["pwd_label"], type="password")
        if pwd:
            if pwd == st.secrets.get("app_password", "ChingLuh@2026"):
                st.session_state["auth"] = True
                st.rerun()
            else:
                st.error(t["pwd_err"])
    return False


if not check_password():
    st.stop()

# ==============================================================================
# SECTION 4 — SHARED RESOURCES  (cached — loaded once per session)
# ==============================================================================
@st.cache_resource
def load_ocr_reader() -> easyocr.Reader:
    return easyocr.Reader(["vi", "en"], gpu=False)


reader = load_ocr_reader()

if "data_p1" not in st.session_state:
    st.session_state.data_p1 = pd.DataFrame()

# ==============================================================================
# SECTION 5 — DATA ENGINE
# ==============================================================================
class DataEngine:
    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _deduplicate_cols(cols: list) -> list:
        """Rename duplicate column names to col, col_1, col_2 …"""
        seen: dict = {}
        result = []
        for c in cols:
            key = str(c).strip()
            if key in seen:
                seen[key] += 1
                result.append(f"{key}_{seen[key]}")
            else:
                seen[key] = 0
                result.append(key)
        return result

    @staticmethod
    def clean_numeric(series) -> pd.Series:
        """Coerce any series to float, stripping thousand-separators and junk."""
        if series is None:
            return pd.Series(dtype=float)
        # Guard: if pandas returns a DataFrame due to duplicate col names
        if isinstance(series, pd.DataFrame):
            series = series.iloc[:, 0]
        cleaned = (
            series.astype(str)
            .str.replace(",", "", regex=False)
            .str.replace(r"[^\d\.]", "", regex=True)
        )
        return pd.to_numeric(cleaned, errors="coerce").fillna(0)

    @staticmethod
    def purify_code(val) -> str:
        """Strip everything except A-Z 0-9; trim trailing decimal artefacts."""
        if pd.isna(val):
            return ""
        return re.sub(r"[^A-Z0-9]", "", str(val).strip().split(".")[0].upper())

    @staticmethod
    def fz_match(code: str, masters: list, master_purified: dict) -> str:
        """Fuzzy-match a code against a master list (purified dict + difflib)."""
        p = DataEngine.purify_code(code)
        if p in master_purified:
            return master_purified[p]
        if code in masters:
            return code
        m = difflib.get_close_matches(code, masters, n=1, cutoff=0.85)
        return m[0] if m else code

    @staticmethod
    def get_col(df: pd.DataFrame, keywords: list[str], fallback_idx: int = 0):
        """
        Find the first column whose name contains any keyword (case-insensitive).
        Falls back to column at fallback_idx, then column 0, never None.
        """
        for col in df.columns:
            for kw in keywords:
                if kw.lower() in str(col).lower():
                    return col
        if fallback_idx is not None and fallback_idx < len(df.columns):
            return df.columns[fallback_idx]
        return df.columns[0] if len(df.columns) > 0 else None

    # ------------------------------------------------------------------
    # Smart multi-format reader
    # ------------------------------------------------------------------
    @staticmethod
    def read_smart(uploaded_file, target_keywords: list[str]) -> pd.DataFrame:
        """
        Read Excel / CSV / TXT / PDF / image files uniformly.
        Auto-detects the true header row by scanning for target_keywords
        and applies forward-fill on the first matched key column (merged cells).
        """
        if uploaded_file is None:
            return pd.DataFrame()

        ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
        dedup = DataEngine._deduplicate_cols  # alias

        try:
            # ── 1. TABULAR (Excel family + CSV/TXT) ──────────────────────
            if ext in {"xlsx", "xls", "xlsb", "xlsm", "csv", "txt"}:
                df_raw: pd.DataFrame | None = None

                if ext in {"csv", "txt"}:
                    for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
                        try:
                            uploaded_file.seek(0)
                            df_raw = pd.read_csv(
                                uploaded_file, on_bad_lines="skip", encoding=enc
                            )
                            break
                        except Exception:
                            continue
                    if df_raw is None:
                        uploaded_file.seek(0)
                        df_raw = pd.read_excel(uploaded_file)
                else:
                    engines = [None, "openpyxl", "xlrd"]
                    for eng in engines:
                        try:
                            uploaded_file.seek(0)
                            df_raw = (
                                pd.read_excel(uploaded_file)
                                if eng is None
                                else pd.read_excel(uploaded_file, engine=eng)
                            )
                            break
                        except Exception:
                            continue
                    if df_raw is None:
                        # Last resort: treat as CSV
                        uploaded_file.seek(0)
                        df_raw = pd.read_csv(uploaded_file, on_bad_lines="skip")

                if df_raw is None or df_raw.empty:
                    return pd.DataFrame()

                df_raw.columns = dedup(df_raw.columns.tolist())

                # Scan first 30 rows for a real header
                header_idx: int | None = None
                for i, row in df_raw.head(30).iterrows():
                    row_str = " ".join(str(v).lower() for v in row.dropna())
                    if any(kw.lower() in row_str for kw in target_keywords):
                        header_idx = i
                        break

                if header_idx is not None:
                    new_headers = df_raw.iloc[header_idx].tolist()
                    df_clean = df_raw.iloc[header_idx + 1 :].copy()
                    df_clean.columns = dedup(new_headers)
                    # Drop fully-unnamed / nan columns and empty rows
                    df_clean = df_clean.loc[
                        :, ~df_clean.columns.str.contains(r"^nan$|^Unnamed", case=False, regex=True)
                    ].dropna(how="all")
                    # Forward-fill the first keyword-matched column (handles merged cells)
                    for col in df_clean.columns:
                        if any(kw.lower() in str(col).lower() for kw in target_keywords):
                            df_clean[col] = df_clean[col].ffill()
                            break
                    return df_clean.reset_index(drop=True)

                return df_raw.dropna(how="all").reset_index(drop=True)

            # ── 2. PDF ────────────────────────────────────────────────────
            elif ext == "pdf":
                all_rows: list = []
                with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        if tables:
                            for tbl in tables:
                                for row in tbl:
                                    clean_row = [
                                        str(c).strip().replace("\n", " ") if c else ""
                                        for c in row
                                    ]
                                    if any(clean_row):
                                        all_rows.append(clean_row)
                        else:
                            text = page.extract_text()
                            if text:
                                for line in text.split("\n"):
                                    all_rows.append([line])
                if all_rows:
                    return pd.DataFrame(
                        all_rows[1:], columns=dedup(all_rows[0])
                    ).reset_index(drop=True)

            # ── 3. SCANNED IMAGE (EasyOCR) ────────────────────────────────
            elif ext in {"png", "jpg", "jpeg"}:
                img = Image.open(uploaded_file).convert("RGB")
                img.thumbnail((1600, 1600), Image.Resampling.LANCZOS)
                ocr_results = reader.readtext(np.array(img), detail=0)
                return pd.DataFrame(ocr_results, columns=["Du_Lieu_OCR"])

        except Exception as exc:
            st.error(
                f"🚨 File **{uploaded_file.name}** — lỗi cấu trúc. Chi tiết: {exc}"
            )

        return pd.DataFrame()

    # ------------------------------------------------------------------
    # ECUS cross-reference parser (kept for legacy Phase 2 support)
    # ------------------------------------------------------------------
    @staticmethod
    def parse_ecus(
        raw_df: pd.DataFrame, master_codes: list, prefix: str
    ) -> pd.DataFrame:
        empty = pd.DataFrame(columns=["Ma_Vat_Tu", f"SL_{prefix}", "HS_Code"])
        if raw_df.empty:
            return empty
        d_col = DataEngine.get_col(raw_df, ["Mô tả", "Tên hàng"], 1)
        q_col = DataEngine.get_col(raw_df, ["Số lượng", "Lượng tính thuế"], 3)
        h_col = DataEngine.get_col(raw_df, ["Mã số", "HS"], 0)
        tmp = raw_df[[d_col, q_col, h_col]].dropna(subset=[d_col]).copy()
        codes = []
        for _, row in tmp.iterrows():
            desc = re.sub(r"[^A-Z0-9]", "", str(row[d_col]).upper())
            found = "UNK"
            for c in master_codes:
                p = DataEngine.purify_code(c)
                if p and p in desc:
                    found = c
                    break
            codes.append(found)
        tmp["Ma_Vat_Tu"] = codes
        valid = tmp[tmp["Ma_Vat_Tu"] != "UNK"].copy()
        valid[f"SL_{prefix}"] = DataEngine.clean_numeric(valid[q_col])
        valid.rename(columns={h_col: "HS_Code"}, inplace=True)
        return valid.groupby(["Ma_Vat_Tu", "HS_Code"], as_index=False)[
            f"SL_{prefix}"
        ].sum()


# ==============================================================================
# SECTION 6 — EXPORT ENGINE
# ==============================================================================
class ExportEngine:
    @staticmethod
    def to_excel_ecus(df: pd.DataFrame) -> bytes:
        """Styled ECUS Excel with traffic-light status column."""
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="ECUS_UPLOAD")
            wb = writer.book
            ws = writer.sheets["ECUS_UPLOAD"]

            fmt_hdr = wb.add_format(
                {
                    "bold": True,
                    "bg_color": "#1E3A8A",
                    "font_color": "white",
                    "border": 1,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
            fmt_red = wb.add_format(
                {"bg_color": "#FEE2E2", "font_color": "#991B1B", "border": 1}
            )
            fmt_yel = wb.add_format(
                {"bg_color": "#FEF08A", "font_color": "#854D0E", "border": 1}
            )
            fmt_grn = wb.add_format(
                {"bg_color": "#D1FAE5", "font_color": "#065F46", "border": 1}
            )
            fmt_def = wb.add_format({"border": 1})

            status_idx = len(df.columns) - 1

            for ci, col_name in enumerate(df.columns):
                ws.write(0, ci, col_name, fmt_hdr)
                col_width = (
                    max(df[col_name].astype(str).map(len).max(), len(col_name)) + 4
                )
                ws.set_column(ci, ci, col_width)

            for ri, row in df.iterrows():
                status = str(row.iloc[status_idx])
                for ci, val in enumerate(row):
                    if ci == status_idx:
                        fmt = (
                            fmt_grn
                            if "🟢" in status
                            else (fmt_yel if "🟡" in status else fmt_red)
                        )
                        ws.write(ri + 1, ci, str(val), fmt)
                    else:
                        ws.write(ri + 1, ci, val if not isinstance(val, float) or not np.isnan(val) else 0, fmt_def)

        return output.getvalue()

    @staticmethod
    def to_word(df: pd.DataFrame, txt: dict) -> bytes:
        """Word confirmation report for discrepancy lines."""
        doc = Document()
        heading = doc.add_heading(txt["btn_docx"].replace("📝 ", ""), 0)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Date/Time: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

        status_col = txt["col_status"]
        err_df = df[~df[status_col].str.contains("🟢", regex=False)]

        if err_df.empty:
            doc.add_paragraph("✅ All rows valid — no discrepancies found.")
        else:
            tbl = doc.add_table(rows=1, cols=4)
            tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            hdr[0].text = txt["col_material"]
            hdr[1].text = txt["col_qty"]
            hdr[2].text = txt["col_tk"]
            hdr[3].text = txt["col_status"]
            for _, row in err_df.head(200).iterrows():
                cells = tbl.add_row().cells
                cells[0].text = str(row[txt["col_material"]])
                cells[1].text = str(row[txt["col_qty"]])
                cells[2].text = str(row[txt["col_tk"]])
                cells[3].text = str(row[txt["col_status"]])

        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()


# ==============================================================================
# SECTION 7 — SIDEBAR CONFIG CONTROLS
# ==============================================================================
with st.sidebar:
    st.markdown(t["config_side"])
    tol_weight = st.slider(t["tol_w"], 0.0, 1.0, 0.20, 0.05)
    tol_count  = st.slider(t["tol_c"], 0.0, 0.10, 0.01, 0.01)

    st.markdown(t["config_skip"])
    skip_inv  = st.number_input(t["skip_inv"],  value=15, step=1)
    skip_cd   = st.number_input(t["skip_cd"],   value=17, step=1)
    skip_erp  = st.number_input(t["skip_erp"],  value=0,  step=1)
    skip_ecus = st.number_input(t["skip_ecus"], value=7,  step=1)

    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["auth"] = False
        st.rerun()

# ==============================================================================
# SECTION 8 — MAIN UI
# ==============================================================================
st.markdown(f"<h1>{t['main_title']}</h1>", unsafe_allow_html=True)
st.markdown(
    f"<p style='font-size:1.15em;color:#475569'>{t['main_sub']}</p>",
    unsafe_allow_html=True,
)
st.markdown("---")

# ── Data Lake ──────────────────────────────────────────────────────────────────
st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
st.markdown(t["dl_title"])
st.caption(t["dl_sub"])

ALLOWED_TYPES = ["xlsx", "csv", "xls", "xlsb", "xlsm", "txt", "pdf", "jpg", "png", "jpeg"]

c1, c2, c3 = st.columns(3)
with c1: f_inv   = st.file_uploader(t["f_inv_lbl"],   type=ALLOWED_TYPES, key="up_inv")
with c2: f_pkl   = st.file_uploader(t["f_pkl_lbl"],   type=ALLOWED_TYPES, key="up_pkl")
with c3: f_hd03  = st.file_uploader(t["f_hd03_lbl"],  type=ALLOWED_TYPES, key="up_hd03")

c4, c5, c6 = st.columns(3)
with c4: f_zmm12 = st.file_uploader(t["f_zmm12_lbl"], type=ALLOWED_TYPES, key="up_zmm12")
with c5: f_iop01 = st.file_uploader(t["f_iop01_lbl"], type=ALLOWED_TYPES, key="up_iop01")
with c6: f_mb52  = st.file_uploader(t["f_mb52_lbl"],  type=ALLOWED_TYPES, key="up_mb52")

st.markdown("</div>", unsafe_allow_html=True)

# ── Engine ─────────────────────────────────────────────────────────────────────
ALL_FILES_LOADED = all([f_inv, f_pkl, f_hd03, f_zmm12, f_iop01, f_mb52])

if ALL_FILES_LOADED:
    st.markdown("<div class='phase-box'>", unsafe_allow_html=True)
    st.markdown(t["eng_title"])

    with st.spinner(t["spin_msg"]):

        # ── Read raw files ────────────────────────────────────────────────
        r_inv   = DataEngine.read_smart(f_inv,   ["Material code", "Material"])
        r_hd03  = DataEngine.read_smart(f_hd03,  ["Mã nguyên liệu", "Material", "Mã NL"])
        r_iop01 = DataEngine.read_smart(f_iop01, ["Material", "NLClass"])
        r_mb52  = DataEngine.read_smart(f_mb52,  ["Material", "Unrestricted"])
        r_zmm12 = DataEngine.read_smart(f_zmm12, ["PO Number", "Material"])

        GC = DataEngine.get_col    # alias
        CN = DataEngine.clean_numeric

        # ── MB52 — physical stock dict ────────────────────────────────────
        df_mb = r_mb52[
            [GC(r_mb52, ["Material", "Mã"], 1), GC(r_mb52, ["Unrestricted", "Tồn"], 3)]
        ].copy()
        df_mb.columns = ["Ma_R", "Ton_Kho_Thuc"]
        df_mb["Ma_R"] = df_mb["Ma_R"].apply(DataEngine.purify_code)
        df_mb["Ton_Kho_Thuc"] = CN(df_mb["Ton_Kho_Thuc"])
        dict_mb52 = df_mb.groupby("Ma_R")["Ton_Kho_Thuc"].sum().to_dict()

        # ── IOP01 — conversion rate dict ─────────────────────────────────
        df_iop = r_iop01[
            [
                GC(r_iop01, ["Material"],        1),
                GC(r_iop01, ["NLRate", "Tỷ lệ"], 7),
                GC(r_iop01, ["NLUnit", "ĐVT"],   6),
            ]
        ].copy()
        df_iop.columns = ["Ma_R", "NLRate", "NLUnit"]
        df_iop["Ma_R"]   = df_iop["Ma_R"].apply(DataEngine.purify_code)
        df_iop["NLRate"] = CN(df_iop["NLRate"])
        df_iop = df_iop.drop_duplicates(subset=["Ma_R"])

        # ── HD03 — customs ledger ─────────────────────────────────────────
        df_hd03 = r_hd03[
            [
                GC(r_hd03, ["Mã nguyên liệu", "Mã NL"],        2),
                GC(r_hd03, ["Số tờ khai", "TK"],                0),
                GC(r_hd03, ["Ngày tờ khai", "Ngày"],            1),
                GC(r_hd03, ["Lượng tồn", "Lượng còn lại", "Balance"], 4),
                GC(r_hd03, ["Đơn giá", "Price"],                5),
            ]
        ].copy()
        df_hd03.columns = ["Ma_R", "So_TK_Goc", "Ngay_TK", "Ton_HD03", "Don_Gia_HQ"]
        df_hd03["Ma_R"]       = df_hd03["Ma_R"].apply(DataEngine.purify_code)
        df_hd03["Ton_HD03"]   = CN(df_hd03["Ton_HD03"])
        df_hd03["Don_Gia_HQ"] = CN(df_hd03["Don_Gia_HQ"])
        # Drop rows with no material code or zero balance
        df_hd03 = df_hd03[df_hd03["Ma_R"] != ""].copy()

        # ── ZMM12 — accounting book dict ──────────────────────────────────
        df_zmm12 = r_zmm12[
            [GC(r_zmm12, ["Material"],        10), GC(r_zmm12, ["In Qty", "Qty"], 13)]
        ].copy()
        df_zmm12.columns = ["Ma_R", "Ton_ZMM12"]
        df_zmm12["Ma_R"]     = df_zmm12["Ma_R"].apply(DataEngine.purify_code)
        df_zmm12["Ton_ZMM12"] = CN(df_zmm12["Ton_ZMM12"])
        dict_zmm12 = df_zmm12.groupby("Ma_R")["Ton_ZMM12"].sum().to_dict()

        # ── Invoice — demand dict ─────────────────────────────────────────
        df_inv = r_inv[
            [GC(r_inv, ["Material code", "Material"], 1), GC(r_inv, ["Quantity", "PO Qty"], 5)]
        ].copy()
        df_inv.columns = ["Ma_R", "Qty_Invoice"]
        df_inv["Ma_R"]       = df_inv["Ma_R"].apply(DataEngine.purify_code)
        df_inv["Qty_Invoice"] = CN(df_inv["Qty_Invoice"])
        inv_grouped = df_inv[df_inv["Ma_R"] != ""].groupby("Ma_R")["Qty_Invoice"].sum().reset_index()

        # ── FIFO allocation loop ──────────────────────────────────────────
        ecus_rows: list[dict] = []

        # Column name map (localised)
        C = t  # alias

        for _, inv_row in inv_grouped.iterrows():
            ma_r    = inv_row["Ma_R"]
            qty_inv = inv_row["Qty_Invoice"]

            if not ma_r or qty_inv <= 0:
                continue

            # Conversion rate & customs UOM
            rate_info = df_iop[df_iop["Ma_R"] == ma_r]
            rate = float(rate_info["NLRate"].iloc[0]) if not rate_info.empty else 1.0
            unit = str(rate_info["NLUnit"].iloc[0])  if not rate_info.empty else "UNK"
            target_hq = round(qty_inv * rate, 4)

            # ── Pre-check 1: physical stock (MB52) ───────────────────────
            ton_mb52 = dict_mb52.get(ma_r, 0)
            if ton_mb52 < qty_inv:
                ecus_rows.append(
                    {
                        C["col_material"]: ma_r,
                        C["col_unit"]:     unit,
                        C["col_qty"]:      target_hq,
                        C["col_tk"]:       "---",
                        C["col_price"]:    0,
                        C["col_value"]:    0,
                        C["col_status"]:   C["st_err_mb52"].format(ton_mb52),
                    }
                )
                continue  # blocked — cannot allocate

            # ── Pre-check 2: accounting book (ZMM12) ─────────────────────
            ton_zmm   = dict_zmm12.get(ma_r, 0)
            zmm_warn  = C["st_warn_zmm"] if ton_zmm < target_hq else ""

            # ── HD03 FIFO split ───────────────────────────────────────────
            hd03_sub = df_hd03[df_hd03["Ma_R"] == ma_r].copy()

            if hd03_sub.empty:
                ecus_rows.append(
                    {
                        C["col_material"]: ma_r,
                        C["col_unit"]:     unit,
                        C["col_qty"]:      target_hq,
                        C["col_tk"]:       "---",
                        C["col_price"]:    0,
                        C["col_value"]:    0,
                        C["col_status"]:   C["st_err_hd03_miss"],
                    }
                )
                continue

            remaining = target_hq
            for _, hd in hd03_sub.iterrows():
                if remaining <= 0:
                    break
                avail = hd["Ton_HD03"]
                if avail <= 0:
                    continue

                take      = min(remaining, avail)
                remaining = round(remaining - take, 4)

                status = (
                    zmm_warn + C["st_warn_chk"]
                    if zmm_warn
                    else C["st_ok"]
                )
                ecus_rows.append(
                    {
                        C["col_material"]: ma_r,
                        C["col_unit"]:     unit,
                        C["col_qty"]:      round(take, 2),
                        C["col_tk"]:       hd["So_TK_Goc"],
                        C["col_price"]:    hd["Don_Gia_HQ"],
                        C["col_value"]:    round(take * hd["Don_Gia_HQ"], 2),
                        C["col_status"]:   status,
                    }
                )

            # Residual after exhausting HD03
            if remaining > 0:
                ecus_rows.append(
                    {
                        C["col_material"]: ma_r,
                        C["col_unit"]:     unit,
                        C["col_qty"]:      round(remaining, 2),
                        C["col_tk"]:       "---",
                        C["col_price"]:    0,
                        C["col_value"]:    0,
                        C["col_status"]:   C["st_err_hd03_empty"],
                    }
                )

        df_final = pd.DataFrame(ecus_rows)

    # ── Success banner ────────────────────────────────────────────────────────
    st.success(t["succ_msg"])

    # ── Summary metrics ───────────────────────────────────────────────────────
    if not df_final.empty:
        n_total = len(df_final)
        n_ok    = df_final[t["col_status"]].str.contains("🟢", regex=False).sum()
        n_warn  = df_final[t["col_status"]].str.contains("🟡", regex=False).sum()
        n_err   = df_final[t["col_status"]].str.contains("🔴", regex=False).sum()
        total_v = df_final[t["col_value"]].sum()

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric(t["metric_total"], n_total)
        m2.metric(t["metric_ok"],   n_ok)
        m3.metric(t["metric_warn"], n_warn)
        m4.metric(t["metric_err"],  n_err)
        m5.metric(t["metric_value"], f"${total_v:,.2f}")

    # ── Filter toggle ─────────────────────────────────────────────────────────
    err_filter   = st.toggle(t["err_filter_lbl"], value=False)
    board_display = (
        df_final[~df_final[t["col_status"]].str.contains("🟢", regex=False)]
        if err_filter
        else df_final
    )

    # ── Styled table ──────────────────────────────────────────────────────────
    def _style_status(val: str) -> str:
        if "🔴" in str(val):
            return "background-color:#fee2e2;color:#991b1b;font-weight:bold"
        if "🟡" in str(val):
            return "background-color:#fef08a;color:#854d0e;font-weight:bold"
        if "🟢" in str(val):
            return "background-color:#d1fae5;color:#065f46;font-weight:bold"
        return ""

    st.data_editor(
        board_display.style.map(_style_status, subset=[t["col_status"]]),
        use_container_width=True,
        hide_index=True,
        height=520,
    )

    # ── Download buttons ──────────────────────────────────────────────────────
    dl1, dl2 = st.columns(2)

    with dl1:
        st.download_button(
            label=t["btn_ecus_xlsx"],
            data=ExportEngine.to_excel_ecus(df_final),
            file_name=f"ECUS_Data_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    with dl2:
        st.download_button(
            label=t["btn_docx"],
            data=ExportEngine.to_word(df_final, t),
            file_name=f"ECUS_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)

else:
    st.info(t["miss_file_msg"])
