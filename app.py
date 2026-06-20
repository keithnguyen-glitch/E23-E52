import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import pdfplumber
import easyocr
from PIL import Image
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import difflib
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & BẢO MẬT (WHITE-LABEL)
# ==========================================
st.set_page_config(layout="wide", page_title="CLG SCM EXIM VN E54-E23", page_icon="🌐", initial_sidebar_state="expanded")

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .stDeployButton {display:none;}
            .block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 98%;}
            /* Custom Sidebar */
            [data-testid="stSidebar"] {background-color: #f8fafc;}
            /* Thêm CSS cho các thẻ đồ họa nâng cao */
            .metric-card {background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); border-left: 5px solid #2563eb; margin-bottom: 15px;}
            .metric-card.alert {border-left-color: #ef4444;}
            .metric-card.success {border-left-color: #10b981;}
            .metric-title {font-size: 13px; color: #64748b; font-weight: 700; text-transform: uppercase;}
            .metric-value {font-size: 28px; color: #0f172a; font-weight: 900;}
            .stButton>button {border-radius: 8px; font-weight: 700; padding: 0.5rem 1rem; transition: all 0.3s ease;}
            .stButton>button:hover {transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

def check_password():
    if "auth" not in st.session_state: st.session_state["auth"] = False
    if not st.session_state["auth"]:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.markdown("<h2 style='text-align: center; color: #1E3A8A; padding-top: 50px;'>🌐 CỔNG ĐIỀU HÀNH CHUỖI CUNG ỨNG</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center;'>Hệ thống liên thông dữ liệu Ching Luh & Fu-Luh</p>", unsafe_allow_html=True)
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
def load_ocr_reader(): 
    return easyocr.Reader(['vi', 'en'], gpu=False)

reader = load_ocr_reader()

# =====================================================================
# 2. DATA LAYER (QUẢN LÝ BỘ NHỚ VÀ STATE)
# =====================================================================
class SessionManager:
    @staticmethod
    def init_state():
        default_states = {
            'master_hsqd': pd.DataFrame(),
            'master_thue': pd.DataFrame(),
            'transit_data': pd.DataFrame(),
            'inv_date': datetime.today().date(),
            'audit_logs': []
        }
        for key, val in default_states.items():
            if key not in st.session_state:
                st.session_state[key] = val

    @staticmethod
    def log_action(module, msg):
        st.session_state.audit_logs.append({
            "Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Module": module,
            "Action": msg
        })

SessionManager.init_state()

# =====================================================================
# 3. UTILITIES & EXTRACTORS (CÔNG CỤ TRÍCH XUẤT)
# =====================================================================
class DataExtractor:
    @staticmethod
    def read_file(file_obj, skiprows=0):
        if not file_obj: return pd.DataFrame()
        ext = file_obj.name.split('.')[-1].lower()
        try:
            if ext in ['xlsx', 'xls', 'csv']:
                df = pd.read_csv(file_obj, skiprows=skiprows, on_bad_lines='skip') if ext == 'csv' else pd.read_excel(file_obj, skiprows=skiprows)
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all')
                df.columns = [str(c).strip() for c in df.columns]
                return df
            elif ext == 'pdf':
                rows = []
                with pdfplumber.open(io.BytesIO(file_obj.read())) as pdf:
                    for p in pdf.pages:
                        for t in p.extract_tables():
                            for r in t:
                                cr = [str(c).replace('\n', ' ') if c else '' for c in r]
                                if any(cr): rows.append(cr)
                return pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame()
        except Exception as e:
            st.error(f"Lỗi Extract {file_obj.name}: {e}")
        return pd.DataFrame()

    @staticmethod
    def get_col(df, keywords, fallback_idx=0):
        for c in df.columns:
            for k in keywords:
                if k.lower() in str(c).lower(): return c
        return df.columns[fallback_idx] if len(df.columns) > fallback_idx else None

class DataCleaner:
    @staticmethod
    def norm_code(code_str):
        if pd.isna(code_str): return "UNK"
        return str(code_str).strip().upper()

    @staticmethod
    def norm_uom(uom_str):
        if pd.isna(uom_str) or str(uom_str).strip() == '': return "PCE"
        m = {'PC':'PCE', 'UNIT':'PCE', 'PCS':'PCE', 'EA':'PCE',
             'PR':'PRS', 'PAIRS':'PRS', 'PAIR':'PRS', 'PAA':'PRS',
             'KG':'KGM', 'KGS':'KGM', 'NW':'KGM', 'GW':'KGM', 'M':'MTR', 'YD':'YRD', 'YDS':'YRD', 'M2':'MTK'}
        return m.get(str(uom_str).strip().upper(), str(uom_str).strip().upper())

    @staticmethod
    def fuzzy_match(target, master_list, threshold=0.85):
        if target in master_list: return target, False
        matches = difflib.get_close_matches(target, master_list, n=1, cutoff=threshold)
        return (matches[0], True) if matches else (target, False)

class AnomalyDetector:
    @staticmethod
    def detect_price_outliers(df, price_col='DonGia'):
        if df.empty or price_col not in df.columns: return df
        df['Is_Outlier'] = False
        valid_prices = df[df[price_col] > 0][price_col]
        if len(valid_prices) > 5:
            mean = valid_prices.mean()
            std = valid_prices.std()
            if std > 0:
                df['Is_Outlier'] = ((df[price_col] - mean) / std).abs() > 3
        return df

# =====================================================================
# 4. BUSINESS ENGINES (ĐỘNG CƠ NGHIỆP VỤ XNK)
# =====================================================================
class LogisticsEngine:
    @staticmethod
    def convert_hsqd(df, qty_col, uom_col):
        h_db = st.session_state.master_hsqd
        if not h_db.empty and not df.empty and 'Ma_Vat_Tu' in df.columns:
            df = pd.merge(df, h_db, on='Ma_Vat_Tu', how='left')
            df['He_So_QD'] = df['He_So_QD'].fillna(1.0)
            df[qty_col] = df[qty_col] * df['He_So_QD']
            if uom_col in df.columns:
                df[uom_col] = df['DVT_Bao_Quan'].combine_first(df[uom_col])
            return df.drop(columns=['He_So_QD', 'DVT_Bao_Quan'], errors='ignore')
        return df

    @staticmethod
    def parse_ecus(raw_df, master_codes, prefix):
        if raw_df.empty: return pd.DataFrame(columns=['Ma_Vat_Tu', f'SL_{prefix}', 'HS_Code'])
        d_col = DataExtractor.get_col(raw_df, ['Mô tả', 'Tên hàng'], 1)
        q_col = DataExtractor.get_col(raw_df, ['Số lượng', 'Lượng tính thuế'], 3)
        h_col = DataExtractor.get_col(raw_df, ['Mã số', 'HS'], 0)
        
        tmp = raw_df[[d_col, q_col, h_col]].dropna(subset=[d_col]).copy()
        codes = []
        for _, r in tmp.iterrows():
            desc = str(r[d_col]).upper()
            found = "UNK"
            for c in master_codes:
                if c in desc:
                    found = c; break
            codes.append(found)
        tmp['Ma_Vat_Tu'] = codes
        res = tmp[tmp['Ma_Vat_Tu'] != 'UNK'].copy()
        res[f'SL_{prefix}'] = pd.to_numeric(res[q_col].astype(str).str.replace(',',''), errors='coerce').fillna(0)
        res.rename(columns={h_col: 'HS_Code'}, inplace=True)
        return res.groupby(['Ma_Vat_Tu', 'HS_Code'], as_index=False)[f'SL_{prefix}'].sum()

class ValidationEngine:
    @staticmethod
    def evaluate_e54(r):
        if r['SL_INV'] == 0: return "❌ LỖI: KHUYẾT MÃ INVOICE"
        if r.get('Is_Outlier', False): return "🟡 CẢNH BÁO: ĐƠN GIÁ BẤT THƯỜNG"
        s_pkl = abs(r['SL_INV'] - r['SL_PKL'])
        s_erp = abs(r['SL_INV'] - r['SL_XUAT_ERP']) if 'SL_XUAT_ERP' in r else 0
        s_e54 = abs(r['SL_INV'] - r['SL_E54']) if 'SL_E54' in r else 0
        tol = 0.2 if r['UOM_INV'] in ['KGM', 'MTK'] else 0.01
        
        if s_pkl <= tol and s_erp <= tol and s_e54 <= tol: return "🟢 HỢP LỆ XUẤT (KHỚP)"
        if s_pkl > tol: return "🔴 LỆCH ĐÓNG GÓI PKL"
        if s_e54 > tol: return "🔴 LỆCH KHAI BÁO E54"
        return "🟠 LỆCH ERP XUẤT"

    @staticmethod
    def evaluate_e23(r):
        s_erp = abs(r['SL_INV'] - r['SL_NHAP_ERP']) if 'SL_NHAP_ERP' in r else 0
        s_e23 = abs(r['SL_INV'] - r['SL_E23']) if 'SL_E23' in r else 0
        tol = 0.2 if r['UOM_INV'] in ['KGM', 'MTK'] else 0.01
        
        if s_erp <= tol and s_e23 <= tol: return "🟢 HOÀN TẤT E23 (KHỚP)"
        if s_erp > tol: return "🚨 LỖI THẤT THOÁT: KHO NHẬP THIẾU"
        return "🔴 LỆCH KHAI BÁO E23"

# =====================================================================
# 5. EXPORT ENGINE (BỘ TRÍCH XUẤT BÁO CÁO CỰC MẠNH)
# =====================================================================
class ExportEngine:
    @staticmethod
    def to_excel(df, sheet_name="Master_Data"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            wb = writer.book
            ws = writer.sheets[sheet_name]
            
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#1e293b', 'font_color': 'white', 'border': 1})
            fmt_fatal = wb.add_format({'bg_color': '#7f1d1d', 'font_color': '#fca5a5', 'bold': True, 'border': 1})
            fmt_red = wb.add_format({'bg_color': '#fee2e2', 'font_color': '#991b1b', 'border': 1})
            fmt_green = wb.add_format({'bg_color': '#d1fae5', 'font_color': '#065f46', 'border': 1})
            fmt_warn = wb.add_format({'bg_color': '#fef3c7', 'font_color': '#b45309', 'border': 1})
            
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)
                ws.set_column(col_idx, col_idx, 15)
            
            ws.set_column(0, 0, 30) 
            status_col_idx = len(df.columns) - 1
            ws.set_column(status_col_idx, status_col_idx, 40)
            
            for row_idx, status_val in enumerate(df.iloc[:, status_col_idx]):
                r = row_idx + 1
                val_str = str(status_val)
                fmt = fmt_warn
                if "FATAL" in val_str: fmt = fmt_fatal
                elif "🔴" in val_str or "❌" in val_str: fmt = fmt_red
                elif "🟢" in val_str or "🟡" in val_str: fmt = fmt_green
                ws.write(r, status_col_idx, val_str, fmt)
            
            err_df = df[~df.iloc[:, status_col_idx].str.contains("KHỚP|HOÀN TẤT", regex=True)]
            if not err_df.empty:
                err_df.to_excel(writer, index=False, sheet_name="Errors_Action_Required")
                ws_err = writer.sheets["Errors_Action_Required"]
                for col_idx, col_name in enumerate(err_df.columns): ws_err.write(0, col_idx, col_name, fmt_header)
                ws_err.set_column(0, 0, 30)
                ws_err.set_column(status_col_idx, status_col_idx, 40)
        return output.getvalue()

    @staticmethod
    def to_pdf(df, title, summary, total_tax=0):
        class PDFReport(FPDF):
            def header(self):
                self.set_font('Arial', 'B', 15)
                self.set_text_color(30, 58, 138)
                self.cell(0, 10, 'SUPPLY CHAIN AUDIT REPORT', 0, 1, 'C')
                self.set_font('Arial', 'I', 10)
                self.set_text_color(150, 150, 150)
                self.cell(0, 5, 'Automated Audit System E54-E23', 0, 1, 'C')
                self.line(10, 25, 200, 25)
                self.ln(10)
            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', 'I', 8)
                self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

        pdf = PDFReport()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 14)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(0, 10, title, ln=1)
        pdf.set_font('Arial', '', 11)
        pdf.multi_cell(0, 8, summary)
        
        if total_tax > 0:
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.set_text_color(220, 38, 38)
            pdf.cell(0, 10, f"Total Estimated Tax: ${total_tax:,.2f} USD", ln=1)
            pdf.set_text_color(0, 0, 0)
        
        pdf.ln(5)
        err_df = df[~df.iloc[:, -1].str.contains("KHỚP|HOÀN TẤT", regex=True)]
        if not err_df.empty:
            pdf.set_font('Arial', 'B', 11)
            pdf.cell(0, 10, f"Discrepancies Found ({len(err_df)} items):", ln=1)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(50, 8, "Material Code", 1, 0, 'C', 1)
            pdf.cell(25, 8, "UOM", 1, 0, 'C', 1)
            pdf.cell(115, 8, "Error Description", 1, 1, 'C', 1)
            
            pdf.set_text_color(0, 0, 0)
            pdf.set_font('Arial', '', 8)
            for i, r in err_df.head(20).iterrows():
                fill = True if i % 2 == 0 else False
                if fill: pdf.set_fill_color(245, 245, 245)
                pdf.cell(50, 8, str(r.iloc[0])[:20], 1, 0, 'L', fill)
                pdf.cell(25, 8, str(r.iloc[1])[:10], 1, 0, 'C', fill)
                pdf.cell(115, 8, str(r.iloc[-1])[:60], 1, 1, 'L', fill)
        else:
            pdf.set_font('Arial', 'B', 12)
            pdf.set_text_color(22, 163, 74)
            pdf.cell(0, 10, "SUCCESS: All data matched perfectly. Ready for Customs.", ln=1)
        
        return bytes(pdf.output())

    @staticmethod
    def to_word(df, date_str):
        doc = Document()
        doc.add_heading('BIÊN BẢN XÁC NHẬN SỐ LIỆU E54-E23', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Trích xuất ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Invoice gốc: {date_str}")
        
        p = doc.add_paragraph()
        p.add_run("Bên Giao: ").bold = True
        p.add_run("Công ty Ching Luh\n")
        p.add_run("Bên Nhận: ").bold = True
        p.add_run("Công ty Fu-Luh\n")
        
        doc.add_heading('1. Kết quả kiểm kê dữ liệu', level=1)
        err_df = df[~df.iloc[:, -1].str.contains("KHỚP|HOÀN TẤT", regex=True)]
        
        if err_df.empty:
            doc.add_paragraph("Hai bên xác nhận số liệu đóng gói, vận chuyển và khai báo Hải quan hoàn toàn trùng khớp.").bold = True
        else:
            doc.add_paragraph(f"Phát hiện {len(err_df)} mã vật tư có sự sai lệch. Yêu cầu hai bên kiểm tra lại danh sách dưới đây:").bold = True
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Mã Vật Tư', 'ĐVT', 'Mô tả sai lệch'
            for _, r in err_df.head(20).iterrows():
                row = table.add_row().cells
                row[0].text, row[1].text, row[2].text = str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[-1])
        
        doc.add_paragraph("\n")
        sign_table = doc.add_table(rows=2, cols=2)
        sign_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c1, c2 = sign_table.rows[0].cells
        c1.text = 'ĐẠI DIỆN BÊN GIAO (Ký tên)'
        c2.text = 'ĐẠI DIỆN BÊN NHẬN (Ký tên)'
        c1.paragraphs[0].alignment = c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()

# =====================================================================
# 6. ROUTING VIEW (GIAO DIỆN CÁC PHÂN HỆ)
# =====================================================================

def view_dashboard():
    st.markdown("## 📊 TỔNG QUAN CHUỖI CUNG ỨNG (DASHBOARD BI)")
    st.info("Trang chủ cung cấp cái nhìn toàn cảnh về tình trạng lưu thông hàng hóa và tỷ lệ sai sót chứng từ.")
    
    tdf = st.session_state.transit_data
    if tdf.empty:
        st.warning("Hệ thống chưa ghi nhận lô hàng nào được xuất từ phân hệ Ching Luh.")
        return
        
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='metric-card'><div class='metric-title'>Mã VT Đang Vận Chuyển</div><div class='metric-value'>{len(tdf)}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='metric-card'><div class='metric-title'>Tổng Trị Giá (USD)</div><div class='metric-value'>${tdf['TriGia'].sum():,.2f}</div></div>", unsafe_allow_html=True)
    
    errs = len(tdf[~tdf['TRẠNG THÁI XUẤT'].str.contains('🟢')])
    err_rate = (errs / len(tdf)) * 100
    color_class = "success" if err_rate < 5 else "alert"
    c3.markdown(f"<div class='metric-card {color_class}'><div class='metric-title'>Vật Tư Sai Lệch</div><div class='metric-value'>{errs} ({err_rate:.1f}%)</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='metric-card'><div class='metric-title'>Ngày Tạo Lô Hàng</div><div class='metric-value'>{st.session_state.inv_date.strftime('%d/%m/%Y')}</div></div>", unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Phân bổ Tình trạng Đối chiếu**")
        p_data = tdf['TRẠNG THÁI XUẤT'].value_counts().reset_index()
        p_data.columns = ['Status', 'Count']
        fig1 = px.pie(p_data, values='Count', names='Status', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig1, use_container_width=True)
    with g2:
        st.markdown("**Top 10 Vật tư có Giá trị lớn nhất**")
        top10 = tdf.nlargest(10, 'TriGia')
        fig2 = px.bar(top10, x='Ma_Vat_Tu', y='TriGia', color='TriGia', color_continuous_scale='Teal')
        st.plotly_chart(fig2, use_container_width=True)

def view_master_data():
    st.markdown("## ⚙️ QUẢN LÝ MASTER DATA (HSQĐ & THUẾ)")
    st.caption("Dữ liệu gốc nạp tại đây sẽ tự động chạy ngầm cho tất cả các tính toán ở các phân hệ khác.")
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 1. Nạp Bảng Hệ Số Quy Đổi")
        f_h = st.file_uploader("Upload Bảng HSQĐ", type=["xlsx", "csv"])
        if f_h:
            df = DataExtractor.read_file(f_h, 1)
            if not df.empty:
                df = df[[DataExtractor.get_col(df, ['Mã'], 1), DataExtractor.get_col(df, ['Hệ số'], 6), DataExtractor.get_col(df, ['Đơn vị'], 8)]].dropna()
                df.columns = ['Ma_Vat_Tu', 'He_So_QD', 'DVT_Bao_Quan']
                df['Ma_Vat_Tu'] = df['Ma_Vat_Tu'].apply(DataCleaner.norm_code)
                df['He_So_QD'] = pd.to_numeric(df['He_So_QD'], errors='coerce').fillna(1.0)
                df['DVT_Bao_Quan'] = df['DVT_Bao_Quan'].apply(DataCleaner.norm_uom)
                st.session_state.master_hsqd = df
                SessionManager.log_action("Master Data", f"Nạp {len(df)} dòng HSQĐ")
                st.success("✅ Lưu Hệ Số Quy Đổi thành công!")
        if not st.session_state.master_hsqd.empty: st.dataframe(st.session_state.master_hsqd.head(3), use_container_width=True)

    with c2:
        st.markdown("#### 2. Nạp Bảng Biểu Thuế")
        f_t = st.file_uploader("Upload Biểu Thuế", type=["xlsx", "csv"])
        if f_t:
            df = DataExtractor.read_file(f_t, 2)
            if not df.empty:
                df = df[[DataExtractor.get_col(df, ['Mã hàng', 'HS'], 1), DataExtractor.get_col(df, ['Thuế suất', 'NK'], 5)]].dropna()
                df.columns = ['HS_Code', 'Thue_Suat']
                df['HS_Code'] = df['HS_Code'].astype(str).str.replace('.', '').str.strip()
                st.session_state.master_thue = df
                SessionManager.log_action("Master Data", f"Nạp {len(df)} dòng Thuế")
                st.success("✅ Lưu Biểu Thuế thành công!")
        if not st.session_state.master_thue.empty: st.dataframe(st.session_state.master_thue.head(3), use_container_width=True)

def view_chingluh_e54():
    st.markdown("## 📤 PHÂN HỆ CHING LUH (XUẤT HÀNG E54)")
    st.session_state.inv_date = st.date_input("🗓️ Ngày lập Invoice Xuất:", st.session_state.inv_date)
    
    with st.expander("⚙️ Tùy chỉnh dòng tiêu đề thừa (Skiprows)"):
        c_sk1, c_sk2, c_sk3 = st.columns(3)
        with c_sk1: sk_i = st.number_input("Invoice/PKL:", value=15)
        with c_sk2: sk_e = st.number_input("ERP Kho:", value=0)
        with c_sk3: sk_hq = st.number_input("ECUS E54:", value=18)

    c1, c2, c3 = st.columns(3)
    with c1: 
        f_inv = st.file_uploader("1. INVOICE KHAI BÁO", type=["xlsx", "csv", "pdf"])
        f_pkl = st.file_uploader("2. PACKING LIST", type=["xlsx", "csv", "pdf"])
    with c2:
        f_erp = st.file_uploader("3. ERP XUẤT KHO", type=["xlsx", "csv"])
    with c3:
        f_e54 = st.file_uploader("4. TỜ KHAI E54", type=["xlsx", "csv"])

    if st.button("🚀 XỬ LÝ ĐỐI CHIẾU & CHỐT HÀNG ĐI ĐƯỜNG", type="primary", use_container_width=True):
        if not (f_inv and f_pkl):
            st.error("⚠️ Bắt buộc phải upload INVOICE và PACKING LIST!")
            return
            
        with st.spinner("Động cơ XNK đang xử lý dữ liệu..."):
            r_inv = DataExtractor.read_file(f_inv, sk_i)
            r_pkl = DataExtractor.read_file(f_pkl, sk_i)
            r_erp = DataExtractor.read_file(f_erp, sk_e) if f_erp else pd.DataFrame()
            r_e54 = DataExtractor.read_file(f_e54, sk_hq) if f_e54 else pd.DataFrame()

            # INV
            c_mat = DataExtractor.get_col(r_inv, ['Material', 'Mã'], 1)
            df_inv = r_inv[[c_mat, DataExtractor.get_col(r_inv, ['Quantity'], 6), DataExtractor.get_col(r_inv, ['Unit'], 5), DataExtractor.get_col(r_inv, ['Price'], 7), DataExtractor.get_col(r_inv, ['Amount'], 8)]].dropna(subset=[c_mat]).copy()
            df_inv.columns = ['Ma_Vat_Tu', 'SL_INV', 'UOM_INV', 'DonGia', 'TriGia']
            df_inv['Ma_Vat_Tu'] = df_inv['Ma_Vat_Tu'].apply(DataCleaner.norm_code)
            for c in ['SL_INV', 'DonGia', 'TriGia']: df_inv[c] = pd.to_numeric(df_inv[c].astype(str).str.replace(',',''), errors='coerce').fillna(0)
            
            # Anomaly Detection
            df_inv = AnomalyDetector.detect_price_outliers(df_inv, 'DonGia')
            df_inv = LogisticsEngine.convert_hsqd(df_inv, 'SL_INV', 'UOM_INV')
            df_inv['UOM_INV'] = df_inv['UOM_INV'].apply(DataCleaner.norm_uom)
            masters = df_inv['Ma_Vat_Tu'].tolist()

            # PKL
            c_mat_pkl = DataExtractor.get_col(r_pkl, ['Material'], 0)
            df_pkl = r_pkl[[c_mat_pkl, DataExtractor.get_col(r_pkl, ["Q'TY"], 5)]].dropna(subset=[c_mat_pkl]).copy()
            df_pkl.columns = ['Ma_Vat_Tu', 'SL_PKL']
            df_pkl['Ma_Vat_Tu'] = df_pkl['Ma_Vat_Tu'].apply(DataCleaner.norm_code).apply(lambda x: DataCleaner.fuzzy_match(x, masters)[0])
            df_pkl['SL_PKL'] = pd.to_numeric(df_pkl['SL_PKL'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
            df_pkl['U_TMP'] = None
            df_pkl = LogisticsEngine.convert_hsqd(df_pkl, 'SL_PKL', 'U_TMP').drop(columns=['U_TMP'])
            df_pkl = df_pkl.groupby('Ma_Vat_Tu', as_index=False)['SL_PKL'].sum()

            # ERP
            df_erp = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_XUAT_ERP'])
            if not r_erp.empty:
                df_erp = r_erp[[r_erp.columns[0], r_erp.columns[1]]].dropna().copy()
                df_erp.columns = ['Ma_Vat_Tu', 'SL_XUAT_ERP']
                df_erp['Ma_Vat_Tu'] = df_erp['Ma_Vat_Tu'].apply(DataCleaner.norm_code).apply(lambda x: DataCleaner.fuzzy_match(x, masters)[0])
                df_erp['SL_XUAT_ERP'] = pd.to_numeric(df_erp['SL_XUAT_ERP'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                df_erp['U_TMP'] = None
                df_erp = LogisticsEngine.convert_hsqd(df_erp, 'SL_XUAT_ERP', 'U_TMP').drop(columns=['U_TMP'])

            # E54
            df_e54 = LogisticsEngine.parse_ecus(r_e54, masters, "E54")

            # Merge
            mg = pd.merge(df_inv, df_pkl, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_erp, on='Ma_Vat_Tu', how='outer')
            mg = pd.merge(mg, df_e54, on='Ma_Vat_Tu', how='outer').fillna(0)
            mg['TRẠNG THÁI XUẤT'] = mg.apply(ValidationEngine.evaluate_e54, axis=1)
            mg = mg.sort_values(by='TRẠNG THÁI XUẤT', ascending=False)
            
            # Save State
            st.session_state.transit_data = mg[['Ma_Vat_Tu', 'UOM_INV', 'SL_INV', 'DonGia', 'TriGia', 'TRẠNG THÁI XUẤT']].copy()
            SessionManager.log_action("ChingLuh E54", f"Chốt xuất lô hàng {len(mg)} mã.")
            
            st.success("✅ Đã tạo lô hàng Transit. Dữ liệu sẵn sàng chuyển sang Fu-Luh.")
            st.dataframe(mg.style.applymap(lambda x: 'background-color:#d1fae5; color:#065f46; font-weight:bold' if '🟢' in str(x) else 'background-color:#fee2e2; color:#991b1b; font-weight:bold' if '🔴' in str(x) or '❌' in str(x) else '', subset=['TRẠNG THÁI XUẤT']), use_container_width=True, hide_index=True, height=400)
            
            # EXPORT BUTTONS
            st.markdown("### 📥 TRÍCH XUẤT BÁO CÁO XUẤT HÀNG (E54)")
            col_x1, col_x2 = st.columns(2)
            with col_x1:
                st.download_button("📊 TẢI FILE EXCEL TỔNG HỢP (E54)", ExportEngine.to_excel(mg, "ChingLuh_E54"), "Report_ChingLuh_E54.xlsx", "primary", use_container_width=True)
            with col_x2:
                st.download_button("📑 TẢI BIÊN BẢN PDF (E54)", ExportEngine.to_pdf(mg, "BIEN BAN KIEM KE XUAT KHO E54", f"So luong vat tu: {len(mg)}"), "Report_ChingLuh_E54.pdf", use_container_width=True)

def view_fuluh_e23():
    st.markdown("## 📥 PHÂN HỆ FU-LUH (NHẬP QUYẾT TOÁN E23)")
    
    tdf = st.session_state.transit_data
    if tdf.empty:
        st.warning("📭 Không có lô hàng nào đang chờ nhập. Vui lòng xử lý phần Xuất trước.")
        return
        
    deadline = st.session_state.inv_date + timedelta(days=15)
    days_left = (deadline - datetime.today().date()).days
    status_color = "#ef4444" if days_left < 0 else "#f59e0b" if days_left <= 3 else "#10b981"
    status_text = "QUÁ HẠN HẢI QUAN!" if days_left < 0 else f"Còn {days_left} ngày"
    
    st.markdown(f"""
        <div class="metric-card" style="border-left-color: {status_color};">
            <div class="metric-title">⏰ THEO DÕI HẠN CHÓT TỜ KHAI E23</div>
            <div class="metric-value" style="color: {status_color};">Hạn chót: {deadline.strftime('%d/%m/%Y')} ({status_text})</div>
        </div>
    """, unsafe_allow_html=True)

    with st.expander("⚙️ Cấu hình Header dòng trống"):
        c_sk1, c_sk2 = st.columns(2)
        with c_sk1: sk_e = st.number_input("Header ERP Kho Nhập:", value=0)
        with c_sk2: sk_hq = st.number_input("Header ECUS E23:", value=18)

    c1, c2 = st.columns(2)
    with c1: f_erp = st.file_uploader("1. ERP NHẬP KHO THỰC TẾ (Fu-Luh)", type=["xlsx", "csv"])
    with c2: f_e23 = st.file_uploader("2. TỜ KHAI HẢI QUAN E23", type=["xlsx", "csv"])

    if st.button("🔍 KIỂM ĐỊNH QUYẾT TOÁN E54 - E23", type="primary", use_container_width=True):
        with st.spinner("Đang chốt sổ và tính thuế..."):
            r_erp = DataExtractor.read_file(f_erp, sk_e) if f_erp else pd.DataFrame()
            r_e23 = DataExtractor.read_file(f_e23, sk_hq) if f_e23 else pd.DataFrame()
            masters = tdf['Ma_Vat_Tu'].tolist()

            # ERP
            df_erp = pd.DataFrame(columns=['Ma_Vat_Tu', 'SL_NHAP_ERP'])
            if not r_erp.empty:
                df_erp = r_erp[[r_erp.columns[0], r_erp.columns[1]]].dropna().copy()
                df_erp.columns = ['Ma_Vat_Tu', 'SL_NHAP_ERP']
                df_erp['Ma_Vat_Tu'] = df_erp['Ma_Vat_Tu'].apply(DataCleaner.norm_code).apply(lambda x: DataCleaner.fuzzy_match(x, masters)[0])
                df_erp['SL_NHAP_ERP'] = pd.to_numeric(df_erp['SL_NHAP_ERP'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                df_erp['U_TMP'] = None
                df_erp = LogisticsEngine.convert_hsqd(df_erp, 'SL_NHAP_ERP', 'U_TMP').drop(columns=['U_TMP'])

            # E23
            df_e23 = LogisticsEngine.parse_ecus(r_e23, masters, "E23")

            mg = pd.merge(tdf, df_erp, on='Ma_Vat_Tu', how='left')
            mg = pd.merge(mg, df_e23, on='Ma_Vat_Tu', how='left').fillna(0)
            
            # Tính Thuế
            mg['Thue_NK_DuKien'] = 0.0
            if not st.session_state.master_thue.empty and 'HS_Code' in mg.columns:
                tax_dict = dict(zip(st.session_state.master_thue['HS_Code'], st.session_state.master_thue['Thue_Suat']))
                mg['Thue_Suat_ApDung'] = mg['HS_Code'].astype(str).str.replace('.','').map(tax_dict).fillna(0)
                mg['Thue_NK_DuKien'] = mg['TriGia'] * (mg['HS_Code'].astype(str).str.replace('.','').map(tax_dict).fillna(0).astype(float) / 100)

            mg['KẾT LUẬN CUỐI CÙNG'] = mg.apply(ValidationEngine.evaluate_e23, axis=1)
            mg = mg.sort_values(by='KẾT LUẬN CUỐI CÙNG', ascending=False)
            
            st.dataframe(mg.style.applymap(lambda x: 'background-color:#fee2e2; color:#7f1d1d; font-weight:bold' if '🚨' in str(x) else 'background-color:#d1fae5; color:#065f46' if '🟢' in str(x) else '', subset=['KẾT LUẬN CUỐI CÙNG']), use_container_width=True, hide_index=True)
            
            total_tax = mg['Thue_NK_DuKien'].sum() if 'Thue_NK_DuKien' in mg.columns else 0

            # XUẤT ĐA ĐỊNH DẠNG
            st.markdown("---")
            st.markdown("### 📥 TRÍCH XUẤT KẾT QUẢ QUYẾT TOÁN TỔNG HỢP (E23/E54)")
            cx1, cx2, cx3 = st.columns(3)
            with cx1:
                st.download_button("📊 TẢI EXCEL (TÔ MÀU)", ExportEngine.to_excel(mg, "Quyet_Toan_E23"), "Final_Audit_E54_E23.xlsx", "primary", use_container_width=True)
            with cx2:
                st.download_button("📑 TẢI PDF (BÁO CÁO CÔNG SỞ)", ExportEngine.to_pdf(mg, "BIEN BAN QUYET TOAN CHUOI CUNG UNG", "So lieu tich hop E54-E23.", total_tax), "Final_Audit_E54_E23.pdf", use_container_width=True)
            with cx3:
                st.download_button("📝 TẢI WORD (TRÌNH KÝ)", ExportEngine.to_word(mg, st.session_state.inv_date.strftime('%d/%m/%Y')), "Agreement_E54_E23.docx", use_container_width=True)

# =====================================================================
# 8. APP ROUTER
# =====================================================================
st.sidebar.markdown("### 🏢 MENU ĐIỀU HƯỚNG")
menu = st.sidebar.radio("Chọn chức năng:", [
    "📊 Dashboard Tổng Quan",
    "⚙️ Master Data (Dữ Liệu Nền)",
    "📤 Phân Hệ Ching Luh (E54)",
    "📥 Phân Hệ Fu-Luh (E23)"
])

if menu == "📊 Dashboard Tổng Quan": view_dashboard()
elif menu == "⚙️ Master Data (Dữ Liệu Nền)": view_master_data()
elif menu == "📤 Phân Hệ Ching Luh (E54)": view_chingluh_export()
elif menu == "📥 Phân Hệ Fu-Luh (E23)": view_fuluh_e23()
