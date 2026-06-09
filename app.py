import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl
import textwrap
import re

# ==================================================
# PAGE CONFIG
# ==================================================
st.set_page_config(
    page_title="PEA Executive Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================================================
# SESSION STATE INITIALIZATION
# ==================================================
if 'app_state' not in st.session_state:
    st.session_state.app_state = 'upload' 
if 'summary_df' not in st.session_state:
    st.session_state.summary_df = None
if 'all_df' not in st.session_state:
    st.session_state.all_df = None

# ==================================================
# CSS GLOBAL (สำหรับหน้า Dashboard)
# ==================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"]  {
    font-family: 'Sarabun', sans-serif;
}

/* พื้นหลังเว็บสี Slate-50 ให้การ์ดสีขาวดูโดดเด่น */
[data-testid="stAppViewContainer"] {
    background-color: #F8FAFC; 
    color: #1E293B; /* Slate-800 */
}
[data-testid="stHeader"] {
    background-color: transparent;
}
.block-container { padding-top: 2rem; padding-bottom: 2rem; }

/* HEADER: Deep Navy -> Indigo Purple -> Azure Blue */
.pea-header {
    background: linear-gradient(135deg, #0A192F, #4C1D95, #0284C7); 
    color: white;
    padding: 32px;
    border-radius: 20px;
    margin-bottom: 32px;
    box-shadow: 0 10px 25px -5px rgba(76, 29, 149, 0.4); 
    border: 1px solid rgba(255,255,255,0.1);
}
.pea-title { font-size: 36px; font-weight: 700; margin-bottom: 8px; color: #FFFFFF; letter-spacing: -0.5px;}
.pea-subtitle { font-size: 16px; font-weight: 400; color: #E0E7FF; letter-spacing: 0.5px;}

/* GLASSMORPHISM CARDS (การ์ดกระจก) */
.glass-card {
    background: rgba(255, 255, 255, 0.9); 
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    border-radius: 20px;
    padding: 24px;
    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -2px rgba(0, 0, 0, 0.025); 
    border: 1px solid rgba(255, 255, 255, 0.6);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
    margin-bottom: 24px;
}
.glass-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
}

/* Typography สำหรับ KPI */
.kpi-title { font-size: 15px; color: #64748B; margin-bottom: 8px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;}
.kpi-value { font-size: 36px; font-weight: 700; line-height: 1.2; }

/* Category Cards Specifics */
.category-card { min-height: 220px; }
.category-number { color: #3730A3; font-size: 24px; font-weight: 800; opacity: 0.2; float: right;}
.category-title { font-size: 18px; font-weight: 700; margin: 0 0 12px 0; color: #0F172A;}
.category-value { font-size: 15px; margin-bottom: 8px; color: #475569; font-weight: 500;}
.category-note { color: #0EA5E9; margin-top: 16px; font-size: 13px; border-top: 1px solid #F1F5F9; padding-top: 12px; font-weight: 500;}

/* Table Headers */
.table-header-done { color: #0BFF8D; font-size: 18px; font-weight: 700; margin-top: 25px; border-bottom: 2px solid #0BFF8D; padding-bottom: 8px;}
.table-header-prog { color: #D58A19; font-size: 18px; font-weight: 700; margin-top: 25px; border-bottom: 2px solid #D58A19; padding-bottom: 8px;}
.table-header-not  { color: #F43F5E; font-size: 18px; font-weight: 700; margin-top: 25px; border-bottom: 2px solid #F43F5E; padding-bottom: 8px;}

/* 📌 Custom CSS สำหรับ HTML Table แบบจัดระเบียบ */
.styled-table {
    width: 100%;
    border-collapse: collapse;
    margin: 15px 0 20px 0;
    font-size: 14px;
    background-color: white;
    border: 1px solid #E2E8F0;
}
.styled-table thead tr {
    background-color: #F8FAFC;
    color: #475569;
}
.styled-table th {
    padding: 12px 15px;
    border: 1px solid #E2E8F0;
    text-align: center !important; /* จัดกึ่งกลางหัวตาราง */
    vertical-align: middle;
    white-space: nowrap; /* หัวตารางไม่ตกบรรทัด */
}
.styled-table td {
    padding: 12px 15px;
    border: 1px solid #E2E8F0;
    vertical-align: top;
    color: #334155;
    white-space: nowrap; /* ค่าเริ่มต้นให้ข้อมูลทุกช่องอยู่บรรทัดเดียว (เช่น จังหวัด วงเงิน) */
}
/* 📌 บังคับให้คอลัมน์ 'ชื่องาน' (คอลัมน์ 2) และ 'หมายเหตุ' (คอลัมน์ 8) ปัดบรรทัดใหม่ได้ */
.styled-table td:nth-child(2),
.styled-table td:nth-child(8) {
    white-space: normal !important; 
    word-wrap: break-word;
    min-width: 250px; 
}
.styled-table tbody tr:hover {
    background-color: #F1F5F9;
}

/* Global Text Colors */
p, li, span, div { color: #334155; }
h1, h2, h3, h4, h5, h6 { color: #0F172A !important; font-weight: 700 !important; }

/* Scrollbar สวยๆ */
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #94A3B8; }
</style>
""", unsafe_allow_html=True)

# ==================================================
# DATA PROCESSING FUNCTIONS
# ==================================================
THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

def extract_year_from_text(text):
    if not text: return ""
    text = str(text)
    match = re.search(r'SCOD\s*(\d{2})', text, re.IGNORECASE)
    if match: return match.group(1)
    match = re.search(r'\((\d{2})\)', text)
    if match: return match.group(1)
    match = re.search(r'ปี\s*(25\d{2}|\d{2})', text)
    if match:
        year_str = match.group(1)
        if len(year_str) == 4: return year_str[-2:]
        return year_str
    match = re.search(r'\s(\d{2})$', text)
    if match: return match.group(1)
    return ""

@st.cache_data(show_spinner=False)
def get_all_excel_data(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    all_data = []
    
    keywords_mapping = {
        "โครงการ": ["โครงการ", "ประเภทงาน", "หมวดงาน"],
        "ชื่องาน": ["ชื่องาน", "ชื่อโครงการย่อย", "รายละเอียดงาน", "รายการ"],
        "พื้นที่": ["พื้นที่", "กฟข.", "เขต"], 
        "จังหวัด": ["จังหวัด"],
        "ผู้รับผิดชอบ": ["ผู้รับผิดชอบ", "ผู้ดำเนินการ"],
        "วงเงิน": ["วงเงินตามอนุมัติ", "วงเงินงบประมาณ", "งบประมาณ", "วงเงิน"],
        "หมายเหตุ": ["อนุมัติ/หมายเหตุ/ปัญหา", "หมายเหตุ", "ปัญหา"],
        "สถานะ": ["สถานะงาน", "ตรวจสอบสถานะงาน", "ตรวจสอบงานแล้วเสร็จ", "สถานะ"]
    }

    for sheet_name in wb.sheetnames:
        if sheet_name == "สรุป": continue
        
        ws = wb[sheet_name]
        col_map = {}
        month_cols = {}
        header_row = 1
        
        for r in range(1, 15):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(r, c).value
                if isinstance(val, str):
                    val_str = val.strip()
                    val_lower = val_str.lower()
                    
                    for std_key, aliases in keywords_mapping.items():
                        if std_key not in col_map:
                            for alias in aliases:
                                if alias.lower() in val_lower:
                                    col_map[std_key] = c
                                    header_row = max(header_row, r)
                                    break
                    
                    if val_str in THAI_MONTHS:
                        month_cols[c] = val_str

        if "โครงการ" not in col_map: continue
        
        proj_c = col_map["โครงการ"]
        task_c = col_map.get("ชื่องาน", proj_c + 2) 

        current_project = ""

        for r in range(header_row + 1, ws.max_row + 1):
            has_data = False
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if v is not None and str(v).strip() != "":
                    has_data = True
                    break
            if not has_data: continue

            project_val = ws.cell(r, proj_c).value
            if project_val is not None and str(project_val).strip() != "":
                current_project = str(project_val).strip()
                
            task_val = ws.cell(r, task_c).value
            task_str = str(task_val).strip() if task_val is not None else ""
            
            if not task_str or task_str in THAI_MONTHS or "ไตรมาส" in task_str:
                continue
            
            task_clean = task_str.replace(" ", "")
            proj_clean = current_project.replace(" ", "")
            if task_clean in ["รวม", "รวมทั้งสิ้น", "ยอดรวม", "รวมทั้งหมด"] or proj_clean in ["รวม", "รวมทั้งสิ้น", "ยอดรวม"]:
                continue
            
            year_suffix = extract_year_from_text(current_project)
            if not year_suffix:
                year_suffix = extract_year_from_text(sheet_name)
            
            deadline_str = "-"
            sort_index = 9999 
            
            for c in sorted(month_cols.keys(), reverse=True):
                cell = ws.cell(r, c)
                if cell.fill and cell.fill.patternType == 'solid':
                    color = cell.fill.start_color.rgb
                    if color and str(color) not in ['00000000', 'FFFFFFFF', '000000', 'None']:
                        month_str = month_cols[c]
                        deadline_str = f"{month_str} {year_suffix}".strip() if year_suffix else month_str
                        sort_index = c 
                        break 
            
            row = {
                "แหล่งที่มา (ชีต)": sheet_name,
                "โครงการหลัก": current_project,
                "ชื่องาน": task_str,
                "พื้นที่": str(ws.cell(r, col_map["พื้นที่"]).value or "-").strip() if "พื้นที่" in col_map else "-",
                "จังหวัด": str(ws.cell(r, col_map["จังหวัด"]).value or "-").strip() if "จังหวัด" in col_map else "-",
                "ผู้รับผิดชอบ": str(ws.cell(r, col_map["ผู้รับผิดชอบ"]).value or "-").strip() if "ผู้รับผิดชอบ" in col_map else "-",
                "วงเงิน": ws.cell(r, col_map["วงเงิน"]).value if "วงเงิน" in col_map else 0,
                "หมายเหตุ/ปัญหา": str(ws.cell(r, col_map["หมายเหตุ"]).value or "-").strip() if "หมายเหตุ" in col_map else "-"
            }
            
            raw_status = ws.cell(r, col_map["สถานะ"]).value if "สถานะ" in col_map else None
            row["สถานะดิบ"] = str(raw_status).strip() if raw_status is not None else ""
            row["Deadline"] = deadline_str
            row["Sort_Index"] = sort_index 
            
            all_data.append(row)

    df_combined = pd.DataFrame(all_data)
    
    if not df_combined.empty:
        def map_3_status(val):
            val = str(val).strip().lower()
            if val in ["nan", "none", "", "nat", "-", "0", "0.0"]: return "❌ ยังไม่ได้ดำเนินการ"
            if "ยังไม่ได้" in val or "ยังไม่เริ่ม" in val: return "❌ ยังไม่ได้ดำเนินการ"
            if any(x in val for x in ["1", "1.0", "สำเร็จ", "แล้วเสร็จ", "เสร็จ"]): return "✅ แล้วเสร็จ"
            if any(x in val for x in ["กำลัง", "ระหว่าง", "ดำเนินการ"]): return "⏳ อยู่ระหว่างดำเนินการ"
            return "❌ ยังไม่ได้ดำเนินการ" 
            
        df_combined["สถานะ"] = df_combined["สถานะดิบ"].apply(map_3_status)

    return df_combined

@st.cache_data(show_spinner=False)
def read_summary_sheet(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    if "สรุป" not in wb.sheetnames: return pd.DataFrame()
    ws = wb["สรุป"]
    
    header_row = None
    for r in range(1, 20):
        row_vals = [str(ws.cell(r, c).value).strip() if ws.cell(r, c).value else "" for c in range(1, ws.max_column + 1)]
        if "ประเภทงาน" in row_vals or "โครงการ" in row_vals:
            header_row = r
            break
            
    if not header_row: return pd.DataFrame()

    c_proj, c_total, c_done = None, None, None
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(header_row, c).value).strip()
        if "ประเภทงาน" in val or "โครงการ" in val: c_proj = c
        elif "จำนวน" in val: c_total = c
        elif "แล้วเสร็จ" in val: c_done = c

    data = []
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(r, c_proj).value if c_proj else None
        if not name or "รวม" in str(name): continue
        
        total = ws.cell(r, c_total).value if c_total else 0
        done = ws.cell(r, c_done).value if c_done else 0
        try: total = int(total) if total else 0
        except: total = 0
        try: done = int(done) if done else 0
        except: done = 0

        data.append({
            "โครงการ": str(name).strip(),
            "จำนวนงาน": total,
            "งานแล้วเสร็จ": done,
            "งานรอดำเนินการ": total - done,
            "หมายเหตุ": "" 
        })
    return pd.DataFrame(data)

def render_html_table(df, cols):
    display_df = df[cols].copy()
    if "วงเงิน" in display_df.columns:
        display_df["วงเงิน"] = pd.to_numeric(display_df["วงเงิน"].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        display_df["วงเงิน"] = display_df["วงเงิน"].apply(lambda x: f"{x:,.0f}")
        
    return display_df.to_html(index=False, classes="styled-table", escape=False)

# ==================================================
# UI COMPONENTS
# ==================================================
def create_category_cards(summary_df):
    if summary_df.empty: return
    st.markdown("## 📂 สรุปตามหมวดงาน")
    cols = st.columns(4)

    for i, row in summary_df.iterrows():
        with cols[i % 4]:
            note = row.get("หมายเหตุ", "")
            total = row["จำนวนงาน"]
            done = row["งานแล้วเสร็จ"]
            pending = row["งานรอดำเนินการ"]
            percent = (done / total * 100) if total > 0 else 0
            
            st.markdown(f"""
            <div class="glass-card category-card">
                <div class="category-number">{str(i+1).zfill(2)}</div>
                <div class="category-title">{row["โครงการ"]}</div>
                <div class="category-value"><b>{total}</b> งานทั้งหมด</div>
                <div class="category-value">
                    <span style="color:#0BFF8D; font-weight:700;">✅ เสร็จ {done} งาน ({percent:.1f}%)</span>
                </div>
                <div style="width: 100%; background-color: #E2E8F0; border-radius: 999px; height: 6px; margin: 6px 0 12px 0;">
                    <div style="background-color: #0BFF8D; height: 6px; border-radius: 999px; width: {percent}%;"></div>
                </div>
                <div class="category-value"><span style="color:#D58A19; font-weight:700;">⏳ คงเหลือ {pending} งาน</span></div>
                <div class="category-note">{note}</div>
            </div>
            """, unsafe_allow_html=True)

def create_pea_chart(summary_df):
    if summary_df.empty: return
    def wrap_label(text):
        text = str(text)
        if len(text) > 20:
            if " " in text:
                return "<br>".join(textwrap.wrap(text, width=20))
            elif "/" in text:
                return text.replace("/", "/<br>")
            else:
                mid = len(text) // 2
                return text[:mid] + "<br>" + text[mid:]
        return text

    wrapped_x = summary_df["โครงการ"].apply(wrap_label)

    fig = go.Figure()
    
    fig.add_bar(
        x=wrapped_x, y=summary_df["จำนวนงาน"], name="จำนวนงานทั้งหมด", 
        marker_color="#3730A3", text=summary_df["จำนวนงาน"], textposition="outside", cliponaxis=False,
        marker_line_width=0
    )
    fig.add_bar(
        x=wrapped_x, y=summary_df["งานแล้วเสร็จ"], name="งานแล้วเสร็จ", 
        marker_color="#0BFF8D", text=summary_df["งานแล้วเสร็จ"], textposition="outside", cliponaxis=False,
        marker_line_width=0
    )

    fig.update_layout(
        barmode="group", 
        height=500, 
        paper_bgcolor="rgba(0,0,0,0)", 
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Sarabun, sans-serif", size=14, color="#1E293B"),
        legend=dict(orientation="h", y=1.12, font=dict(size=14, color="#1E293B")), 
        margin=dict(l=20, r=20, t=40, b=80), 
        xaxis=dict(showgrid=False, tickangle=0), 
        yaxis=dict(showgrid=True, gridcolor="#E2E8F0", gridwidth=1)
    )
    st.plotly_chart(fig, use_container_width=True)

# ==================================================
# PAGE 1: UPLOAD PAGE
# ==================================================
if st.session_state.app_state == 'upload':
    st.markdown("""
        <style>
        [data-testid="collapsedControl"], [data-testid="stSidebar"] { display: none !important; }
        
        @keyframes liquidGradient {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        
        [data-testid="stAppViewContainer"] {
            background: linear-gradient(-45deg, #FF9A9E, #FECFEF, #A1C4FD, #E0C3FC) !important;
            background-size: 400% 400% !important;
            animation: liquidGradient 12s ease infinite !important;
        }

        .liquid-header-box {
            background: rgba(255, 255, 255, 0.4);
            backdrop-filter: blur(15px);
            -webkit-backdrop-filter: blur(15px);
            border: 1px solid rgba(255, 255, 255, 0.7);
            border-bottom: 1px solid rgba(255, 255, 255, 0.3);
            border-right: 1px solid rgba(255, 255, 255, 0.3);
            border-radius: 30px;
            padding: 40px;
            box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.15);
            text-align: center;
            margin-bottom: 30px;
        }

        .upload-title {
            color: #4A4A4A;
            font-size: 42px;
            font-weight: 800;
            margin-bottom: 10px;
            text-shadow: 1px 1px 3px rgba(255,255,255,0.7);
        }
        
        .upload-subtitle {
            color: #4A4A4A;
            font-size: 18px;
            font-weight: 500;
        }

        [data-testid="stFileUploader"] {
            background: rgba(255, 255, 255, 0.45) !important;
            backdrop-filter: blur(20px) !important;
            -webkit-backdrop-filter: blur(20px) !important;
            border: 1px solid rgba(255, 255, 255, 0.8) !important;
            border-radius: 24px !important;
            padding: 30px !important;
            box-shadow: 0 8px 32px 0 rgba(100, 100, 100, 0.1) !important;
        }
        
        [data-testid="stFileUploadDropzone"] {
            background: rgba(255, 255, 255, 0.5) !important;
            border: 2px dashed rgba(255, 255, 255, 0.8) !important;
            border-radius: 16px !important;
        }
        [data-testid="stFileUploadDropzone"]:hover {
            background: rgba(255, 255, 255, 0.8) !important;
            border-color: #FFFFFF !important;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div class="liquid-header-box">
            <div class="upload-title">📊 PEA Executive Dashboard</div>
            <div class="upload-subtitle">ระบบรายงานสรุปผลการดำเนินงานแบบเรียลไทม์</div>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("📂 กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้นใช้งาน", type=["xlsx", "xls"])

        if uploaded_file:
            with st.spinner('กำลังประมวลผลข้อมูลและตรวจสอบความถูกต้อง...'):
                summary_df = read_summary_sheet(uploaded_file)
                all_df = get_all_excel_data(uploaded_file)
                
            if summary_df.empty:
                st.error("❌ ไม่สามารถดึงข้อมูลได้ กรุณาตรวจสอบว่ามีชีต 'สรุป' และคอลัมน์ถูกต้อง")
            elif all_df.empty:
                st.warning("❌ ไม่พบข้อมูลรายละเอียดโครงการในชีตอื่นๆ กรุณาตรวจสอบฟอร์แมตตาราง")
            else:
                st.session_state.summary_df = summary_df
                st.session_state.all_df = all_df
                st.session_state.app_state = 'dashboard'
                st.rerun()

# ==================================================
# PAGE 2 & 3: DASHBOARD PAGE
# ==================================================
elif st.session_state.app_state == 'dashboard':
    summary_df = st.session_state.summary_df
    all_df = st.session_state.all_df

    # ---------------------------------------------
    # SIDEBAR
    # ---------------------------------------------
    st.sidebar.markdown("### ⚙️ การจัดการข้อมูล")
    if st.sidebar.button("⬅️ อัปโหลดไฟล์ใหม่", use_container_width=True):
        st.session_state.app_state = 'upload'
        st.session_state.summary_df = None
        st.session_state.all_df = None
        st.rerun()
        
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔎 เจาะลึกรายละเอียดงาน")
    st.sidebar.caption("เลือกระดับหมวดงานเพื่อดูรายละเอียดเชิงลึก")
    
    main_projects = summary_df["โครงการ"].dropna().astype(str).unique().tolist()
    main_projects = [p for p in main_projects if p.strip() and p.lower() != "nan"]
    main_projects.insert(0, "-- ภาพรวมทั้งหมด --")
    
    selected_main_proj = st.sidebar.selectbox("📌 เลือกหมวดงาน:", main_projects)

    # ---------------------------------------------
    # RENDER OVERVIEW (หน้าที่ 2)
    # ---------------------------------------------
    if selected_main_proj == "-- ภาพรวมทั้งหมด --":
        st.markdown("""
        <div class="pea-header">
            <div class="pea-title">Executive Dashboard</div>
            <div class="pea-subtitle">รายงานสรุปผลการดำเนินงานแบบเรียลไทม์</div>
        </div>
        """, unsafe_allow_html=True)

        total_all = summary_df["จำนวนงาน"].sum() if not summary_df.empty else 0
        done_all = summary_df["งานแล้วเสร็จ"].sum() if not summary_df.empty else 0
        
        percent_all = (done_all / total_all * 100) if total_all > 0 else 0
        
        budget_series = all_df["วงเงิน"].astype(str).str.replace(',', '').str.replace(' ', '')
        budget_sum = pd.to_numeric(budget_series, errors='coerce').fillna(0).sum()

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f'''
            <div class="glass-card">
                <div class="kpi-title">📌 ปริมาณงานทั้งหมด</div>
                <div class="kpi-value" style="color: #3730A3;">{total_all:,} <span style="font-size:16px; color:#94A3B8;">งาน</span></div>
            </div>''', unsafe_allow_html=True)
        with col2:
            st.markdown(f'''
            <div class="glass-card">
                <div class="kpi-title">✅ งานแล้วเสร็จ</div>
                <div class="kpi-value" style="color: #0BFF8D;">{done_all:,} <span style="font-size:16px; color:#94A3B8;">งาน</span> <span style="font-size:24px; color:#059669;">({percent_all:.1f}%)</span></div>
                <div style="width: 100%; background-color: #E2E8F0; border-radius: 999px; height: 8px; margin-top: 12px;">
                    <div style="background-color: #0BFF8D; height: 8px; border-radius: 999px; width: {percent_all}%;"></div>
                </div>
            </div>''', unsafe_allow_html=True)
        with col3:
            st.markdown(f'''
            <div class="glass-card">
                <div class="kpi-title">💰 วงเงินก่อสร้างรวม</div>
                <div class="kpi-value" style="color: #D97706;">{budget_sum:,.0f} <span style="font-size:16px; color:#94A3B8;">บาท</span></div>
            </div>''', unsafe_allow_html=True)

        if not summary_df.empty:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            create_pea_chart(summary_df)
            st.markdown('</div>', unsafe_allow_html=True)
            create_category_cards(summary_df)
        else:
            st.warning("⚠️ ไม่พบข้อมูลสรุปโครงการ")

    # ---------------------------------------------
    # RENDER DRILL-DOWN (หน้าที่ 3)
    # ---------------------------------------------
    else:
        st.markdown(f"## 🔎 เจาะลึกรายละเอียดงาน: <span style='color:#3730A3;'>{selected_main_proj}</span>", unsafe_allow_html=True)
        
        search_term = str(selected_main_proj).strip().upper()
        detail_df = all_df[all_df["โครงการหลัก"].str.upper().str.contains(search_term, regex=False)].copy()

        if detail_df.empty:
            st.info(f"ไม่พบข้อมูลย่อยของ '{selected_main_proj}'")
        else:
            detail_df["โครงการหลัก"] = selected_main_proj 

            pie_data = detail_df["สถานะ"].value_counts().reset_index()
            pie_data.columns = ["Status", "Count"]

            color_map = {
                "✅ แล้วเสร็จ": "#0BFF8D",            
                "⏳ อยู่ระหว่างดำเนินการ": "#D58A19", 
                "❌ ยังไม่ได้ดำเนินการ": "#F43F5E"    
            }
            colors = [color_map.get(s, "#94A3B8") for s in pie_data["Status"]]

            sub_budget_series = detail_df["วงเงิน"].astype(str).str.replace(',', '').str.replace(' ', '')
            sub_budget = pd.to_numeric(sub_budget_series, errors='coerce').fillna(0).sum()

            st.markdown(f"<h4 style='color:#475569; margin-top:20px;'>📊 สัดส่วนสถานะงาน และข้อมูลวงเงินเฉพาะโครงการ</h4>", unsafe_allow_html=True)
            
            scol1, scol2, scol3 = st.columns([1, 1, 1.5])
            with scol1:
                total_sub_tasks = len(detail_df)
                st.markdown(f'''
                <div class="glass-card" style="border-left: 6px solid #3730A3;">
                    <p class="kpi-title" style="margin:0;">📌 งานย่อยทั้งหมด</p>
                    <p class="kpi-value" style="color:#0F172A; margin: 5px 0 0 0;">{total_sub_tasks}</p>
                </div>
                ''', unsafe_allow_html=True)
            with scol2:
                done_count = len(detail_df[detail_df["สถานะ"] == "✅ แล้วเสร็จ"])
                st.markdown(f'''
                <div class="glass-card" style="border-left: 6px solid #0BFF8D;">
                    <p class="kpi-title" style="margin:0;">✅ เสร็จสมบูรณ์แล้ว</p>
                    <p class="kpi-value" style="color:#0BFF8D; margin: 5px 0 0 0;">{done_count}</p>
                </div>
                ''', unsafe_allow_html=True)
            with scol3:
                st.markdown(f'''
                <div class="glass-card" style="border-left: 6px solid #D97706;">
                    <p class="kpi-title" style="margin:0;">💰 วงเงินเฉพาะหมวดนี้</p>
                    <p class="kpi-value" style="color:#D97706; margin: 5px 0 0 0;">{sub_budget:,.0f} ฿</p>
                </div>
                ''', unsafe_allow_html=True)

            col_chart, col_stat = st.columns([1, 1.5])
            with col_chart:
                fig_pie = go.Figure(data=[go.Pie(
                    labels=pie_data["Status"],
                    values=pie_data["Count"],
                    hole=0.55, 
                    marker=dict(colors=colors, line=dict(color='#FFFFFF', width=2)), 
                    textinfo='label+percent+value',
                    textposition='outside', 
                    hoverinfo='label+value'
                )])
                
                fig_pie.update_layout(
                    margin=dict(l=30, r=30, t=20, b=20), 
                    height=350,
                    showlegend=True,
                    legend=dict(orientation="h", y=-0.2, font=dict(color="#1E293B")),
                    font=dict(family="Sarabun, sans-serif", size=14, color="#1E293B"),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)"
                )
                st.plotly_chart(fig_pie, use_container_width=True)
                
            with col_stat:
                st.markdown("<br><br>", unsafe_allow_html=True)
                for _, row in pie_data.iterrows():
                    color_hex = color_map.get(row['Status'], "#1E293B")
                    st.markdown(f"**<span style='color:{color_hex}; font-size:18px;'>{row['Status']}</span> :** <span style='font-size:18px; font-weight:bold;'>{row['Count']}</span> งาน", unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 📋 ตารางรายละเอียดงาน (เรียงตามกำหนดแล้วเสร็จ)")
            
            detail_df.rename(columns={"โครงการหลัก": "โครงการ (หมวดหลัก)"}, inplace=True)
            
            filter_col1, filter_col2 = st.columns(2)
            
            area_options = ["ทั้งหมด"] + sorted([str(x) for x in detail_df["พื้นที่"].unique() if pd.notna(x) and str(x).strip() not in ["", "-"]])
            owner_options = ["ทั้งหมด"] + sorted([str(x) for x in detail_df["ผู้รับผิดชอบ"].unique() if pd.notna(x) and str(x).strip() not in ["", "-"]])
            
            with filter_col1:
                selected_table_area = st.selectbox("📍 กรองข้อมูลตามพื้นที่ (เลือกได้ 1 พื้นที่):", options=area_options, index=0)
            with filter_col2:
                selected_table_owner = st.selectbox("👤 กรองข้อมูลตามผู้รับผิดชอบ (เลือกได้ 1 ผู้รับผิดชอบ):", options=owner_options, index=0)
                
            filtered_table_df = detail_df.copy()
            if selected_table_area != "ทั้งหมด":
                filtered_table_df = filtered_table_df[filtered_table_df["พื้นที่"].astype(str) == selected_table_area]
            if selected_table_owner != "ทั้งหมด":
                filtered_table_df = filtered_table_df[filtered_table_df["ผู้รับผิดชอบ"].astype(str) == selected_table_owner]
            
            display_cols = ["โครงการ (หมวดหลัก)", "ชื่องาน", "แหล่งที่มา (ชีต)", "พื้นที่", "จังหวัด", "ผู้รับผิดชอบ", "วงเงิน", "หมายเหตุ/ปัญหา", "Deadline"]

            df_done = filtered_table_df[filtered_table_df["สถานะ"] == "✅ แล้วเสร็จ"].sort_values(by="Sort_Index")
            st.markdown('<div class="table-header-done">✅ งานที่แล้วเสร็จ</div>', unsafe_allow_html=True)
            if not df_done.empty:
                st.markdown(render_html_table(df_done, display_cols), unsafe_allow_html=True)
            else:
                st.info("ไม่มีข้อมูลที่ตรงกับเงื่อนไขในสถานะนี้")

            df_prog = filtered_table_df[filtered_table_df["สถานะ"] == "⏳ อยู่ระหว่างดำเนินการ"].sort_values(by="Sort_Index")
            st.markdown('<div class="table-header-prog">⏳ งานที่อยู่ระหว่างดำเนินการ</div>', unsafe_allow_html=True)
            if not df_prog.empty:
                st.markdown(render_html_table(df_prog, display_cols), unsafe_allow_html=True)
            else:
                st.info("ไม่มีข้อมูลที่ตรงกับเงื่อนไขในสถานะนี้")

            df_not = filtered_table_df[filtered_table_df["สถานะ"] == "❌ ยังไม่ได้ดำเนินการ"].sort_values(by="Sort_Index")
            st.markdown('<div class="table-header-not">❌ งานที่ยังไม่ได้ดำเนินการ</div>', unsafe_allow_html=True)
            if not df_not.empty:
                st.markdown(render_html_table(df_not, display_cols), unsafe_allow_html=True)
            else:
                st.info("ไม่มีข้อมูลที่ตรงกับเงื่อนไขในสถานะนี้")
