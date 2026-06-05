import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl
import textwrap

# ==================================================
# PAGE CONFIG
# ==================================================
st.set_page_config(
    page_title="PEA Executive Dashboard",
    layout="wide"
)

# ==================================================
# CSS (CUSTOM COLOR THEME)
# ==================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"]  {
    font-family: 'Sarabun', sans-serif;
}

/* พื้นหลังหลักใช้สี #F5DEB3 (Wheat) */
.main { background-color: #F5DEB3; }
.block-container { padding-top: 2rem; padding-bottom: 2rem; }

/* HEADER: ไล่ระดับสี #483D8B -> #660099 -> #5F9EA0 */
.pea-header {
    background: linear-gradient(135deg, #483D8B, #660099, #5F9EA0); 
    color: white;
    padding: 30px;
    border-radius: 16px;
    margin-bottom: 30px;
    box-shadow: 0 10px 30px rgba(72, 61, 139, 0.2);
}
.pea-title { font-size: 34px; font-weight: 700; margin-bottom: 8px;}
.pea-subtitle { font-size: 16px; font-weight: 300; opacity: 0.9; letter-spacing: 0.5px;}

/* KPI Cards */
.kpi-card {
    background: white;
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 20px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    border: 1px solid #EBEBEB;
    transition: transform 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}
.kpi-card:hover {
    transform: translateY(-5px);
    box-shadow: 0 8px 25px rgba(102, 0, 153, 0.15);
}
.kpi-title { font-size: 16px; color: #483D8B; margin-bottom: 5px; font-weight: 600;}
.kpi-value { font-size: 32px; font-weight: 700; }

/* Category Cards */
.category-card {
    background: white;
    border-radius: 16px;
    padding: 24px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    margin-bottom: 20px;
    min-height: 200px;
    border: 1px solid #EBEBEB;
    transition: transform 0.2s ease;
}
.category-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 25px rgba(102, 0, 153, 0.15);
    border-color: #660099;
}
.category-number { color: #660099; font-size: 24px; font-weight: 700; opacity: 0.6;}
.category-title { font-size: 18px; font-weight: 700; margin: 10px 0; color: #483D8B; line-height: 1.4;}
.category-value { font-size: 15px; margin-bottom: 6px; color: #555;}
.category-note { color: #5F9EA0; margin-top: 12px; font-size: 13px; border-top: 1px dashed #EEE; padding-top: 10px; font-weight: 500;}

/* Table Section Headers */
.table-header-done { color: MediumSeaGreen; font-size: 18px; font-weight: 700; margin-top: 25px; border-bottom: 2px solid MediumSeaGreen; padding-bottom: 8px;}
.table-header-prog { color: #5F9EA0; font-size: 18px; font-weight: 700; margin-top: 25px; border-bottom: 2px solid #5F9EA0; padding-bottom: 8px;}
.table-header-not  { color: #660099; font-size: 18px; font-weight: 700; margin-top: 25px; border-bottom: 2px solid #660099; padding-bottom: 8px;}

/* Sub-KPI (Drilldown) */
.sub-kpi {
    background: white;
    padding: 20px 24px;
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    margin-bottom: 20px;
    border: 1px solid #F0F0F0;
}
.sub-kpi-title { font-size: 14px; color: #483D8B; margin: 0; font-weight: 600;}
.sub-kpi-val { font-size: 26px; font-weight: 700; margin: 5px 0 0 0; }
</style>
""", unsafe_allow_html=True)

# ==================================================
# DATA PROCESSING FUNCTIONS
# ==================================================
THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

@st.cache_data
def get_all_excel_data(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    all_data = []
    
    keywords_mapping = {
        "โครงการ": ["โครงการ", "ประเภทงาน", "หมวดงาน"],
        "ชื่องาน": ["ชื่องาน", "ชื่อโครงการย่อย", "รายละเอียดงาน", "รายการ"],
        "พื้นที่": ["พื้นที่", "กฟข.", "เขต", "จังหวัด"],
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
            
            deadline_str = "-"
            sort_index = 99999 
            
            for c in sorted(month_cols.keys(), reverse=True):
                cell = ws.cell(r, c)
                if cell.fill and cell.fill.patternType == 'solid':
                    color = cell.fill.start_color.rgb
                    if color and str(color) not in ['00000000', 'FFFFFFFF', '000000', 'None']:
                        deadline_str = month_cols[c]
                        sort_index = c 
                        break 
            
            row = {
                "แหล่งที่มา (ชีต)": sheet_name,
                "โครงการหลัก": current_project,
                "ชื่องาน": task_str,
                "พื้นที่": str(ws.cell(r, col_map["พื้นที่"]).value or "-").strip() if "พื้นที่" in col_map else "-",
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

@st.cache_data
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

# ==================================================
# UI COMPONENTS
# ==================================================
def create_category_cards(summary_df):
    st.markdown("## 📂 สรุปตามหมวดงาน")
    cols = st.columns(4)

    for i, row in summary_df.iterrows():
        with cols[i % 4]:
            note = row.get("หมายเหตุ", "")
            st.markdown(f"""
            <div class="category-card">
                <div class="category-number">[{i+1}]</div>
                <div class="category-title">📁 {row["โครงการ"]}</div>
                <div class="category-value"><b>{row["จำนวนงาน"]}</b> งานทั้งหมด</div>
                <div class="category-value"><span style="color:MediumSeaGreen; font-weight:600;">✅ เสร็จ {row["งานแล้วเสร็จ"]} งาน</span></div>
                <div class="category-value"><span style="color:#5F9EA0; font-weight:600;">⏳ คงเหลือ {row["งานรอดำเนินการ"]} งาน</span></div>
                <div class="category-note">{note}</div>
            </div>
            """, unsafe_allow_html=True)

def create_pea_chart(summary_df):
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
        marker_color="#483D8B", text=summary_df["จำนวนงาน"], textposition="outside", cliponaxis=False
    )
    fig.add_bar(
        x=wrapped_x, y=summary_df["งานแล้วเสร็จ"], name="งานแล้วเสร็จ", 
        marker_color="MediumSeaGreen", text=summary_df["งานแล้วเสร็จ"], textposition="outside", cliponaxis=False
    )

    fig.update_layout(
        barmode="group", 
        height=550, 
        paper_bgcolor="rgba(0,0,0,0)", # พื้นหลังโปร่งใสเพื่อให้กลืนกับ #F5DEB3
        plot_bgcolor="rgba(255,255,255,0.6)", # พื้นหลังกราฟสีขาวขุ่น
        font=dict(family="Sarabun, sans-serif", size=14, color="#483D8B"),
        legend=dict(orientation="h", y=1.12, font=dict(size=14)), 
        margin=dict(l=20, r=20, t=40, b=80), 
        xaxis=dict(showgrid=False, tickangle=0), 
        yaxis=dict(showgrid=True, gridcolor="#EBEBEB")
    )
    st.plotly_chart(fig, use_container_width=True)

# ==================================================
# MAIN APP
# ==================================================
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel แผนดำเนินการ (ระบบจะอ่านรวมอัตโนมัติ)", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner('กำลังประมวลผลข้อมูลและตรวจสอบความถูกต้อง...'):
        summary_df = read_summary_sheet(uploaded_file)
        all_df = get_all_excel_data(uploaded_file)
        
    if summary_df.empty:
        st.error("❌ ไม่สามารถดึงข้อมูลได้ กรุณาตรวจสอบว่ามีชีต 'สรุป' และคอลัมน์ถูกต้อง")
        st.stop()

    if all_df.empty:
        st.warning("❌ ไม่พบข้อมูลรายละเอียดโครงการในชีตอื่นๆ กรุณาตรวจสอบฟอร์แมตตาราง")
        st.stop()

    # ---------------------------------------------
    # เมนูเจาะลึกรายละเอียดงานใน Sidebar 
    # ---------------------------------------------
    st.sidebar.markdown("### 🔎 เจาะลึกรายละเอียดงาน")
    st.sidebar.caption("ระบบรวบรวมข้อมูลจากทุกชีตอัตโนมัติ")
    
    main_projects = summary_df["โครงการ"].dropna().astype(str).unique().tolist()
    main_projects = [p for p in main_projects if p.strip() and p.lower() != "nan"]
    main_projects.insert(0, "-- กรุณาเลือกหมวดงาน --")
    
    selected_main_proj = st.sidebar.selectbox("📌 เลือกหมวดงาน (ค้นหาข้ามชีต):", main_projects)

    # ==================================================
    # HEADER & OVERVIEW
    # ==================================================
    st.markdown("""
    <div class="pea-header">
        <div class="pea-title">📊 Dashboard แผนดำเนินการ กรพ.</div>
        <div class="pea-subtitle">Executive Progress Report | ภาพรวมโครงการทั้งหมดทั่วประเทศ</div>
    </div>
    """, unsafe_allow_html=True)

    total_all = summary_df["จำนวนงาน"].sum()
    done_all = summary_df["งานแล้วเสร็จ"].sum()
    
    budget_series = all_df["วงเงิน"].astype(str).str.replace(',', '').str.replace(' ', '')
    budget_sum = pd.to_numeric(budget_series, errors='coerce').fillna(0).sum()

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">📌 ปริมาณงานทั้งหมด (ทุกชีต)</div><div class="kpi-value" style="color:#483D8B;">{total_all:,}</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">✅ งานแล้วเสร็จ (ภาพรวม)</div><div class="kpi-value" style="color:MediumSeaGreen;">{done_all:,}</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">💰 วงเงินก่อสร้างรวม (ทั่วประเทศ)</div><div class="kpi-value" style="color:#5F9EA0;">{budget_sum:,.0f} ฿</div></div>', unsafe_allow_html=True)

    # Main Chart
    create_pea_chart(summary_df)

    st.markdown("---")
    create_category_cards(summary_df)
    st.markdown("---")

    # ==================================================
    # DRILL-DOWN SECTION 
    # ==================================================
    if selected_main_proj != "-- กรุณาเลือกหมวดงาน --":
        
        st.markdown(f"## 🔎 เจาะลึกรายละเอียดงาน: <span style='color:#660099;'>{selected_main_proj}</span>", unsafe_allow_html=True)
        
        search_term = str(selected_main_proj).strip().upper()
        detail_df = all_df[all_df["โครงการหลัก"].str.upper().str.contains(search_term, regex=False)].copy()

        if detail_df.empty:
            st.info(f"ไม่พบข้อมูลย่อยของ '{selected_main_proj}'")
        else:
            detail_df["โครงการหลัก"] = selected_main_proj 

            pie_data = detail_df["สถานะ"].value_counts().reset_index()
            pie_data.columns = ["Status", "Count"]

            # โทนสีตามที่ Request มา
            color_map = {
                "✅ แล้วเสร็จ": "MediumSeaGreen",            
                "⏳ อยู่ระหว่างดำเนินการ": "#5F9EA0", 
                "❌ ยังไม่ได้ดำเนินการ": "#660099"    
            }
            colors = [color_map.get(s, "#483D8B") for s in pie_data["Status"]]

            sub_budget_series = detail_df["วงเงิน"].astype(str).str.replace(',', '').str.replace(' ', '')
            sub_budget = pd.to_numeric(sub_budget_series, errors='coerce').fillna(0).sum()

            st.markdown(f"#### 📊 สัดส่วนสถานะงาน และข้อมูลวงเงินเฉพาะโครงการ")
            
            scol1, scol2, scol3 = st.columns([1, 1, 1.5])
            with scol1:
                st.markdown(f'''
                <div class="sub-kpi" style="border-left: 5px solid #483D8B;">
                    <p class="sub-kpi-title">📌 งานย่อยทั้งหมด</p>
                    <p class="sub-kpi-val" style="color:#483D8B;">{len(detail_df)} งาน</p>
                </div>
                ''', unsafe_allow_html=True)
            with scol2:
                done_count = len(detail_df[detail_df["สถานะ"] == "✅ แล้วเสร็จ"])
                st.markdown(f'''
                <div class="sub-kpi" style="border-left: 5px solid MediumSeaGreen;">
                    <p class="sub-kpi-title">✅ เสร็จสมบูรณ์แล้ว</p>
                    <p class="sub-kpi-val" style="color:MediumSeaGreen;">{done_count} งาน</p>
                </div>
                ''', unsafe_allow_html=True)
            with scol3:
                st.markdown(f'''
                <div class="sub-kpi" style="border-left: 5px solid #5F9EA0;">
                    <p class="sub-kpi-title">💰 วงเงินเฉพาะหมวดนี้</p>
                    <p class="sub-kpi-val" style="color:#5F9EA0;">{sub_budget:,.0f} ฿</p>
                </div>
                ''', unsafe_allow_html=True)

            col_chart, col_stat = st.columns([1, 1.5])
            with col_chart:
                fig_pie = go.Figure(data=[go.Pie(
                    labels=pie_data["Status"],
                    values=pie_data["Count"],
                    hole=0.45,
                    marker=dict(colors=colors),
                    textinfo='label+percent+value',
                    textposition='outside', 
                    hoverinfo='label+value'
                )])
                fig_pie.update_layout(
                    margin=dict(l=30, r=30, t=20, b=20), 
                    height=350,
                    showlegend=True,
                    legend=dict(orientation="h", y=-0.2),
                    font=dict(family="Sarabun, sans-serif", size=14, color="#483D8B"),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)"
                )
                st.plotly_chart(fig_pie, use_container_width=True)
                
            with col_stat:
                st.markdown("<br>", unsafe_allow_html=True)
                for _, row in pie_data.iterrows():
                    color_hex = color_map.get(row['Status'], "#483D8B")
                    st.markdown(f"**<span style='color:{color_hex};'>{row['Status']}</span> :** {row['Count']} งาน", unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 📋 ตารางรายละเอียดงาน (เรียงตามกำหนดแล้วเสร็จ)")
            
            detail_df.rename(columns={"โครงการหลัก": "โครงการ (หมวดหลัก)"}, inplace=True)
            display_cols = ["โครงการ (หมวดหลัก)", "ชื่องาน", "แหล่งที่มา (ชีต)", "พื้นที่", "ผู้รับผิดชอบ", "วงเงิน", "หมายเหตุ/ปัญหา", "Deadline"]

            df_done = detail_df[detail_df["สถานะ"] == "✅ แล้วเสร็จ"].sort_values(by="Sort_Index")
            st.markdown('<div class="table-header-done">✅ งานที่แล้วเสร็จ</div>', unsafe_allow_html=True)
            if not df_done.empty:
                st.dataframe(df_done[display_cols], use_container_width=True, hide_index=True)
            else:
                st.info("ไม่มีงานที่แล้วเสร็จในหมวดนี้")

            df_prog = detail_df[detail_df["สถานะ"] == "⏳ อยู่ระหว่างดำเนินการ"].sort_values(by="Sort_Index")
            st.markdown('<div class="table-header-prog">⏳ งานที่อยู่ระหว่างดำเนินการ</div>', unsafe_allow_html=True)
            if not df_prog.empty:
                st.dataframe(df_prog[display_cols], use_container_width=True, hide_index=True)
            else:
                st.info("ไม่มีงานที่กำลังดำเนินการในหมวดนี้")

            df_not = detail_df[detail_df["สถานะ"] == "❌ ยังไม่ได้ดำเนินการ"].sort_values(by="Sort_Index")
            st.markdown('<div class="table-header-not">❌ งานที่ยังไม่ได้ดำเนินการ</div>', unsafe_allow_html=True)
            if not df_not.empty:
                st.dataframe(df_not[display_cols], use_container_width=True, hide_index=True)
            else:
                st.info("ไม่มีงานค้างที่ยังไม่ดำเนินการ")
