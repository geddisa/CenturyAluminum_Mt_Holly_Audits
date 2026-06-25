import streamlit as st
import pandas as pd
from datetime import date
import os
import re

# -----------------------------
# PAGE SETUP
# -----------------------------
st.set_page_config(layout="wide", page_title="Audit Management System")

st.markdown("""
    <style>
    .main .block-container { padding-top: 2rem; }
    div[data-testid="stMetricValue"] { font-size: 28px; font-weight: bold; color: #1E3A8A; }
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 4px; }
    h1, h2, h3 { color: #0F172A; font-family: 'Segoe UI', Helvetica, Arial, sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("📌 Audit Management System")

excel_file = "Audit Schedule - Internal - LPA.xlsx"

if not os.path.exists(excel_file):
    st.error("❌ Excel file not found")
    st.write("Files found in folder:", os.listdir())
    st.stop()

xls = pd.ExcelFile(excel_file)

# -----------------------------
# ✅ MASTER NAME REGISTRY
# -----------------------------
all_names = set()
COMMON_NON_NAMES = ["Room", "West", "East", "Side", "Shop", "Tank", "Tanks", "Mill", "House", "Area", "Café", "Snacks", "Shift", "Job", "Details"]

for sheet in xls.sheet_names:
    if sheet in ["Jobs and shifts", "Sheet1"]:
        continue
    try:
        df_temp = pd.read_excel(excel_file, sheet_name=sheet)
        for row in df_temp.itertuples(index=False):
            if len(row) > 0 and pd.notna(row[0]):
                name_str = str(row[0]).strip().replace('"', '')
                if ',' in name_str:
                    parts = [p.strip() for p in name_str.split(',')]
                    if len(parts) == 2:
                        name_str = f"{parts[1]} {parts[0]}"
                words = name_str.split()
                if len(words) >= 2 and words[0][0].isupper() and not any(w in COMMON_NON_NAMES for w in words):
                    all_names.add(name_str)
    except:
        pass

name_list = sorted(list(all_names))

# -----------------------------
# ⚙️ ULTIMATE EXCEL SPREADSHEET PARSER
# -----------------------------
@st.cache_data
def parse_excel_history():
    compiled_records = []
    
    def parse_header_date(header_str):
        if pd.isna(header_str):
            return None
        match = re.search(r'(\d+)/(\d+)/(\d+)', str(header_str))
        if match:
            return pd.to_datetime(f"20{match.group(3)}-{match.group(1)}-{match.group(2)}")
        match_iso = re.search(r'(\d{4})-(\d{2})-(\d{2})', str(header_str))
        if match_iso:
            return pd.to_datetime(match_iso.group(0))
        return None

    # 1. Housekeeping Plan Assignments (HK)
    if "HK" in xls.sheet_names:
        df_hk = pd.read_excel(excel_file, sheet_name="HK")
        dates_by_col = {}
        for col_idx in range(2, len(df_hk.columns)):
            for row_idx in range(min(15, len(df_hk))):
                d = parse_header_date(df_hk.iloc[row_idx, col_idx])
                if d:
                    dates_by_col[col_idx] = d
                    break
        for idx, row in df_hk.iterrows():
            auditor_raw = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            matched_auditor = next((n for n in name_list if n.lower() in auditor_raw.lower()), None)
            if matched_auditor:
                for col_idx in range(2, len(row)):
                    loc = row.iloc[col_idx]
                    if pd.notna(loc) and str(loc).strip() not in ["", "1", "2", "3"]:
                        compiled_records.append({
                            "Date": dates_by_col.get(col_idx, pd.to_datetime("2026-01-01")),
                            "Auditor": matched_auditor,
                            "Area": str(loc).strip(),
                            "Type": "HK Assignment",
                            "Score": 100.0,
                            "Notes": "Scheduled housekeeping inspection loop."
                        })

    # 2. Housekeeping Historical Grades (HK Scores)
    if "HK Scores" in xls.sheet_names:
        df_hks = pd.read_excel(excel_file, sheet_name="HK Scores")
        dates_by_col = {}
        for col_idx in range(2, len(df_hks.columns)):
            for row_idx in range(min(15, len(df_hks))):
                d = parse_header_date(df_hks.iloc[row_idx, col_idx])
                if d:
                    dates_by_col[col_idx] = d
                    break
        for idx, row in df_hks.iterrows():
            area_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            if area_val in ["Maintenance", "Carbon", "Cast House", "Potline", "Environmental"]:
                for col_idx in range(2, len(row)):
                    score = row.iloc[col_idx]
                    if pd.notna(score) and isinstance(score, (int, float)):
                        compiled_records.append({
                            "Date": dates_by_col.get(col_idx, pd.to_datetime("2026-01-01")),
                            "Auditor": "Department Metric",
                            "Area": area_val,
                            "Type": "HK Scorecard Grade",
                            "Score": float(score),
                            "Notes": "Aggregated area housekeeping performance rating."
                        })

    # 3. Lockout / Tagout Matrix (LOTO)
    if "LOTO" in xls.sheet_names:
        df_loto = pd.read_excel(excel_file, sheet_name="LOTO")
        dates_by_col = {}
        for col_idx in range(2, len(df_loto.columns)):
            for row_idx in range(min(15, len(df_loto))):
                d = parse_header_date(df_loto.iloc[row_idx, col_idx])
                if d:
                    dates_by_col[col_idx] = d
                    break
        for idx, row in df_loto.iterrows():
            auditor_raw = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            area_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "General Plant"
            matched_auditor = next((n for n in name_list if n.lower() in auditor_raw.lower()), None)
            if matched_auditor:
                for col_idx in range(2, len(row)):
                    score = row.iloc[col_idx]
                    if pd.notna(score) and isinstance(score, (int, float)):
                        compiled_records.append({
                            "Date": dates_by_col.get(col_idx, pd.to_datetime("2026-04-01")),
                            "Auditor": matched_auditor,
                            "Area": area_val,
                            "Type": "LOTO Audit",
                            "Score": float(score),
                            "Notes": "LOTO compliance validation score entry."
                        })

    # 4. Standard Operational Inspections (PPE & Mobile Equip)
    for sheet_name in ["PPE", "Mobile Equip"]:
        if sheet_name in xls.sheet_names:
            df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
            dates_by_col = {}
            for col_idx in range(2, len(df_sheet.columns)):
                for row_idx in range(min(15, len(df_sheet))):
                    d = parse_header_date(df_sheet.iloc[row_idx, col_idx])
                    if d:
                        dates_by_col[col_idx] = d
                        break
            for idx, row in df_sheet.iterrows():
                area_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "Plant Wide"
                for col_idx in range(2, len(row)):
                    val = row.iloc[col_idx]
                    if pd.notna(val) and str(val).strip() in ["1", "2", "3", "4"]:
                        compiled_records.append({
                            "Date": dates_by_col.get(col_idx, pd.to_datetime("2026-01-01")),
                            "Auditor": "Scheduled Tracker",
                            "Area": area_val,
                            "Type": f"{sheet_name} Track",
                            "Score": 100.0,
                            "Notes": f"Target checks assigned inside row index: {idx}."
                        })

    # 5. Behavior-Based Safety Observations (Leadership, GS and EHS)
    for tab in ["Safe Obs - Leadership", "Safe Obs - GS and EHS", "Safe Obs GS EHS"]:
        if tab in xls.sheet_names:
            df_so = pd.read_excel(excel_file, sheet_name=tab)
            dates_by_col = {}
            for col_idx in range(2, len(df_so.columns)):
                for row_idx in range(min(15, len(df_so))):
                    d = parse_header_date(df_so.iloc[row_idx, col_idx])
                    if d:
                        dates_by_col[col_idx] = d
                        break
            for idx, row in df_so.iterrows():
                auditor_raw = str(row.iloc[0]).strip().replace('"', '') if pd.notna(row.iloc[0]) else ""
                if ',' in auditor_raw:
                    parts = [p.strip() for p in auditor_raw.split(',')]
                    auditor_raw = f"{parts[1]} {parts[0]}"
                area_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "General Plant"
                
                matched_auditor = next((n for n in name_list if n.lower() in auditor_raw.lower()), None)
                if matched_auditor:
                    for col_idx in range(2, len(row)):
                        cell_indicator = str(row.iloc[col_idx]).strip().lower()
                        if cell_indicator in ['c', 'x', 'o', 'ooo']:
                            score_val = 100.0 if cell_indicator in ['c', 'x'] else 0.0
                            note_text = "Completed - Safe" if cell_indicator == 'c' else "Completed - Action Plan Saved" if cell_indicator == 'x' else "Omit / Out of Office"
                            compiled_records.append({
                                "Date": dates_by_col.get(col_idx, pd.to_datetime("2026-06-01")),
                                "Auditor": matched_auditor,
                                "Area": "Plant Wide" if area_val.lower() == "any" else area_val,
                                "Type": "Safety Observation Loop",
                                "Score": score_val,
                                "Notes": f"{note_text} (Tab: {tab})"
                            })
                            
    return pd.DataFrame(compiled_records)

# -----------------------------
# RUN TIME PERSISTENCE ENGINE
# -----------------------------
if os.path.exists("audit_data.csv"):
    user_data = pd.read_csv("audit_data.csv")
    user_data["Date"] = pd.to_datetime(user_data["Date"])
else:
    user_data = parse_excel_history()

def save_user_data(df):
    df.to_csv("audit_data.csv", index=False)

# -----------------------------
# SIDEBAR NAVIGATION
# -----------------------------
st.sidebar.markdown("### 🏢 Facility Management")
page = st.sidebar.radio(
    "Navigation Options",
    ["📊 Dashboard", "📋 Enter Audit", "📁 Excel Viewer", "👥 Names"],
    key="main_navigation_radio"
)

# -----------------------------
# 📊 DASHBOARD PAGE
# -----------------------------
if page == "📊 Dashboard":
    st.header("Audit & Performance Dashboard")
    
    if user_data.empty:
        st.warning("No historical matrix logs compiled yet.")
    else:
        auditor_filter = st.sidebar.selectbox("Filter Auditor", ["All"] + sorted(list(user_data["Auditor"].unique())), key="dash_auditor_filter")
        area_filter = st.sidebar.selectbox("Filter Plant Area", ["All"] + sorted(list(user_data["Area"].unique())), key="dash_area_filter")
        
        df = user_data.copy()
        if auditor_filter != "All":
            df = df[df["Auditor"] == auditor_filter]
        if area_filter != "All":
            df = df[df["Area"] == area_filter]

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Historical Records", f"{len(df):,}")
        col2.metric("Overall Average Rating", f"{round(df['Score'].mean(), 1)}%" if not df['Score'].empty else "N/A")
        col3.metric("Monitored Zones Active", df["Area"].nunique())
        
        st.markdown("---")
        
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            st.subheader("Performance Tracking by Inspector/Zone")
            if not df.empty and df["Score"].notna().any():
                chart_data = df.groupby("Auditor")["Score"].mean().sort_values()
                st.bar_chart(chart_data)
        
        with col_chart2:
            st.subheader("Inspection Trends Over Time")
            if not df.empty:
                trend_df = df.groupby("Date")["Score"].mean().reset_index().set_index("Date")
                st.line_chart(trend_df)

        st.subheader("📋 Historical & Real-Time Consolidated Data Ledger")
        st.dataframe(df.sort_values(by="Date", ascending=False), use_container_width=True)

# -----------------------------
# 📋 ENTER AUDIT FORM
# -----------------------------
elif page == "📋 Enter Audit":
    st.header("Enter New Audit Sheet Records")

    with st.form("audit_form", clear_on_submit=True):
        auditor = st.selectbox("Select Auditor Name", name_list)
        area = st.selectbox("Plant Operational Area", ["Maintenance", "Carbon", "Cast House", "Potline", "Environmental", "Green Mill", "Coke Tank", "Rod Shop"])
        audit_type = st.selectbox("Audit Type Classification", ["LPA", "Safe Observation", "PPE", "LOTO", "Mobile Equipment", "HK Score"])
        score = st.number_input("Recorded Performance Score (%)", 0.0, 100.0, step=1.0)
        audit_date = st.date_input("Audit Execution Date", value=date.today())
        notes = st.text_area("Observations & Notes")

        submitted = st.form_submit_button("Submit Entry to Database")

        if submitted:
            new_row = pd.DataFrame([{
                "Date": pd.to_datetime(audit_date),
                "Auditor": auditor,
                "Area": area,
                "Type": audit_type,
                "Score": score,
                "Notes": notes
            }])
            
            user_data = pd.concat([user_data, new_row], ignore_index=True)
            save_user_data(user_data)
            st.success(f"✅ Audit records saved successfully into ledger database!")

# -----------------------------
# 📁 EXCEL VIEWER
# -----------------------------
elif page == "📁 Excel Viewer":
    st.header("Spreadsheet Tab Visualizer")
    sheet = st.selectbox("Choose Sheet Tab to View", xls.sheet_names, key="excel_viewer_sheet_select")
    
    df_excel = pd.read_excel(excel_file, sheet_name=sheet)
    st.markdown(f"#### 📄 Displaying Tab: `{sheet}`")
    st.dataframe(df_excel, use_container_width=True)

# -----------------------------
# 👥 EMPLOYEE NAMES VERIFICATION
# -----------------------------
elif page == "👥 Names":
    st.header("Verified Clean Auditor Registry")
    st.write(f"Total Unique Filtered Personnel Names Identified: **{len(name_list)}**")
    
    name_df = pd.DataFrame(name_list, columns=["Employee Name Listing"])
    st.dataframe(name_df, use_container_width=True)
