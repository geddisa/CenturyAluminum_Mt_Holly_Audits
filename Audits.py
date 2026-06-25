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
# ✅ EMPLOYEE REGISTRY FILTER
# -----------------------------
all_names = set()
COMMON_NON_NAMES = [
    "Room", "West", "East", "Side", "Shop", "Tank", "Tanks", "Mill", 
    "House", "Laboratory", "Moldshop", "EQUIP", "Carbon", "Potline", 
    "Area", "Café", "Snacks", "Shift", "Job", "Details"
]
EXPLICIT_NON_NAMES = {
    "CH Laboratory", "CH Moldshop", "Cast House", "HK Scores", 
    "MOBILE EQUIP", "MT Carbon", "MT Potline", "Maintenance Area", 
    "Midnight Snacks", "Monmartre Café", "Pitch Tanks", "City Center"
}

for sheet in xls.sheet_names:
    if sheet in ["Jobs and shifts", "Sheet1"]:
        continue
    try:
        df_temp = pd.read_excel(excel_file, sheet_name=sheet)
        for col in df_temp.columns:
            for val in df_temp[col].dropna():
                if isinstance(val, str):
                    val = val.strip().replace('"', '')
                    if "," in val:  # Fix Last, First formats
                        parts = [p.strip() for p in val.split(",")]
                        if len(parts) == 2:
                            val = f"{parts[1]} {parts[0]}"
                    words = val.split()
                    if (
                        len(words) == 2 
                        and all(w[0].isupper() for w in words if w.isalpha()) 
                        and not any(w in COMMON_NON_NAMES for w in words)
                        and val not in EXPLICIT_NON_NAMES
                    ):
                        all_names.add(val)
    except:
        pass

all_names.add("Wilson Smith")
name_list = sorted(list(all_names))

# -----------------------------
# ⚙️ FULL EXCEL HISTORY PARSER
# -----------------------------
@st.cache_data
def parse_excel_history():
    compiled_records = []
    
    # Helper to clean text ranges like '6/1/26 - 6/5/26' into a standard start date
    def extract_start_date(date_cell, default_year="2026"):
        if pd.isna(date_cell):
            return None
        match = re.search(r'(\d+)/(\d+)/(\d+)', str(date_cell))
        if match:
            return pd.to_datetime(f"20{match.group(3)}-{match.group(1)}-{match.group(2)}")
        return None

    # 1. Parse HK Scores Tab
    if "HK Scores" in xls.sheet_names:
        df_hk = pd.read_excel(excel_file, sheet_name="HK Scores")
        valid_areas = ["Maintenance", "Carbon", "Cast House", "Potline", "Environmental"]
        
        for col_idx in range(2, len(df_hk.columns)):
            # Look for date patterns in the header rows
            target_date = None
            for r_idx in range(min(12, len(df_hk))):
                d_val = df_hk.iloc[r_idx, col_idx]
                parsed_d = extract_start_date(d_val)
                if parsed_d:
                    target_date = parsed_d
                    break
            if not target_date:
                target_date = pd.to_datetime("2026-06-01")
                
            for row_idx in range(len(df_hk)):
                area_val = str(df_hk.iloc[row_idx, 1]).strip() if pd.notna(df_hk.iloc[row_idx, 1]) else ""
                if area_val in valid_areas:
                    score = df_hk.iloc[row_idx, col_idx]
                    if pd.notna(score) and isinstance(score, (int, float)):
                        compiled_records.append({
                            "Date": target_date,
                            "Auditor": "System Record",
                            "Area": area_val,
                            "Type": "HK Score",
                            "Score": float(score),
                            "Notes": "Historical benchmark score imported from HK Scores tab."
                        })

    # 2. Parse LOTO Tab
    if "LOTO" in xls.sheet_names:
        df_loto = pd.read_excel(excel_file, sheet_name="LOTO")
        for idx, row in df_loto.iterrows():
            auditor_raw = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            area_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # Verify if this row belongs to an actual auditor name
            matched_auditor = next((n for n in name_list if n.lower() in auditor_raw.lower()), None)
            if matched_auditor:
                for col_idx in range(2, len(row)):
                    score_val = row.iloc[col_idx]
                    if pd.notna(score_val) and isinstance(score_val, (int, float)):
                        compiled_records.append({
                            "Date": pd.to_datetime("2026-04-06"),
                            "Auditor": matched_auditor,
                            "Area": area_val if area_val else "General Plant",
                            "Type": "LOTO",
                            "Score": float(score_val),
                            "Notes": "Historical verification loop data parsed from LOTO tab matrix."
                        })

    # 3. Parse Safe Obs tabs
    safe_obs_tabs = ["Safe Obs - Leadership", "Safe Obs - GS and EHS"]
    for tab_name in safe_obs_tabs:
        if tab_name in xls.sheet_names:
            df_sol = pd.read_excel(excel_file, sheet_name=tab_name)
            for idx, row in df_sol.iterrows():
                auditor_raw = str(row.iloc[0]).strip().replace('"', '') if pd.notna(row.iloc[0]) else ""
                area_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                
                if "," in auditor_raw:
                    parts = [p.strip() for p in auditor_raw.split(",")]
                    auditor_raw = f"{parts[1]} {parts[0]}"
                    
                matched_auditor = next((n for n in name_list if n.lower() in auditor_raw.lower()), None)
                if matched_auditor:
                    # Count any cells containing 'c' or 'x' indicators for compliance metrics
                    completions = sum(1 for cell in row.values[2:] if str(cell).strip().lower() in ['c', 'x'])
                    if completions > 0:
                        compiled_records.append({
                            "Date": pd.to_datetime("2026-06-01"),
                            "Auditor": matched_auditor,
                            "Area": "Plant-wide Operations" if area_val.lower() == "any" else area_val,
                            "Type": "Safe Observation",
                            "Score": 100.0,
                            "Notes": f"Logged execution: completed {completions} checks in schedule window."
                        })

    return pd.DataFrame(compiled_records)

# Local transactional DB checks
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
# 📊 DASHBOARD PAGE (ALL PAST SPREADSHEET INFO)
# -----------------------------
if page == "📊 Dashboard":
    st.header("Audit & Performance Dashboard")
    
    if user_data.empty:
        st.warning("No historical spreadsheet logs found.")
    else:
        auditor_filter = st.sidebar.selectbox("Filter Auditor", ["All"] + sorted(list(user_data["Auditor"].unique())), key="dash_auditor_filter")
        area_filter = st.sidebar.selectbox("Filter Plant Area", ["All"] + sorted(list(user_data["Area"].unique())), key="dash_area_filter")
        
        df = user_data.copy()
        if auditor_filter != "All":
            df = df[df["Auditor"] == auditor_filter]
        if area_filter != "All":
            df = df[df["Area"] == area_filter]

        # Key indicators computed over historical inputs
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Historical Records", len(df))
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
        area = st.selectbox("Plant Operational Area", [
            "Maintenance", "Carbon", "Cast House", "Potline", "Environmental",
            "Green Mill", "Coke Tank", "Rod Shop"
        ])
        audit_type = st.selectbox("Audit Type Classification", [
            "LPA", "Safe Observation", "PPE", "LOTO", "Mobile Equipment", "HK Score"
        ])
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
