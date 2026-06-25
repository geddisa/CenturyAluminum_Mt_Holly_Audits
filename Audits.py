import streamlit as st
import pandas as pd
from datetime import date
import os
import hashlib

# -----------------------------
# PAGE SETUP & MODERN PLATFORM DESIGN
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum - EHSQ Auditing Platform")

st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 32px; font-weight: 700; color: #0F172A; }
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; }
    h1, h2, h3 { color: #0F172A; font-family: 'Inter', system-ui, sans-serif; font-weight: 600; }
    .login-box { padding: 2.5rem; border-radius: 12px; border: 1px solid #E2E8F0; background-color: #FFFFFF; max-width: 480px; margin: 4rem auto; box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.05); }
    .metric-card { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 1.25rem; border-radius: 10px; text-align: center; }
    </style>
""", unsafe_allow_html=True)

# -----------------------------
# 🔒 SECURE USER DATABASE ENGINE 
# -----------------------------
DB_FILE = "users_db.csv"

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def load_users():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE).set_index("email").to_dict(orient="index")
    else:
        default_users = {
            "admin@centuryaluminum.com": {
                "name": "System Administrator",
                "password_hash": hash_password("Century2026!")
            }
        }
        df = pd.DataFrame.from_dict(default_users, orient="index").reset_index().rename(columns={"index": "email"})
        df.to_csv(DB_FILE, index=False)
        return default_users

def save_new_user(email, name, password):
    users = load_users()
    if email in users:
        return False
    new_row = pd.DataFrame([{"email": email, "name": name, "password_hash": hash_password(password)}])
    new_row.to_csv(DB_FILE, mode='a', header=not os.path.exists(DB_FILE), index=False)
    return True

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "auth_user_email" not in st.session_state:
    st.session_state.auth_user_email = ""
if "auth_user_name" not in st.session_state:
    st.session_state.auth_user_name = ""

# -----------------------------
# 🛡️ SYSTEM AUTHENTICATION WALL
# -----------------------------
if not st.session_state.authenticated:
    st.title("🛡️ Identity Access Gateway")
    st.markdown("### Century Aluminum Company — Mt. Holly EHSQ Hub")
    st.markdown("---")
    
    auth_mode = st.radio("Access Type:", ["Sign In to Account", "Register New Profile"], horizontal=True)
    users_database = load_users()

    if "Sign In" in auth_mode:
        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.subheader("Sign In")
        email_input = st.text_input("Corporate Email Address", placeholder="name@centuryaluminum.com")
        password_input = st.text_input("Security Password", type="password")
        
        if st.button("Verify & Open Workspace", use_container_width=True):
            clean_email = email_input.strip().lower()
            if not clean_email.endswith("@centuryaluminum.com"):
                st.error("❌ Access Denied. Requires a verified corporate domain.")
            elif clean_email in users_database and hash_password(password_input) == users_database[clean_email]["password_hash"]:
                st.session_state.authenticated = True
                st.session_state.auth_user_email = clean_email
                st.session_state.auth_user_name = users_database[clean_email]["name"]
                st.success("Identity Confirmed.")
                st.rerun()
            else:
                st.error("❌ Invalid password or user profile matching details.")
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.subheader("Register Profile")
        reg_name = st.text_input("Full Name", placeholder="Alex Smith")
        reg_email = st.text_input("Corporate Email", placeholder="alex.smith@centuryaluminum.com")
        reg_password = st.text_input("Create Password", type="password")
        confirm_password = st.text_input("Confirm Password", type="password")
        
        if st.button("Create System Profile", use_container_width=True):
            clean_reg_email = reg_email.strip().lower()
            if not reg_name.strip() or not clean_reg_email.endswith("@centuryaluminum.com") or len(reg_password) < 6:
                st.error("❌ Registration error: Complete all fields and verify domain.")
            elif reg_password != confirm_password:
                st.error("❌ Password validation match failed.")
            elif clean_reg_email in users_database:
                st.error("❌ This email has already been claimed.")
            else:
                if save_new_user(clean_reg_email, reg_name.strip(), reg_password):
                    st.success("🎉 Registered successfully! Please toggle back to the 'Sign In' screen.")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# -----------------------------
# 📊 AUTOMATED PARSING ENGINE
# -----------------------------
primary_file = "Audit Schedule - Internal - LPA_2.xlsx"
fallback_file = "Audit Schedule - Internal - LPA.xlsx"
excel_file = primary_file if os.path.exists(primary_file) else fallback_file

@st.cache_data
def load_and_transform_schedule(file_path):
    if not os.path.exists(file_path):
        return pd.DataFrame(), []
    
    xls = pd.ExcelFile(file_path)
    all_records = []
    auditors_found = set()
    
    system_blacklist = {
        "week", "sheet", "audit", "shift", "score", "nan", "ppe", "loto", 
        "total", "average", "target", "date", "dept", "department", "mobile", 
        "equipment", "hk", "score", "operational", "summary", "status", "scheduled", "area", "vacation"
    }
    
    for sheet in xls.sheet_names:
        if sheet in ["Jobs and shifts", "Sheet1"]:
            continue
        try:
            df = pd.read_excel(file_path, sheet_name=sheet, header=None)
            current_active_date = "Baseline 2026"
            
            for _, row in df.iterrows():
                if row.dropna().empty: 
                    continue
                
                # Dynamic timestamp/date row detection to ensure every shifting date marker is preserved
                row_str_combined = " ".join([str(val).strip() for val in row.values if pd.notna(val)])
                
                # Matches patterns like '2026-02-01' or date ranges like '3/2/26 - 3/6/26'
                if "2026-" in row_str_combined or "/26" in row_str_combined:
                    for val in row.values:
                        val_s = str(val).strip()
                        if "2026-" in val_s:
                            current_active_date = val_s.split(" ")[0]
                            break
                        elif "/26" in val_s:
                            current_active_date = val_s
                            break
                    continue
                
                # Keep names exactly as they are listed in the sheet verbatim
                raw_name = str(row.iloc[0]).strip()
                if raw_name == "nan" or not raw_name:
                    raw_name = str(row.iloc[1]).strip()
                
                if any(k in raw_name.lower() for k in system_blacklist) or raw_name.isdigit() or len(raw_name) < 3 or raw_name == "nan":
                    continue
                
                auditors_found.add(raw_name)
                
                # Extract and cross-map distinct target columns natively
                for col_idx, val in enumerate(row.iloc[2:]):
                    val_str = str(val).strip()
                    if val_str != "nan" and len(val_str) > 0:
                        area_out = str(row.iloc[1]) if (len(str(row.iloc[1])) > 1 and str(row.iloc[1]) != "nan") else "General Plant boundary"
                        
                        all_records.append({
                            "Scheduled Target Date": current_active_date,
                            "Auditor Name": raw_name,
                            "Department/Area": area_out,
                            "Classification Type": sheet,
                            "Current Assignment / Status": val_str,
                        })
        except:
            pass
            
    return pd.DataFrame(all_records), sorted(list(auditors_found))

excel_records, parsed_names = load_and_transform_schedule(excel_file)

if not parsed_names:
    parsed_names = ["Freddie Gamble", "Anthony Wall", "Tim Kass", "Bryan Profit", "Reggie Coleman", "Miguel Frias"]

# Load persistent user data inputs
if os.path.exists("audit_data.csv"):
    live_records = pd.read_csv("audit_data.csv")
    live_records.columns = ["Scheduled Target Date", "Auditor Name", "Department/Area", "Classification Type", "Current Assignment / Status"]
else:
    live_records = pd.DataFrame(columns=["Scheduled Target Date", "Auditor Name", "Department/Area", "Classification Type", "Current Assignment / Status"])

combined_dataset = pd.concat([live_records, excel_records], ignore_index=True) if not excel_records.empty else live_records

# -----------------------------
# SIDE PANEL USER SYSTEM NAVIGATION
# -----------------------------
st.sidebar.markdown(f"**Logged In As:**\n💡 **{st.session_state.auth_user_name}**\n`{st.session_state.auth_user_email}`")
if st.sidebar.button("Logout of Session", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🗺️ Navigation Hub")
view_mode = st.sidebar.radio("Go To View:", ["📊 Executive Chart Dashboard", "📋 Direct Entry Log"])

# -----------------------------
# VIEW A: CHARTS & METRIC GRIDS LAYOUT
# -----------------------------
if view_mode == "📊 Executive Chart Dashboard":
    st.title("Century Aluminum Corporate Audit Hub")
    st.markdown("Interactive performance overview tracking clean individual assignments and precise time windows.")
    st.markdown("---")
    
    total_audits = len(combined_dataset)
    unique_types = combined_dataset["Classification Type"].nunique() if total_audits > 0 else 0
    
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        st.markdown(f'<div class="metric-card">📋 <span style="color:#64748B;">Total Assigned Records</span><h3>{total_audits:,}</h3></div>', unsafe_allow_html=True)
    with m_col2:
        st.markdown(f'<div class="metric-card">🏗️ <span style="color:#64748B;">Monitored Categories</span><h3>{unique_types}</h3></div>', unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    if total_audits > 0:
        st.subheader("📈 Audits Tracked by Program Category")
        chart_data = combined_dataset.groupby("Classification Type").size().reset_index(name="Total Count")
        chart_data.columns = ["Audit Classification Type", "Total Active Assignments"]
        st.bar_chart(data=chart_data, x="Audit Classification Type", y="Total Active Assignments", use_container_width=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📋 Master Compliance Schedule Ledger")
    
    all_categories = list(combined_dataset["Classification Type"].unique()) if total_audits > 0 else ["No Active Logs Found"]
    selected_tab_category = st.selectbox("Filter Ledger View Framework:", ["View All Unified Rows"] + all_categories)
    
    display_filter_df = combined_dataset.copy()
    if selected_tab_category != "View All Unified Rows":
        display_filter_df = display_filter_df[display_filter_df["Classification Type"] == selected_tab_category]
        
    if not display_filter_df.empty:
        st.dataframe(display_filter_df[["Scheduled Target Date", "Auditor Name", "Department/Area", "Classification Type", "Current Assignment / Status"]], use_container_width=True, height=500)
    else:
        st.info("No tracked audits recorded in this category yet.")

# -----------------------------
# VIEW B: USER ENTRY WEB INTERFACE
# -----------------------------
elif view_mode == "📋 Direct Entry Log":
    st.title("Log Completed Safety Actions")
    st.markdown("Bypass spreadsheet cells by executing precise direct entries below.")
    st.markdown("---")
    
    with st.form("web_entry_form", clear_on_submit=True):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            in_auditor = st.selectbox("Select Certified Auditor Name", parsed_names)
            in_area = st.text_input("Department / Area Name", placeholder="e.g., Maintenance, Carbon, Cast House")
            in_date = st.date_input("Exact Execution Date", value=date.today())
        with col_f2:
            in_type = st.selectbox("Audit Program Classification System", ["LPA", "Safe Obs - GS and EHS", "Safe Obs - Leadership", "PPE", "LOTO", "Mobile Equip", "HK Scores"])
            in_status = st.text_input("Assignment Status / Location Note", placeholder="e.g., Complete, Crane Shop, 100%")
            
        if st.form_submit_button("Securely Write Entry to Audit Files", type="primary"):
            new_audit_record = pd.DataFrame([{
                "Scheduled Target Date": str(in_date),
                "Auditor Name": in_auditor,
                "Department/Area": in_area if in_area.strip() else "General Plant boundary",
                "Classification Type": in_type,
                "Current Assignment / Status": in_status if in_status.strip() else "Complete"
            }])
            
            if os.path.exists("audit_data.csv"):
                base_df = pd.read_csv("audit_data.csv")
                base_df.columns = ["Scheduled Target Date", "Auditor Name", "Department/Area", "Classification Type", "Current Assignment / Status"]
                pd.concat([base_df, new_audit_record], ignore_index=True).to_csv("audit_data.csv", index=False)
            else:
                new_audit_record.to_csv("audit_data.csv", index=False)
                
            st.success("🎉 Entry registered successfully! Check the main dashboard ledger to review updates.")
            st.rerun()
