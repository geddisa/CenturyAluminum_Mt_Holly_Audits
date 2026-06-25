import streamlit as st
import pandas as pd
from datetime import date
import os
import hashlib

# -----------------------------
# PAGE SETUP & MODERN PLATFORM DESIGN
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum - EHSQ Auditing Platform")

# Injecting modern web-app component styles (Overrides spreadsheet look with card components)
st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 32px; font-weight: 700; color: #0F172A; }
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; }
    h1, h2, h3 { color: #0F172A; font-family: 'Inter', system-ui, sans-serif; font-weight: 600; }
    .login-box { padding: 2.5rem; border-radius: 12px; border: 1px solid #E2E8F0; background-color: #FFFFFF; max-width: 480px; margin: 4rem auto; box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.05); }
    .metric-card { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 1.25rem; border-radius: 10px; text-align: center; }
    .audit-badge { background-color: #EFF6FF; color: #1E40AF; padding: 0.35rem 0.75rem; border-radius: 9999px; font-weight: 600; font-size: 0.85rem; }
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
# 📊 AUTOMATED PARSING ENGINE (Schedules to Chart Objects)
# -----------------------------
excel_file = "Audit Schedule - Internal - LPA.xlsx"

@st.cache_data
def load_and_transform_schedule(file_path):
    if not os.path.exists(file_path):
        return pd.DataFrame(), []
    
    xls = pd.ExcelFile(file_path)
    all_records = []
    auditors_found = set()
    
    # Specific layout tab-parsers mapping spreadsheet formats into web metrics
    for sheet in xls.sheet_names:
        if sheet in ["Jobs and shifts", "Sheet1"]:
            continue
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            for _, row in df.iterrows():
                if row.dropna().empty: continue
                
                raw_name = str(row.iloc[0]).strip()
                if ',' in raw_name:
                    p = raw_name.split(',')
                    raw_name = f"{p[1].strip()} {p[0].strip()}"
                
                if any(k in raw_name.lower() for k in ["week", "sheet", "audit", "shift", "score", "nan", "ppe", "loto"]):
                    continue
                
                if len(raw_name) > 2 and not raw_name.isdigit():
                    auditors_found.add(raw_name)
                    
                    # Transform standard structural tracking indicators safely
                    for val in row.iloc[2:]:
                        v_str = str(val).strip().upper()
                        if v_str in ['C', 'X', '100', '88.89', '90', '95', '75', '86', '96', '64'] or any(char.isdigit() for char in v_str):
                            score_out = "100%" if v_str in ['C', 'X', 'COMPLETE'] else (f"{v_str}%" if "%" not in v_str else v_str)
                            area_out = str(row.iloc[1]) if len(str(row.iloc[1])) > 3 else "General Facility Floor"
                            
                            all_records.append({
                                "Target Period": "Spreadsheet Baseline Schedule",
                                "Auditor": raw_name,
                                "Operational Area": area_out,
                                "Classification Type": sheet,
                                "Current Status / Score": score_out,
                            })
        except:
            pass
    return pd.DataFrame(all_records), sorted(list(auditors_found))

excel_records, parsed_names = load_and_transform_schedule(excel_file)

if not parsed_names:
    parsed_names = ["Anthony Wall", "Art DiFilippo", "Brett Meyer", "Brian Weatherford", "Bryan Profit", "Freddie Gamble", "Tim Kass"]

# Load up persistent web inputs 
if os.path.exists("audit_data.csv"):
    live_records = pd.read_csv("audit_data.csv")
    live_records.columns = ["Target Period", "Auditor", "Operational Area", "Classification Type", "Current Status / Score"]
else:
    live_records = pd.DataFrame(columns=["Target Period", "Auditor", "Operational Area", "Classification Type", "Current Status / Score"])

# Merge data streams seamlessly
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
    st.markdown("Interactive performance overview compiled directly from live operational sheets.")
    st.markdown("---")
    
    # 🌟 METRICS HEADER TILES
    total_audits = len(combined_dataset)
    unique_types = combined_dataset["Classification Type"].nunique() if total_audits > 0 else 0
    
    def calculate_numeric_avg(df):
        if df.empty: return 100.0
        scores = []
        for s in df["Current Status / Score"]:
            s_clean = str(s).replace('%', '').strip().upper()
            if s_clean in ['C', 'X', 'COMPLETE', 'N/A', '']:
                scores.append(100.0)
            else:
                try: scores.append(float(s_clean))
                except: scores.append(100.0)
        return round(sum(scores) / len(scores), 1) if scores else 100.0

    system_compliance = calculate_numeric_avg(combined_dataset)
    
    m_col1, m_col2, m_col3 = st.columns(3)
    with m_col1:
        st.markdown(f'<div class="metric-card">📋 <span style="color:#64748B;">Total Records Unified</span><h3>{total_audits:,}</h3></div>', unsafe_allow_html=True)
    with m_col2:
        st.markdown(f'<div class="metric-card">🎯 <span style="color:#64748B;">Compliance System Average</span><h3>{system_compliance}%</h3></div>', unsafe_allow_html=True)
    with m_col3:
        st.markdown(f'<div class="metric-card">🏗️ <span style="color:#64748B;">Monitored Categories</span><h3>{unique_types}</h3></div>', unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 📈 PERFORMANCE METRIC CHART
    if total_audits > 0:
        st.subheader("📈 Category Performance & Compliance Benchmarks")
        chart_prep = combined_dataset.copy()
        chart_prep["Score_Numeric"] = [float(str(s).replace('%','').strip()) if str(s).replace('%','').strip().replace('.','',1).isdigit() else 100.0 for s in chart_prep["Current Status / Score"]]
        chart_data = chart_prep.groupby("Classification Type")["Score_Numeric"].mean().reset_index()
        chart_data.columns = ["Audit Classification Type", "Compliance Score Rating (%)"]
        st.bar_chart(data=chart_data, x="Audit Classification Type", y="Compliance Score Rating (%)", use_container_width=True)
        
    # 🗂️ CLEAN CATEGORIZED TABS (Instead of big messy rows)
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📋 Dynamic Data Classification Ledger")
    
    all_categories = list(combined_dataset["Classification Type"].unique()) if total_audits > 0 else ["No Active Logs Found"]
    selected_tab_category = st.selectbox("Select Classification View Framework:", ["View All Unified Rows"] + all_categories)
    
    display_filter_df = combined_dataset.copy()
    if selected_tab_category != "View All Unified Rows":
        display_filter_df = display_filter_df[display_filter_df["Classification Type"] == selected_tab_category]
        
    if not display_filter_df.empty:
        st.dataframe(display_filter_df, use_container_width=True, height=380)
    else:
        st.info("No tracked audits recorded in this category yet.")

# -----------------------------
# VIEW B: USER ENTRY WEB INTERFACE
# -----------------------------
elif view_mode == "📋 Direct Entry Log":
    st.title("Log Completed Safety Actions")
    st.markdown("Bypass complicated spreadsheet cells completely by executing data inputs below.")
    st.markdown("---")
    
    with st.form("web_entry_form", clear_on_submit=True):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            in_auditor = st.selectbox("Select Certified Auditor Name", parsed_names)
            in_area = st.selectbox("Target Plant Boundary / Area", ["Maintenance Area", "Carbon Floor", "Cast House", "Potline Grid", "Environmental Management"])
            in_date = st.date_input("Audit Action Execution Date", value=date.today())
        with col_f2:
            in_type = st.selectbox("Audit Program Classification System", ["LPA", "Safe Obs - GS and EHS", "Safe Obs - Leadership", "PPE", "LOTO", "Mobile Equip", "HK Scores"])
            is_c_flag = st.checkbox("Mark execution as Clean / 100% Complete status ('C' Flag indicator)")
            in_score = st.number_input("If numeric evaluation, enter exact score (%)", min_value=0.0, max_value=100.0, value=100.0, step=1.0, disabled=is_c_flag)
            
        if st.form_submit_button("Securely Write Entry to Audit Files", type="primary"):
            final_status_metric = "100%" if is_c_flag else f"{in_score}%"
            new_audit_record = pd.DataFrame([{
                "Target Period": str(in_date),
                "Auditor": in_auditor,
                "Operational Area": in_area,
                "Classification Type": in_type,
                "Current Status / Score": final_status_metric
            }])
            
            if os.path.exists("audit_data.csv"):
                base_df = pd.read_csv("audit_data.csv")
                base_df.columns = ["Target Period", "Auditor", "Operational Area", "Classification Type", "Current Status / Score"]
                pd.concat([base_df, new_audit_record], ignore_index=True).to_csv("audit_data.csv", index=False)
            else:
                new_audit_record.to_csv("audit_data.csv", index=False)
                
            st.success("🎉 Performance metric registered safely! Reloading active metrics engine dashboard views.")
            st.rerun()
