import streamlit as st
import pandas as pd
from datetime import date
import os
import hashlib
import re

# -----------------------------
# PAGE SETUP & CONFIG
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum - Secure Audit System")

# Clean, professional UI stylesheet making it look like an enterprise web app
st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 28px; font-weight: bold; color: #1E3A8A; }
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 8px; box-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1); }
    h1, h2, h3 { color: #0F172A; font-family: 'Segoe UI', Helvetica, Arial, sans-serif; }
    .login-box { padding: 2.5rem; border-radius: 8px; border: 1px solid #CBD5E1; background-color: #F8FAFC; max-width: 480px; margin: 2rem auto; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); }
    .card-container { background-color: #FFFFFF; padding: 1.5rem; border: 1px solid #E2E8F0; border-radius: 8px; margin-bottom: 1rem; }
    .danger-zone { border: 1px solid #FCA5A5; padding: 1.5rem; border-radius: 8px; background-color: #FEF2F2; margin-top: 1.5rem; }
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
# 🛡️ AUTHENTICATION GATEWAY
# -----------------------------
if not st.session_state.authenticated:
    st.title("🔒 Corporate Security Access Gateway")
    st.markdown("### Century Aluminum Company — Internal EHSQ Systems")
    st.markdown("---")
    
    auth_mode = st.radio("Choose Action:", ["Sign In", "Secure Registration (New User)"], horizontal=True)
    users_database = load_users()

    if auth_mode == "Sign In":
        with st.container():
            st.markdown('<div class="login-box">', unsafe_allow_html=True)
            st.subheader("Account Sign In")
            email_input = st.text_input("Corporate Email Address", placeholder="username@centuryaluminum.com", key="login_email")
            password_input = st.text_input("Password", type="password", key="login_pass")
            
            if st.button("Verify Identity & Sign In", use_container_width=True):
                clean_email = email_input.strip().lower()
                if not clean_email.endswith("@centuryaluminum.com"):
                    st.error("❌ Access Denied. Only valid @centuryaluminum.com domains are authorized.")
                elif clean_email in users_database and hash_password(password_input) == users_database[clean_email]["password_hash"]:
                    st.session_state.authenticated = True
                    st.session_state.auth_user_email = clean_email
                    st.session_state.auth_user_name = users_database[clean_email]["name"]
                    st.success("✅ Credentials verified.")
                    st.rerun()
                else:
                    st.error("❌ Invalid credentials entry.")
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        with st.container():
            st.markdown('<div class="login-box">', unsafe_allow_html=True)
            st.subheader("Create Authorized Profile")
            reg_name = st.text_input("Full Name (First Last)", placeholder="John Doe")
            reg_email = st.text_input("Corporate Email Address", placeholder="username@centuryaluminum.com")
            reg_password = st.text_input("Choose Password", type="password")
            confirm_password = st.text_input("Confirm Password", type="password")
            
            if st.button("Register Corporate Profile", use_container_width=True):
                clean_reg_email = reg_email.strip().lower()
                if not reg_name.strip() or not clean_reg_email.endswith("@centuryaluminum.com") or len(reg_password) < 6:
                    st.error("❌ Invalid submission details. Ensure domain is correct and password is 6+ chars.")
                elif reg_password != confirm_password:
                    st.error("❌ Passwords do not match.")
                elif clean_reg_email in users_database:
                    st.error("❌ Email already exists.")
                else:
                    if save_new_user(clean_reg_email, reg_name.strip(), reg_password):
                        st.success("🎉 Registered! Switch to Sign In tab.")
            st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# -----------------------------
# 📊 EXCEL REPOSITORY PARSER ENGINE
# -----------------------------
excel_file = "Audit Schedule - Internal - LPA.xlsx"

@st.cache_data
def parse_all_excel_schedules(file_path):
    if not os.path.exists(file_path):
        return pd.DataFrame(), []
    
    xls = pd.ExcelFile(file_path)
    all_extracted_records = []
    discovered_auditors = set()
    
    # Parser mapping config per tab structure
    for sheet in xls.sheet_names:
        if sheet in ["Jobs and shifts", "Sheet1"]:
            continue
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            
            # Clean and isolate column elements matching standard structured tables
            for idx, row in df.iterrows():
                row_list = [str(x).strip() for x in row.dropna().tolist()]
                if not row_list:
                    continue
                
                # Dynamically extract scores/completions linked to known names
                name_candidate = str(row.iloc[0]).strip()
                if ',' in name_candidate:
                    parts = name_candidate.split(',')
                    name_candidate = f"{parts[1].strip()} {parts[0].strip()}"
                
                # Filter out system header text markers
                if any(k in name_candidate.lower() for k in ["week", "sheet", "audit", "shift", "score", "nan", "mobile equip", "loto", "ppe"]):
                    continue
                
                if len(name_candidate) > 2 and not name_candidate.isdigit():
                    discovered_auditors.add(name_candidate)
                    
                    # Scan row positions for scores or completion 'C' values
                    for cell_val in row.iloc[2:]:
                        val_str = str(cell_val).strip().upper()
                        if val_str in ['C', 'X', '100', '88.89', '90', '95', '75', '86', '96', '64'] or any(v.replace('.','',1).isdigit() for v in [val_str]):
                            score_disp = "C" if val_str in ['C', 'X'] else val_str
                            area_assigned = str(row.iloc[1]) if len(str(row.iloc[1])) > 3 else "Operational Plant Area"
                            
                            all_extracted_records.append({
                                "Date": "Schedule Tracking Period",
                                "Auditor": name_candidate,
                                "Area": area_assigned,
                                "Type": sheet,
                                "Score": score_disp,
                                "Notes": f"Imported directly from legacy tracking grid worksheet tab [{sheet}]."
                            })
        except Exception:
            pass
            
    return pd.DataFrame(all_extracted_records), sorted(list(discovered_auditors))

excel_records, parsed_names = parse_all_excel_schedules(excel_file)

# -----------------------------
# RUN TIME LIVE DATA INTERFACE
# -----------------------------
if os.path.exists("audit_data.csv"):
    live_records = pd.read_csv("audit_data.csv")
else:
    live_records = pd.DataFrame(columns=["Date", "Auditor", "Area", "Type", "Score", "Notes"])

# Blend files with user entered manual overrides cleanly
if not excel_records.empty:
    user_data = pd.concat([live_records, excel_records], ignore_index=True)
else:
    user_data = live_records

def save_user_data(df):
    # Only save locally captured records back to delta storage file
    live_only = df[df["Date"] != "Schedule Tracking Period"]
    live_only.to_csv("audit_data.csv", index=False)

if not parsed_names:
    parsed_names = ["Anthony Wall", "Art DiFilippo", "Brett Meyer", "Brian Weatherford", "Bryan Profit", "Freddie Gamble", "Tim Kass"]

# -----------------------------
# SIDEBAR CONTROL PANEL
# -----------------------------
st.sidebar.markdown(f"👤 **Session Profile:**\n**{st.session_state.auth_user_name}**\n`{st.session_state.auth_user_email}`")
if st.sidebar.button("Secure Logout", use_container_width=True):
    st.session_state.authenticated = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🏢 Facility Management")
page = st.sidebar.radio("Navigation Options", ["📊 Consolidated Charts & Tables", "📋 Enter Audit"])

# -----------------------------
# DYNAMIC NAVIGATION PAGES
# -----------------------------
if page == "📊 Consolidated Charts & Tables":
    st.header("Century Aluminum Performance Dashboard")
    st.markdown("Consolidated real-time views from your audit database log files and uploaded schedule registries.")
    
    # Math execution wrapper transforming 'C' statuses instantly to 100% compliant counts
    total_records = len(user_data)
    if not user_data.empty:
        def parse_score(val):
            if pd.isna(val): return 100.0
            s = str(val).strip().upper()
            if s in ['C', 'X', 'N/A', 'COMPLETE', ''] or 'COMPLETE' in s: return 100.0
            try: return float(val)
            except: return 100.0
            
        clean_scores = user_data['Score'].apply(parse_score)
        avg_score = round(clean_scores.mean(), 1)
        active_zones = user_data["Area"].nunique()
    else:
        avg_score = 100.0
        active_zones = 0
        
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Records Loaded", f"{total_records:,}")
    col2.metric("Overall Compliance Rating", f"{avg_score}%")
    col3.metric("Monitored Zones Active", str(active_zones))
    
    st.markdown("---")
    
    # 📈 MATRIX STYLED VISUAL CHART SECTION
    st.subheader("📈 Type Breakdown Compliance Charts")
    if not user_data.empty:
        type_summary = user_data.copy()
        type_summary["NumericScore"] = type_summary["Score"].apply(parse_score)
        chart_df = type_summary.groupby("Type")["NumericScore"].mean().reset_index()
        chart_df.columns = ["Audit Type", "Compliance Rating Average (%)"]
        
        st.bar_chart(data=chart_df, x="Audit Type", y="Compliance Rating Average (%)", use_container_width=True)
    
    # 📋 THE MAIN CLEAN DATA LEDGER TABLE View
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📋 Unified Performance Ledger Table")
    
    # Add a global filter mechanism to switch views quickly
    filter_type = st.selectbox("Filter Chart View by Audit Type:", ["All Data Categories"] + list(user_data["Type"].unique()))
    
    display_df = user_data.copy()
    if filter_type != "All Data Categories":
        display_df = display_df[display_df["Type"] == filter_type]
        
    if not display_df.empty:
        display_df.index.name = "Record Index"
        st.dataframe(display_df.reset_index(), use_container_width=True, height=400)
        
        # Row item interactive scrubbing tool
        st.markdown('<div class="danger-zone">', unsafe_allow_html=True)
        st.subheader("🗑️ Record Management Panel")
        row_to_delete = st.selectbox(
            "Select Record Index Row to Scrape/Delete", 
            options=list(display_df.index),
            format_func=lambda idx: f"Index {idx} — {display_df.loc[idx, 'Auditor']} | {display_df.loc[idx, 'Area']} [{display_df.loc[idx, 'Type']}]"
        )
        if st.button("Delete Selected Audit Entry", type="primary"):
            # Check if it's a live row or legacy row
            if user_data.loc[row_to_delete, "Date"] == "Schedule Tracking Period":
                st.error("Cannot delete a static spreadsheet record row directly. You can clear them from Option A below.")
            else:
                live_records = pd.read_csv("audit_data.csv")
                # match using content properties safely
                tgt_auditor = user_data.loc[row_to_delete, "Auditor"]
                tgt_notes = user_data.loc[row_to_delete, "Notes"]
                live_records = live_records[~((live_records["Auditor"] == tgt_auditor) & (live_records["Notes"] == tgt_notes))]
                live_records.to_csv("audit_data.csv", index=False)
                st.success("Record scraped.")
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("No matching records found.")

elif page == "📋 Enter Audit":
    st.header("Enter New Audit Sheet Records")
    with st.form("audit_form", clear_on_submit=True):
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            auditor = st.selectbox("Select Auditor Name", parsed_names)
            area = st.selectbox("Plant Operational Area", ["Maintenance", "Carbon", "Cast House", "Potline", "Environmental"])
            audit_date = st.date_input("Audit Execution Date", value=date.today())
        with f_col2:
            audit_type = st.selectbox("Audit Type Classification", ["LPA", "Safe Observation", "PPE", "LOTO", "Mobile Equipment", "HK Score"])
            is_complete_only = st.checkbox("Mark Audit as Complete ('C' status flag)")
            score = st.number_input("Recorded Performance Score (%)", min_value=0.0, max_value=100.0, value=100.0, step=1.0, disabled=is_complete_only)
            
        notes = st.text_area("Observations & Notes")
        if st.form_submit_button("Submit Entry to Database", type="primary"):
            final_score = "C" if is_complete_only else score
            new_row = pd.DataFrame(
