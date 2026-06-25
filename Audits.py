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

# Clean, professional CSS stylesheet injecting custom corporate spacing rules
st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 26px; font-weight: bold; color: #1E3A8A; }
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 6px; }
    h1, h2, h3 { color: #0F172A; font-family: 'Segoe UI', Helvetica, Arial, sans-serif; }
    .login-box { padding: 2.5rem; border-radius: 8px; border: 1px solid #CBD5E1; background-color: #F8FAFC; max-width: 480px; margin: 2rem auto; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); }
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
    
    new_row = pd.DataFrame([{
        "email": email,
        "name": name,
        "password_hash": hash_password(password)
    }])
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
                    st.success("✅ Credentials verified. Access granted.")
                    st.rerun()
                else:
                    st.error("❌ Invalid email or password entry.")
            st.markdown('</div>', unsafe_allow_html=True)

    else:
        with st.container():
            st.markdown('<div class="login-box">', unsafe_allow_html=True)
            st.subheader("Create Authorized Profile")
            st.info("Registration requires an official corporate Century Aluminum email account.")
            
            reg_name = st.text_input("Full Name (First Last)", placeholder="John Doe")
            reg_email = st.text_input("Corporate Email Address", placeholder="username@centuryaluminum.com")
            reg_password = st.text_input("Choose Password", type="password")
            confirm_password = st.text_input("Confirm Password", type="password")
            
            if st.button("Register Corporate Profile", use_container_width=True):
                clean_reg_email = reg_email.strip().lower()
                
                if not reg_name.strip():
                    st.error("❌ Please provide your name.")
                elif not clean_reg_email.endswith("@centuryaluminum.com"):
                    st.error("❌ Registration Blocked. Email domain must explicitly be @centuryaluminum.com.")
                elif len(reg_password) < 6:
                    st.error("❌ Password must be at least 6 characters long.")
                elif reg_password != confirm_password:
                    st.error("❌ Passwords do not match.")
                elif clean_reg_email in users_database:
                    st.error("❌ This email address is already registered.")
                else:
                    success = save_new_user(clean_reg_email, reg_name.strip(), reg_password)
                    if success:
                        st.success("🎉 Registration complete! Switch to the 'Sign In' tab above to authenticate.")
                    else:
                        st.error("❌ Database conflict occurred.")
            st.markdown('</div>', unsafe_allow_html=True)

    st.stop()

# -----------------------------
# 👥 MASTER AUDITOR REGISTRY CLEANER
# -----------------------------
excel_file = "Audit Schedule - Internal - LPA.xlsx"

if not os.path.exists(excel_file):
    st.error("❌ System Critical Error: Source Excel file not found.")
    st.stop()

@st.cache_data
def get_clean_auditor_names(file_path):
    xls = pd.ExcelFile(file_path)
    discovered_names = set()
    
    EXCLUDE_KEYWORDS = {
        "week", "sheet", "audit", "shift", "room", "bake", "casting", "fabrication", 
        "compressor", "hot side", "cold side", "east", "west", "baghouse", "tank", 
        "mill", "shop", "rod", "house", "fluid", "building", "café", "snacks", 
        "catering", "equipment", "score", "formula", "required", "vacation", "complete", "loto"
    }

    for sheet in xls.sheet_names:
        if sheet in ["Jobs and shifts", "Sheet1"]:
            continue
        try:
            df_temp = pd.read_excel(file_path, sheet_name=sheet)
            for row in df_temp.itertuples(index=False):
                for cell in row:
                    if pd.isna(cell):
                        continue
                    cell_str = str(cell).strip().replace('"', '')
                    
                    if not cell_str or any(k in cell_str.lower() for k in EXCLUDE_KEYWORDS):
                        continue
                    
                    if ',' in cell_str:
                        parts = [p.strip() for p in cell_str.split(',')]
                        if len(parts) == 2:
                            first_name = re.sub(r'\(.*?\)', '', parts[1]).strip()
                            last_name = parts[0].strip()
                            cell_str = f"{first_name} {last_name}"
                    
                    words = cell_str.split()
                    if len(words) >= 2 and words[0][0].isupper() and words[1][0].isupper() and len(cell_str) < 30:
                        discovered_names.add(cell_str)
        except:
            pass
            
    return sorted(list(discovered_names))

name_list = get_clean_auditor_names(excel_file)
if not name_list:
    name_list = ["Anthony Wall", "Art DiFilippo", "Brett Meyer", "Brian Weatherford", "Bryan Profit", "Freddie Gamble", "Tim Kass"]

# -----------------------------
# RUN TIME PERSISTENCE ENGINE
# -----------------------------
if os.path.exists("audit_data.csv"):
    user_data = pd.read_csv("audit_data.csv")
else:
    user_data = pd.DataFrame(columns=["Date", "Auditor", "Area", "Type", "Score", "Notes"])

def save_user_data(df):
    df.to_csv("audit_data.csv", index=False)

# -----------------------------
# SIDEBAR CONTROL PANEL
# -----------------------------
st.sidebar.markdown(f"👤 **Session Profile:**\n**{st.session_state.auth_user_name}**\n`{st.session_state.auth_user_email}`")
if st.sidebar.button("Secure Logout", use_container_width=True):
    st.session_state.authenticated = False
    st.session_state.auth_user_email = ""
    st.session_state.auth_user_name = ""
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🏢 Facility Management")
page = st.sidebar.radio(
    "Navigation Options",
    ["📊 Dashboard", "📋 Enter Audit", "📁 Excel Viewer", "👥 Names"],
    key="main_navigation_radio"
)

# -----------------------------
# DYNAMIC NAVIGATION PAGES
# -----------------------------
if page == "📊 Dashboard":
    st.header("Century Aluminum Audit Performance Dashboard")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Records Loaded", f"{len(user_data):,}")
    
    if not user_data.empty:
        try:
            numeric_scores = pd.to_numeric(user_data['Score'], errors='coerce')
            avg_score = round(numeric_scores.mean(), 1) if not numeric_scores.dropna().empty else 100.0
        except:
            avg_score = 100.0
        active_zones = user_data["Area"].nunique()
    else:
        avg_score = 100.0
        active_zones = 0
        
    col2.metric("Overall System Rating", f"{avg_score}%")
    col3.metric("Monitored Zones Active", str(active_zones))
    
    st.markdown("---")
    st.subheader("📋 Historical Consolidated Data Ledger")
    
    if not user_data.empty:
        display_df = user_data.copy()
        display_df.index.name = "Row ID"
        st.dataframe(display_df.reset_index(), use_container_width=True)
        
        st.markdown('<div class="danger-zone">', unsafe_allow_html=True)
        st.subheader("🗑️ Record Management Panel")
        st.markdown("Select a targeted layout entry row index below to permanently scrub it from the storage layer.")
        
        del_col1, del_col2 = st.columns([2, 1])
        with del_col1:
            row_to_delete = st.selectbox(
                "Select Row ID to Delete", 
                options=list(user_data.index),
                format_func=lambda idx: f"Row {idx} — {user_data.loc[idx, 'Auditor']} | {user_data.loc[idx, 'Area']} ({user_data.loc[idx, 'Type']})"
            )
        with del_col2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Delete Selected Audit Entry", type="primary", use_container_width=True):
                user_data = user_data.drop(row_to_delete).reset_index(drop=True)
                save_user_data(user_data)
                st.success(f"Record successfully deleted.")
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("The internal database ledger is currently empty. Submit a log sheet entry in 'Enter Audit' to populate records.")

elif page == "📋 Enter Audit":
    st.header("Enter New Audit Sheet Records")

    with st.form("audit_form", clear_on_submit=True):
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            auditor = st.selectbox("Select Auditor Name", name_list)
            area = st.selectbox("Plant Operational Area", ["Maintenance", "Carbon", "Cast House", "Potline", "Environmental"])
            audit_date = st.date_input("Audit Execution Date", value=date.today())
        with f_col2:
            audit_type = st.selectbox("Audit Type Classification", ["LPA", "Safe Observation", "PPE", "LOTO", "Mobile Equipment", "HK Score"])
            
            is_complete_only = st.checkbox(
                "Mark Audit as Complete (Score not required / observations only)", 
                help="Check this if the inspection type relies on completion status/checkmarks rather than a raw percentage score."
            )
            
            score = st.number_input(
                "Recorded Performance Score (%)", 
                min_value=0.0, max_value=100.0, 
                value=0.0 if is_complete_only else 100.0, 
                step=1.0,
                disabled=is_complete_only,
                help="Unavailable when 'Mark Audit as Complete' is active." if is_complete_only else None
            )
            
        notes = st.text_area("Observations & Notes", placeholder="Type details or observations here...")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.form_submit_button("Submit Entry to Database", type="primary"):
            final_score = "N/A" if is_complete_only else score
            completion_tag = "[COMPLETE]" if is_complete_only else f"[{score}%]"
            
            new_row = pd.DataFrame([{
                "Date": str(audit_date),
                "Auditor": auditor,
                "Area": area,
                "Type": audit_type,
                "Score": final_score,
                "Notes": f"{completion_tag} {notes} (Logged securely by profile: {st.session_state.auth_user_email})"
            }])
            user_data = pd.concat([user_data, new_row], ignore_index=True)
            save_user_data(user_data)
            st.success("✅ Audit logged securely into ledger file database!")

elif page == "📁 Excel Viewer":
    st.header("Spreadsheet Tab Visualizer")
    xls_viewer = pd.ExcelFile(excel_file)
    sheet = st.selectbox("Choose Sheet Tab to View", xls_viewer.sheet_names)
    st.dataframe(pd.read_excel(excel_file, sheet_name=sheet), use_container_width=True)

elif page == "👥 Names":
    st.header("Verified Clean Auditor Registry")
    st.write(f"Total Unique Filtered Personnel Names Extracted: **{len(name_list)}**")
    name_df = pd.DataFrame(name_list, columns=["Employee Name Listing"])
    st.dataframe(name_df, use_container_width=True)

# -----------------------------
# ⚠️ UNIFIED SYSTEM RESET AREA
# -----------------------------
if page == "📊 Dashboard":
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    with st.expander("🚨 Advanced Administrative System Wiping"):
        st.warning("Warning: Clicking the button below wipes all manual records entered up to this second.")
        if st.button("Reset Entire CSV Local Database Layer"):
            if os.path.exists("audit_data.csv"):
                os.remove("audit_data.csv")
            st.cache_data.clear()
            st.success("Database wiped successfully. Reloading system...")
            st.rerun()
