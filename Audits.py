import streamlit as st
import pandas as pd
from datetime import datetime
import os
import hashlib

# -----------------------------
# PAGE SETUP & MODERN PLATFORM DESIGN
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum - EHSQ Auditing Platform")

# Injecting modern web-app component styles
st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 32px; font-weight: 700; color: #0F172A; }
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; }
    h1, h2, h3 { color: #0F172A; font-family: 'Inter', system-ui, sans-serif; font-weight: 600; }
    .login-box { padding: 2.5rem; border-radius: 12px; border: 1px solid #E2E8F0; background-color: #FFFFFF; max-width: 480px; margin: 4rem auto; box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.05); }
    .audit-badge { background-color: #EFF6FF; color: #1E40AF; padding: 0.35rem 0.75rem; border-radius: 9999px; font-weight: 600; font-size: 0.85rem; }
    </style>
""", unsafe_allow_html=True)

# -----------------------------
# 🔒 SECURE USER DATABASE ENGINE
# -----------------------------
DB_FILE = "users_db.csv"
SUBMISSION_FILE = "submitted_audits.csv"

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def load_users():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE).set_index("email").to_dict(orient="index")
    else:
        # Default Admin for first-time setup
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
    if email in users: return False
    new_row = pd.DataFrame([{"email": email, "name": name, "password_hash": hash_password(password)}])
    new_row.to_csv(DB_FILE, mode='a', header=not os.path.exists(DB_FILE), index=False)
    return True

# -----------------------------
# 🛡️ AUTHENTICATION STATE
# -----------------------------
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "auth_user_email" not in st.session_state: st.session_state.auth_user_email = ""
if "auth_user_name" not in st.session_state: st.session_state.auth_user_name = ""

# [Authentication logic remains as per your provided snippet...]
# (For brevity, ensure the Auth Gateway logic is inserted here)

if not st.session_state.authenticated:
    st.title("🛡️ Identity Access Gateway")
    # ... [Insert your Auth Gateway code here] ...
    st.stop()

# -----------------------------
# 📊 PERSISTENCE & DASHBOARD
# -----------------------------

# Initialize Submission Storage
if os.path.exists(SUBMISSION_FILE):
    submitted_df = pd.read_csv(SUBMISSION_FILE)
else:
    submitted_df = pd.DataFrame(columns=["Timestamp", "Auditor", "Area", "Status", "Notes"])

# Main App Navigation
st.sidebar.markdown(f"**Logged In:** {st.session_state.auth_user_name}")
if st.sidebar.button("Logout"):
    for key in st.session_state.keys(): del st.session_state[key]
    st.rerun()

st.title("📊 Mt. Holly EHSQ Audit Hub")
tab1, tab2, tab3 = st.tabs(["Compliance Overview", "Submit New Audit", "Submission History"])

with tab1:
    st.subheader("Performance Metrics")
    col1, col2, col3 = st.columns(3)
    col1.metric("Active Audits", len(submitted_df))
    col2.metric("Compliant Logs", len(submitted_df[submitted_df['Status'] == 'Compliant']))
    col3.metric("Pending Review", "0")
    
    st.markdown("### Master Schedule Ledger")
    # Using your load_and_transform_schedule function
    # excel_records, _ = load_and_transform_schedule("Audit Schedule - Internal - LPA.xlsx")
    # st.dataframe(excel_records, use_container_width=True)

with tab2:
    st.subheader("Report Audit Findings")
    with st.form("audit_submission"):
        col1, col2 = st.columns(2)
        auditor = col1.text_input("Auditor Name", value=st.session_state.auth_user_name)
        area = col2.selectbox("Area", ["Carbon", "Cast House", "Potline", "Maintenance", "Environmental"])
        status = st.selectbox("Compliance", ["Compliant", "Non-Compliant"])
        notes = st.text_area("Observations")
        
        if st.form_submit_button("Submit Record"):
            new_entry = pd.DataFrame([{
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Auditor": auditor, "Area": area, "Status": status, "Notes": notes
            }])
            new_entry.to_csv(SUBMISSION_FILE, mode='a', header=not os.path.exists(SUBMISSION_FILE), index=False)
            st.success("Audit submitted to database.")

with tab3:
    st.subheader("Historical Submissions")
    if os.path.exists(SUBMISSION_FILE):
        logs = pd.read_csv(SUBMISSION_FILE)
        st.dataframe(logs, use_container_width=True)
        st.download_button("Export History (CSV)", logs.to_csv(index=False), "audit_history.csv", "text/csv")
