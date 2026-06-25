import streamlit as st
import pandas as pd
import os
import hashlib
from datetime import datetime

# -----------------------------
# PAGE CONFIGURATION
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum | EHSQ Management")

# -----------------------------
# AUTHENTICATION
# -----------------------------
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "user_name" not in st.session_state: st.session_state.user_name = None

def hash_password(password): return hashlib.sha256(password.encode()).hexdigest()
def get_users(): 
    if os.path.exists("users_db.csv"): return pd.read_csv("users_db.csv")
    return pd.DataFrame(columns=["email", "name", "password"])

# -----------------------------
# AUTH GATEWAY
# -----------------------------
if not st.session_state.authenticated:
    st.title("🛡️ Century Aluminum Identity Access")
    col1, col2 = st.columns([1, 1])
    with col1:
        email = st.text_input("Corporate Email")
        password = st.text_input("Password", type="password")
        if st.button("Access Portal"):
            users = get_users()
            user = users[users['email'] == email]
            if not user.empty and user.iloc[0]['password'] == hash_password(password):
                st.session_state.authenticated = True
                st.session_state.user_name = user.iloc[0]['name']
                st.rerun()
    st.stop()

# -----------------------------
# SUBMISSION DATABASE ENGINE
# -----------------------------
SUBMISSION_FILE = "submitted_audits.csv"

def save_audit(data_dict):
    df_new = pd.DataFrame([data_dict])
    if os.path.exists(SUBMISSION_FILE):
        df_old = pd.read_csv(SUBMISSION_FILE)
        df_final = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_final = df_new
    df_final.to_csv(SUBMISSION_FILE, index=False)

# -----------------------------
# DASHBOARD
# -----------------------------
st.sidebar.title("Mt. Holly Navigation")
if st.sidebar.button("Logout"):
    st.session_state.authenticated = False
    st.rerun()

st.title("📊 Mt. Holly EHSQ Audit Hub")
tab1, tab2, tab3 = st.tabs(["Compliance Dashboard", "Submit Audit", "Submission History"])

# TAB 1: DASHBOARD
with tab1:
    st.subheader("Master Compliance Ledger")
    # Load your main schedule file here
    if os.path.exists("Audit Schedule - Internal - LPA.xlsx"):
        df_main = pd.read_excel("Audit Schedule - Internal - LPA.xlsx")
        st.dataframe(df_main, use_container_width=True)

# TAB 2: SUBMIT AUDIT
with tab2:
    st.subheader("New Audit Submission")
    with st.form("audit_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        area = col1.selectbox("Department/Area", ["Carbon", "Cast House", "Potline", "Maintenance", "Environmental"])
        auditor = col2.text_input("Auditor Name", value=st.session_state.user_name)
        
        status = st.selectbox("Compliance Status", ["Compliant", "Non-Compliant", "Risk Corrected"])
        notes = st.text_area("Observations / Notes")
        
        submitted = st.form_submit_button("Submit Audit Record")
        
        if submitted:
            data = {
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Auditor": auditor,
                "Area": area,
                "Status": status,
                "Notes": notes
            }
            save_audit(data)
            st.success("Audit submitted successfully! It has been added to the local log.")

# TAB 3: HISTORY
with tab3:
    st.subheader("Audit Submission Log")
    if os.path.exists(SUBMISSION_FILE):
        df_logs = pd.read_csv(SUBMISSION_FILE)
        st.dataframe(df_logs, use_container_width=True)
        
        # Download button to get the data out
        with open(SUBMISSION_FILE, "rb") as file:
            st.download_button("Download All Submissions (CSV)", file, "audit_history.csv")
    else:
        st.info("No audits submitted yet.")
