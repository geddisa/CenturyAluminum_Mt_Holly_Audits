import streamlit as st
import pandas as pd
import os
import hashlib

# -----------------------------
# PAGE SETUP
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum | EHSQ Management")

# -----------------------------
# AUTHENTICATION & SESSION STATE
# -----------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "auth_user_email" not in st.session_state:
    st.session_state.auth_user_email = ""
if "auth_user_name" not in st.session_state:
    st.session_state.auth_user_name = ""

DB_FILE = "users_db.csv"

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def get_users_df():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE)
    return pd.DataFrame(columns=["email", "name", "password"])

# -----------------------------
# AUTH GATEWAY
# -----------------------------
if not st.session_state.authenticated:
    st.title("🛡️ Century Aluminum Identity Access")
    st.markdown("### Mt. Holly EHSQ Auditing System")
    
    tab1, tab2 = st.tabs(["Login", "Sign Up"])
    
    with tab1:
        email = st.text_input("Corporate Email", key="login_email")
        password = st.text_input("Security Password", type="password", key="login_pass")
        if st.button("Authenticate"):
            users = get_users_df()
            user = users[users['email'] == email]
            if not user.empty and user.iloc[0]['password'] == hash_password(password):
                st.session_state.authenticated = True
                st.session_state.auth_user_email = email
                st.session_state.auth_user_name = user.iloc[0]['name']
                st.rerun()
            else:
                st.error("Invalid credentials.")
                
    with tab2:
        name = st.text_input("Full Name")
        email = st.text_input("Email", key="reg_email")
        password = st.text_input("Create Password", type="password", key="reg_pass")
        if st.button("Register Profile"):
            users = get_users_df()
            if email in users['email'].values:
                st.error("Email already exists.")
            else:
                new_user = pd.DataFrame([{"email": email, "name": name, "password": hash_password(password)}])
                pd.concat([users, new_user]).to_csv(DB_FILE, index=False)
                st.success("Registration successful. Please login.")
    st.stop()

# -----------------------------
# DASHBOARD INTERFACE
# -----------------------------
st.sidebar.title("System Navigation")
st.sidebar.info(f"User: **{st.session_state.auth_user_name}**")
if st.sidebar.button("Logout"):
    for key in st.session_state.keys():
        del st.session_state[key]
    st.rerun()

st.title("📊 EHSQ Performance & Audit Ledger")
st.markdown("---")

@st.cache_data
def load_data():
    file_path = "Audit Schedule - Internal - LPA.xlsx"
    if os.path.exists(file_path):
        return pd.read_excel(file_path)
    return pd.DataFrame()

df = load_data()

# Professional Metrics
if not df.empty:
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Active Audits", len(df))
    col2.metric("System Status", "Operational")
    col3.metric("Last Updated", "June 2026")
    
    st.subheader("Master Compliance Ledger")
    st.dataframe(df, use_container_width=True)
else:
    st.warning("Ledger data is currently unavailable. Please verify 'Audit Schedule - Internal - LPA.xlsx' is present.")
