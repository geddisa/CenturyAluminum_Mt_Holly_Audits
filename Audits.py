import streamlit as st
import pandas as pd
import os
import hashlib

# -----------------------------
# PAGE CONFIGURATION
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum | Mt. Holly EHSQ Portal")

# Custom Styling
st.markdown("""
    <style>
    .stMetric { background-color: #f0f2f6; padding: 15px; border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

# -----------------------------
# AUTH & SESSION STATE
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
            else: st.error("Invalid credentials.")
    st.stop()

# -----------------------------
# DASHBOARD LOGIC
# -----------------------------
st.sidebar.title("Mt. Holly Navigation")
st.sidebar.info(f"User: **{st.session_state.user_name}**")
if st.sidebar.button("Logout"):
    st.session_state.authenticated = False
    st.rerun()

st.title("📊 Mt. Holly EHSQ Audit Hub")

# Data Loader
@st.cache_data
def load_audit_data(sheet_name):
    file_path = "Audit Schedule - Internal - LPA.xlsx"
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

# Create Tabs
tab1, tab2, tab3, tab4 = st.tabs(["Overview", "LPA Compliance", "Safety Observations", "Specialized Audits"])

with tab1:
    st.subheader("Performance KPIs")
    c1, c2, c3 = st.columns(3)
    c1.metric("Active Audits", "42") # Dynamic counts can be added here
    c2.metric("Compliance Score", "94%")
    c3.metric("System Status", "Operational")
    st.write("Welcome to the centralized audit tracking system for the Mt. Holly facility.")

with tab2:
    st.subheader("LPA & HK Schedule")
    hk_data = load_audit_data("HK")
    if not hk_data.empty: st.dataframe(hk_data, use_container_width=True)
    else: st.warning("HK Data missing.")

with tab3:
    st.subheader("Safe Observations (GS/EHS & Leadership)")
    obs_data = load_audit_data("Safe Obs GS EHS")
    if not obs_data.empty: st.dataframe(obs_data, use_container_width=True)

with tab4:
    st.subheader("LOTO, PPE & Mobile Equipment")
    loto = load_audit_data("LOTO")
    if not loto.empty: 
        st.write("### LOTO Log")
        st.dataframe(loto, use_container_width=True)
