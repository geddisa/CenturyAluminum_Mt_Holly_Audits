import streamlit as st
import pandas as pd
import os
import hashlib

# -----------------------------
# PAGE CONFIGURATION
# -----------------------------
st.set_page_config(
    layout="wide", 
    page_title="Century Aluminum | EHSQ Management",
    menu_items={'Get Help': None, 'Report a bug': None}
)

# Professional CSS
st.markdown("""
    <style>
    .main { background-color: #F8FAFC; }
    .stMetric { background-color: #FFFFFF; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

# -----------------------------
# AUTHENTICATION ENGINE
# -----------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def get_users_db():
    if os.path.exists("users_db.csv"):
        return pd.read_csv("users_db.csv")
    return pd.DataFrame(columns=["email", "name", "password"])

# -----------------------------
# AUTH GATEWAY
# -----------------------------
if not st.session_state.authenticated:
    st.title("🛡️ Century Aluminum EHSQ Portal")
    col1, col2 = st.columns([1, 1])
    with col1:
        email = st.text_input("Corporate Email")
        password = st.text_input("Password", type="password")
        if st.button("Access Dashboard"):
            users = get_users_db()
            user = users[users['email'] == email]
            if not user.empty and user.iloc[0]['password'] == hash_password(password):
                st.session_state.authenticated = True
                st.session_state.user_name = user.iloc[0]['name']
                st.rerun()
            else:
                st.error("Invalid credentials.")
    st.stop()

# -----------------------------
# DASHBOARD INTERFACE
# -----------------------------
st.sidebar.title("Navigation")
st.sidebar.write(f"Logged in as: **{st.session_state.user_name}**")
if st.sidebar.button("Logout"):
    st.session_state.authenticated = False
    st.rerun()

st.title("📊 EHSQ Performance & Audit Ledger")

# Data loading engine
@st.cache_data
def load_data():
    file = "Audit Schedule - Internal - LPA.xlsx"
    if os.path.exists(file):
        return pd.read_excel(file)
    return pd.DataFrame()

df = load_data()

# Executive Metrics
if not df.empty:
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Active Audits", len(df))
    m2.metric("Departments Monitored", df['Department'].nunique() if 'Department' in df.columns else "N/A")
    m3.metric("System Status", "Operational")
    
    st.subheader("Master Compliance Ledger")
    st.dataframe(df, use_container_width=True)
else:
    st.warning("Ledger data is currently unavailable. Please confirm file path.")
