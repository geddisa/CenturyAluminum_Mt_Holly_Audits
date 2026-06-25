import streamlit as st
import pandas as pd
from datetime import date
import os
import hashlib

# -----------------------------
# PAGE SETUP
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum - EHSQ Auditing Platform")

# -----------------------------
# AUTHENTICATION STATE
# -----------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "auth_user_email" not in st.session_state:
    st.session_state.auth_user_email = ""
if "auth_user_name" not in st.session_state:
    st.session_state.auth_user_name = ""

# -----------------------------
# USER DB ENGINE
# -----------------------------
DB_FILE = "users_db.csv"

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def load_users():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE).set_index("email").to_dict(orient="index")
    return {"admin@centuryaluminum.com": {"name": "System Administrator", "password_hash": hash_password("Century2026!")}}

# -----------------------------
# AUTH WALL
# -----------------------------
if not st.session_state.authenticated:
    st.title("🛡️ Identity Access Gateway")
    email_input = st.text_input("Corporate Email")
    password_input = st.text_input("Password", type="password")
    
    if st.button("Login"):
        users = load_users()
        if email_input in users and hash_password(password_input) == users[email_input]["password_hash"]:
            st.session_state.authenticated = True
            st.session_state.auth_user_email = email_input
            st.session_state.auth_user_name = users[email_input]["name"]
            st.rerun()
        else:
            st.error("Invalid credentials.")
    st.stop()

# -----------------------------
# MAIN APP
# -----------------------------
st.sidebar.markdown(f"**Logged In:** {st.session_state.auth_user_name}")
if st.sidebar.button("Logout"):
    for key in st.session_state.keys():
        del st.session_state[key]
    st.rerun()

st.title("Century Aluminum Corporate Audit Hub")

# Data loading engine
@st.cache_data
def load_data():
    file_path = "Audit Schedule - Internal - LPA.xlsx"
    if os.path.exists(file_path):
        return pd.read_excel(file_path)
    return pd.DataFrame()

df = load_data()
if not df.empty:
    st.dataframe(df, use_container_width=True)
else:
    st.info("No data file found. Ensure 'Audit Schedule - Internal - LPA.xlsx' is uploaded.")
