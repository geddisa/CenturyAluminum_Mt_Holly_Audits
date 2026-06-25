import streamlit as st
import pandas as pd
import os
import hashlib

# -----------------------------
# PAGE CONFIGURATION
# -----------------------------
st.set_page_config(
    layout="wide", 
    page_title="Century Aluminum | EHSQ Audit Portal",
    menu_items={'Get Help': None, 'Report a bug': None, 'About': "Century Aluminum EHSQ Platform"}
)

# -----------------------------
# AUTHENTICATION & SECURITY
# -----------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_name" not in st.session_state:
    st.session_state.user_name = None

DB_FILE = "users_db.csv"

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def get_users():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE)
    return pd.DataFrame(columns=["email", "name", "password"])

# -----------------------------
# INTERFACE LAYER
# -----------------------------
if not st.session_state.authenticated:
    st.title("🛡️ Century Aluminum Identity Access")
    st.markdown("### Mt. Holly EHSQ Auditing System")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        choice = st.radio("Access Credentials", ["Login", "Sign Up"])
        email = st.text_input("Corporate Email")
        password = st.text_input("Security Password", type="password")
        
        if choice == "Login":
            if st.button("Authenticate", type="primary"):
                users = get_users()
                user = users[users['email'] == email]
                if not user.empty and user.iloc[0]['password'] == hash_password(password):
                    st.session_state.authenticated = True
                    st.session_state.user_name = user.iloc[0]['name']
                    st.rerun()
                else:
                    st.error("Invalid credentials.")
        else:
            name = st.text_input("Full Name")
            if st.button("Register Profile"):
                users = get_users()
                if email in users['email'].values:
                    st.error("Email already registered.")
                else:
                    new_user = pd.DataFrame([{"email": email, "name": name, "password": hash_password(password)}])
                    pd.concat([users, new_user]).to_csv(DB_FILE, index=False)
                    st.success("Registration successful. Please login.")
    st.stop()

# -----------------------------
# SECURE DASHBOARD VIEW
# -----------------------------
st.sidebar.title("Navigation")
st.sidebar.info(f"User: **{st.session_state.user_name}**")
if st.sidebar.button("Logout"):
    st.session_state.authenticated = False
    st.rerun()

st.title("📊 EHSQ Performance Dashboard")
st.markdown("---")

# Data loading with caching for performance
@st.cache_data
def load_schedule():
    file = "Audit Schedule - Internal - LPA.xlsx"
    if os.path.exists(file):
        # Loading from primary sheet for demonstration
        return pd.read_excel(file, sheet_name=0) 
    return None

data = load_schedule()

if data is not None:
    st.dataframe(data, use_container_width=True)
else:
    st.warning("Audit ledger file not detected. Please verify 'Audit Schedule - Internal - LPA.xlsx' is present.")
