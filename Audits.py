import streamlit as st
import pandas as pd
from datetime import date
import os
import hashlib
from streamlit_javascript import st_javascript

# -----------------------------
# PAGE SETUP & MODERN PLATFORM DESIGN
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum - EHSQ Auditing Platform")

st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 32px; font-weight: 700; color: #0F172A; }
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; }
    h1, h2, h3 { color: #0F172A; font-family: 'Inter', system-ui, sans-serif; font-weight: 600; }
    .login-box { padding: 2.5rem; border-radius: 12px; border: 1px solid #E2E8F0; background-color: #FFFFFF; max-width: 480px; margin: 4rem auto; box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.05); }
    .metric-card { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 1.25rem; border-radius: 10px; text-align: center; }
    </style>
""", unsafe_allow_html=True)

# -----------------------------
# 🌐 BROWSER STORAGE PERSISTENCE ENGINE
# -----------------------------
# Read persistent token values from the browser session context on load
js_auth = st_javascript("sessionStorage.getItem('authenticated');")
js_email = st_javascript("sessionStorage.getItem('auth_user_email');")
js_name = st_javascript("sessionStorage.getItem('auth_user_name');")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = True if js_auth == "true" else False
if "auth_user_email" not in st.session_state:
    st.session_state.auth_user_email = js_email if js_email else ""
if "auth_user_name" not in st.session_state:
    st.session_state.auth_user_name = js_name if js_name else ""

# Sync session state if browser storage values just returned from async JS engine
if js_auth == "true" and not st.session_state.authenticated:
    st.session_state.authenticated = True
    st.session_state.auth_user_email = js_email
    st.session_state.auth_user_name = js_name
    st.rerun()

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
                
                # Write state explicitly into browser memory to prevent refresh wipeout
                st_javascript(f"sessionStorage.setItem('authenticated', 'true');")
                st_javascript(f"sessionStorage.setItem('auth_user_email', '{clean_email}');")
                st_javascript(f"sessionStorage.setItem('auth_user_name', '{users_database[clean_email]['name']}');")
                
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
# 📊 AUTOMATED PARSING ENGINE
# -----------------------------
primary_file = "Audit Schedule - Internal - LPA_2.xlsx"
fallback_file = "Audit Schedule - Internal - LPA.xlsx"
excel_file = primary_file if os.path.exists(primary_file) else fallback_file

@st.cache
