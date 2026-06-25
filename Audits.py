import streamlit as st
import pandas as pd
from datetime import date
import os
import hashlib

# -----------------------------
# PAGE SETUP & MODERN PLATFORM DESIGN
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum - EHSQ Auditing Platform")

# Injecting modern web-app component styles (Overrides spreadsheet look with card components)
st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 32px; font-weight: 700; color: #0F172A; }
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; }
    h1, h2, h3 { color: #0F172A; font-family: 'Inter', system-ui, sans-serif; font-weight: 600; }
    .login-box { padding: 2.5rem; border-radius: 12px; border: 1px solid #E2E8F0; background-color: #FFFFFF; max-width: 480px; margin: 4rem auto; box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.05); }
    .metric-card { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 1.25rem; border-radius: 10px; text-align: center; }
    .audit-badge { background-color: #EFF6FF; color: #1E40AF; padding: 0.35rem 0.75rem; border-radius: 9999px; font-weight: 600; font-size: 0.85rem; }
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
# 📊 AUTOMATED PARSING ENGINE (Schedules to Chart Objects)
# -----------------------------
excel_file = "Audit Schedule - Internal - LPA.xlsx"

@st.cache_data
def load_and_transform_schedule(file_path):
    if not os.path.exists(file_path):
        return pd.DataFrame(), []
    
    xls = pd.ExcelFile(file_path)
    all_records = []
    auditors_found = set()
    
    # Specific layout tab-parsers mapping spreadsheet formats into web metrics
    for sheet in xls.sheet_names:
        if sheet in ["Jobs and shifts", "Sheet1"]:
            continue
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            for _, row in df.iterrows():
                if row.dropna().empty: continue
                
                raw_name = str(row.iloc[0]).strip()
                if ',' in raw_name:
                    p = raw_name.split(',')
                    raw_name = f"{p[1].strip()} {p[0].strip()}"
                
                if any(k in raw_name.lower() for k in ["week", "sheet", "audit", "shift", "score", "nan", "ppe", "loto"]):
                    continue
                
                if len(raw_name) > 2 and not raw_name.isdigit():
                    auditors_found.add(raw_name)
                    
                    # Transform standard structural tracking indicators safely
                    for val in row.iloc[2:]:
                        v_str = str(val).strip().upper()
                        if v_str in ['C', 'X', '100', '88.89', '90', '95', '75', '86', '96', '64'] or any(char.isdigit() for char in v_str):
                            score_out = "100%" if v_str in ['C', 'X', 'COMPLETE'] else (f"{v_str}%" if "%" not in v_str else v_str)
                            area_out = str(row.iloc[1]) if len(str(row.iloc[1])) > 3 else "General Facility Floor"
                            
                            all_records.append({
                                "Target Period": "Spreadsheet Baseline Schedule",
                                "Auditor": raw_name,
                                "Operational Area": area_out,
                                "Classification Type": sheet,
                                "Current Status / Score": score_out,
                            })
        except:
            pass
    return pd.DataFrame(all_records), sorted(list(auditors_found))

excel_records, parsed_names = load_and_transform_schedule(excel_file)

if not parsed_names:
    parsed_names = ["Anthony Wall", "Art DiFilippo", "Brett Meyer", "Brian Weatherford", "Bryan Profit", "Freddie Gamble", "Tim Kass"]

# Load up persistent web inputs 
if os.path.exists("
