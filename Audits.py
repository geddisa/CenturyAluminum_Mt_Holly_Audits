import streamlit as st
import pandas as pd
import os
import hashlib

# -----------------------------
# PAGE CONFIGURATION
# -----------------------------
st.set_page_config(layout="wide", page_title="Century Aluminum - EHSQ Auditing Platform")

# -----------------------------
# AUTHENTICATION & SESSION STATE
# -----------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_name" not in st.session_state:
    st.session_state.user_name = None

DB_FILE = "users_db.csv"

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def get_users_df():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE)
    return pd.DataFrame(columns=["email", "name", "password"])

# -----------------------------
# LOGIN / SIGNUP INTERFACE
# -----------------------------
if not st.session_state.authenticated:
    st.title("🛡️ Identity Access Gateway")
    
    choice = st.radio("Select Action:", ["Login", "Sign Up"])
    
    if choice == "Login":
        email = st.text_input("Email")
        password = st.text_input("Password", type="password")
        if st.button("Login"):
            users = get_users_df()
            user = users[users['email'] == email]
            if not user.empty and user.iloc[0]['password'] == hash_password(password):
                st.session_state.authenticated = True
                st.session_state.user_name = user.iloc[0]['name']
                st.rerun()
            else:
                st.error("Invalid email or password.")
                
    else:  # Sign Up
        name = st.text_input("Full Name")
        email = st.text_input("Email")
        password = st.text_input("Create Password", type="password")
        if st.button("Register"):
            users = get_users_df()
            if email in users['email'].values:
                st.error("Email already exists.")
            else:
                new_user = pd.DataFrame([{"email": email, "name": name, "password": hash_password(password)}])
                pd.concat([users, new_user]).to_csv(DB_FILE, index=False)
                st.success("Account created! Please login.")
    st.stop()

# -----------------------------
# SECURED APP CONTENT
# -----------------------------
st.sidebar.write(f"**Welcome, {st.session_state.user_name}**")
if st.sidebar.button("Logout"):
    st.session_state.authenticated = False
    st.rerun()

st.title("Century Aluminum Corporate Audit Hub")
st.write("You have successfully accessed the secure audit dashboard.")

# Your data loading logic here
