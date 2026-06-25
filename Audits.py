import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Audit Dashboard", layout="wide")

st.title("Internal Audit Management System")

# Define file mapping
audit_files = {
    "LOTO": "Audit Schedule - Internal - LPA.xlsx - LOTO.csv",
    "PPE": "Audit Schedule - Internal - LPA.xlsx - PPE.csv",
    "HK": "Audit Schedule - Internal - LPA.xlsx - HK.csv",
    "Safe Obs - GS and EHS": "Audit Schedule - Internal - LPA.xlsx - Safe Obs - GS and EHS.csv",
    "Safe Obs - Leadership": "Audit Schedule - Internal - LPA.xlsx - Safe Obs - Leadership.csv",
    "Mobile Equip": "Audit Schedule - Internal - LPA.xlsx - Mobile Equip.csv"
}

# Sidebar Navigation
selection = st.sidebar.selectbox("Select Audit Type", list(audit_files.keys()))

# Display Logic
if os.path.exists(audit_files[selection]):
    try:
        df = pd.read_csv(audit_files[selection])
        st.header(f"Dashboard: {selection}")
        
        # Simple search/filter
        search = st.text_input(f"Filter {selection} data...", "")
        if search:
            df = df[df.apply(lambda row: row.astype(str).str.contains(search, case=False).any(), axis=1)]
            
        st.dataframe(df, use_container_width=True)
        
        # Download button
        st.download_button("Download Current View", df.to_csv(index=False), f"{selection}_export.csv", "text/csv")
        
    except Exception as e:
        st.error(f"Error loading file: {e}")
else:
    st.error(f"File not found: {audit_files[selection]}")
