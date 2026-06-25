import streamlit as st
import pandas as pd
from datetime import date
import os

# -----------------------------
# PAGE SETUP
# -----------------------------
st.set_page_config(layout="wide")

st.title("Audit Management System")

excel_file = "Audit Schedule - Internal - LPA.xlsx"

if not os.path.exists(excel_file):
    st.error("❌ Excel file not found")
    st.write("Files in folder:", os.listdir())
    st.stop()

xls = pd.ExcelFile(excel_file)

st.success("✅ File loaded")

# -----------------------------
# ✅ GET ALL NAMES FROM ALL SHEETS
# -----------------------------
all_names = set()

for sheet in xls.sheet_names:
    df_temp = pd.read_excel(excel_file, sheet_name=sheet)

    for col in df_temp.columns:
        for val in df_temp[col]:
            if isinstance(val, str):
                val = val.strip()

                # simple filter (keeps names, removes junk)
                if val and len(val.split()) >= 2:
                    all_names.add(val)

# convert to sorted list
name_list = sorted(all_names)

# -----------------------------
# Sidebar navigation
# -----------------------------
page = st.sidebar.radio(
    "Navigation",
    ["📊 Dashboard", "📋 Enter Audit", "📁 Excel Data Viewer", "👥 All Names"]
)

# -----------------------------
# LOAD USER DATA
# -----------------------------
def load_user_data():
    if os.path.exists("audit_data.csv"):
        return pd.read_csv("audit_data.csv")
    else:
        return pd.DataFrame(columns=[
            "Date", "Auditor", "Area", "Type", "Score", "Notes"
        ])

def save_user_data(df):
    df.to_csv("audit_data.csv", index=False)

user_data = load_user_data()

# -----------------------------
# 📋 ENTER AUDIT
# -----------------------------
if page == "📋 Enter Audit":
    st.header("Enter New Audit")

    with st.form("audit_form"):

        # ✅ dynamic names instead of hardcoded
        auditor = st.selectbox("Auditor", name_list)

        area = st.selectbox("Area", [
            "Maintenance",
            "Carbon",
            "Cast House",
            "Potline",
            "Environmental"
        ])

        audit_type = st.selectbox("Audit Type", [
            "LPA",
            "Safe Observation",
            "PPE",
            "LOTO",
            "Mobile Equipment"
        ])

        score = st.number_input("Score (%)", 0.0, 100.0)

        audit_date = st.date_input("Date", value=date.today())

        notes = st.text_area("Notes")

        submitted = st.form_submit_button("Submit")

        if submitted:
            new_row = pd.DataFrame([{
                "Date": audit_date,
                "Auditor": auditor,
                "Area": area,
                "Type": audit_type,
                "Score": score,
                "Notes": notes
            }])

            user_data = pd.concat([user_data, new_row], ignore_index=True)
            save_user_data(user_data)

            st.success("✅ Audit Saved!")

# -----------------------------
# 📊 DASHBOARD
# -----------------------------
elif page == "📊 Dashboard":
    st.header("Audit Dashboard")

    if user_data.empty:
        st.warning("No audits yet.")
    else:
        st.sidebar.subheader("Filters")

        auditor_filter = st.sidebar.selectbox(
            "Auditor",
            ["All"] + list(user_data["Auditor"].unique())
        )

        area_filter = st.sidebar.selectbox(
            "Area",
            ["All"] + list(user_data["Area"].unique())
        )

        df = user_data.copy()

        if auditor_filter != "All":
            df = df[df["Auditor"] == auditor_filter]

        if area_filter != "All":
            df = df[df["Area"] == area_filter]

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Audits", len(df))
        col2.metric("Avg Score", round(df["Score"].mean(), 1))
        col3.metric("Areas", df["Area"].nunique())

        st.bar_chart(df.groupby("Auditor")["Score"].mean())

        df["Date"] = pd.to_datetime(df["Date"])
        st.line_chart(df.groupby("Date")["Score"].mean())

        st.dataframe(df)

# -----------------------------
# 📁 EXCEL VIEWER
# -----------------------------
elif page == "📁 Excel Data Viewer":
    st.header("Excel Data")

    sheet = st.selectbox("Select Sheet", xls.sheet_names)

    df_excel = pd.read_excel(excel_file, sheet_name=sheet)

    st.dataframe(df_excel)

# -----------------------------
# 👥 ALL NAMES PAGE
# -----------------------------
elif page == "👥 All Names":
    st.header("All Names Found")

    st.write(f"Total: {len(name_list)}")
    st.dataframe(pd.DataFrame(name_list, columns=["Names"]))
