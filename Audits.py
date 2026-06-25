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
st.write("Sheets:", xls.sheet_names)


# Sidebar navigation
page = st.sidebar.radio(
    "Navigation",
    ["📊 Dashboard", "📋 Enter Audit", "📁 Excel Data Viewer"]
)

# -----------------------------
# LOAD USER DATA (CSV STORAGE)
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
# 📋 ENTER AUDIT PAGE
# -----------------------------
if page == "📋 Enter Audit":
    st.header("Enter New Audit")

    with st.form("audit_form"):
        col1, col2 = st.columns(2)

        with col1:
            auditor = st.selectbox("Auditor", [
                "Freddie Gamble",
                "Anthony Wall",
                "Tim Kass",
                "Bryan Profit",
                "Reggie Coleman",
                "Miguel Frias"
            ])

            area = st.selectbox("Area", [
                "Maintenance",
                "Carbon",
                "Cast House",
                "Potline",
                "Environmental"
            ])

        with col2:
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
# 📊 DASHBOARD PAGE
# -----------------------------
elif page == "📊 Dashboard":
    st.header("Audit Dashboard")

    if user_data.empty:
        st.warning("No user-entered audits yet.")
    else:
        # Filters
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

        # KPI Metrics
        col1, col2, col3 = st.columns(3)

        col1.metric("Total Audits", len(df))

        if "Score" in df:
            col2.metric("Avg Score", round(df["Score"].mean(), 1))

        col3.metric("Areas Covered", df["Area"].nunique())

        # Charts
        st.subheader("Performance by Auditor")
        st.bar_chart(df.groupby("Auditor")["Score"].mean())

        st.subheader("Trend Over Time")
        df["Date"] = pd.to_datetime(df["Date"])
        st.line_chart(df.groupby("Date")["Score"].mean())

        # Raw data
        st.subheader("Audit Records")
        st.dataframe(df)

# -----------------------------
# 📁 EXCEL DATA VIEWER
# -----------------------------
elif page == "📁 Excel Data Viewer":
    st.header("Original Excel Data")

    selected_sheet = st.selectbox("Select Sheet", xls.sheet_names)

    df_excel = pd.read_excel(excel_file, sheet_name=selected_sheet)

    st.subheader(f"Sheet: {selected_sheet}")
    st.dataframe(df_excel)

    # Optional quick insights
    if not df_excel.select_dtypes(include="number").empty:
        st.subheader("Quick Chart")
        st.line_chart(df_excel.select_dtypes(include="number"))
