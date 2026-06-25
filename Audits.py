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

# -----------------------------
# ✅ GET ONLY REAL PERSON NAMES
# -----------------------------
all_names = set()

COMMON_NON_NAMES = ["Room", "West", "East", "Side", "Shop", "Tank", "Mill"]

for sheet in xls.sheet_names:
    df_temp = pd.read_excel(excel_file, sheet_name=sheet)

    for col in df_temp.columns:
        for val in df_temp[col].dropna():
            if isinstance(val, str):
                val = val.strip()
                words = val.split()

                if (
                    len(words) == 2                            # exactly 2 words
                    and all(w[0].isupper() for w in words)     # capitalized
                    and all(w.isalpha() for w in words)        # letters only
                    and not any(w in COMMON_NON_NAMES for w in words)
                ):
                    all_names.add(val)

# convert to sorted list
name_list = sorted(all_names)

# -----------------------------
# SIDEBAR NAV
# -----------------------------
page = st.sidebar.radio(
    "Navigation",
    ["📊 Dashboard", "📋 Enter Audit", "📁 Excel Viewer", "👥 Names"]
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
        auditor = st.selectbox("Auditor", name_list)

        area = st.selectbox("Area", [
            "Maintenance", "Carbon", "Cast House", "Potline", "Environmental"
        ])

        audit_type = st.selectbox("Audit Type", [
            "LPA", "Safe Observation", "PPE", "LOTO", "Mobile Equipment"
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
        auditor_filter = st.sidebar.selectbox(
            "Auditor", ["All"] + list(user_data["Auditor"].unique())
        )

        area_filter = st.sidebar.selectbox(
            "Area", ["All"] + list(user_data["Area"].unique())
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

        st.subheader("Performance by Auditor")
        st.bar_chart(df.groupby("Auditor")["Score"].mean())

        st.subheader("Trend Over Time")
        df["Date"] = pd.to_datetime(df["Date"])
        st.line_chart(df.groupby("Date")["Score"].mean())

        st.dataframe(df)

# -----------------------------
# 📁 EXCEL VIEWER
# -----------------------------
elif page == "📁 Excel Viewer":
    st.header("Excel Data")

    sheet = st.selectbox("Select Sheet", xls.sheet_names)
    df_excel = pd.read_excel(excel_file, sheet_name=sheet)

    st.dataframe(df_excel)

# -----------------------------
# 👥 VIEW ALL NAMES
# -----------------------------
elif page == "👥 Names":
    st.header("All Employee Names")

    st.write(f"Total Names Found: {len(name_list)}")
    st.dataframe(pd.DataFrame(name_list, columns=["Names"]))
