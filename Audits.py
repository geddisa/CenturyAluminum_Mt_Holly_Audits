import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import os

FILE_NAME = "Audit Schedule - Internal - LPA.xlsx"
LOG_SHEET = "Audit_Log"

# -----------------------------
# CREATE LOG SHEET IF MISSING
# -----------------------------
def ensure_log_sheet():
    wb = load_workbook(FILE_NAME)

    if LOG_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(LOG_SHEET)

        headers = [
            "Audit ID", "Date Entered", "Week Of",
            "Auditor Name", "Area", "Location",
            "Audit Type", "Score", "Completed"
        ]

        ws.append(headers)
        wb.save(FILE_NAME)
        print("✅ Created Audit_Log sheet")

# -----------------------------
# ADD AUDIT
# -----------------------------
def add_audit():
    wb = load_workbook(FILE_NAME)
    ws = wb[LOG_SHEET]

    print("\n--- ADD NEW AUDIT ---")

    auditor = input("Auditor Name: ")
    area = input("Area: ")
    location = input("Location: ")
    audit_type = input("Audit Type: ")
    week = input("Week (ex: 6/1/26 - 6/5/26): ")
    score = int(input("Score (1-100): "))
    completed = input("Completed (C): ")

    # ✅ FORCE RULES
    if not (1 <= score <= 100):
        print("❌ Score must be 1–100")
        return

    if completed.lower() != "c":
        print("❌ Must be C")
        return

    next_row = ws.max_row + 1

    ws.cell(next_row, 1, next_row - 1)
    ws.cell(next_row, 2, datetime.now().strftime("%Y-%m-%d"))
    ws.cell(next_row, 3, f"Week of: {week}")
    ws.cell(next_row, 4, auditor)
    ws.cell(next_row, 5, area)
    ws.cell(next_row, 6, location)
    ws.cell(next_row, 7, audit_type)
    ws.cell(next_row, 8, score)
    ws.cell(next_row, 9, "C")

    wb.save(FILE_NAME)

    print("✅ Audit saved successfully!")

# -----------------------------
# VIEW HISTORY
# -----------------------------
def view_history():
    try:
        df = pd.read_excel(FILE_NAME, sheet_name=LOG_SHEET)
        print("\n===== HISTORY =====")
        print(df)
    except:
        print("❌ No history found yet")

# -----------------------------
# MAIN MENU
# -----------------------------
def run():
    if not os.path.exists(FILE_NAME):
        print("❌ Excel file not found!")
        return

    ensure_log_sheet()

    while True:
        print("\n--- MENU ---")
        print("1 = Add Audit")
        print("2 = View History")
        print("3 = Exit")

        choice = input("Choice: ")

        if choice == "1":
            add_audit()
        elif choice == "2":
            view_history()
        elif choice == "3":
            break
        else:
            print("Invalid")

# -----------------------------
run()
