import pandas as pd
from datetime import datetime
import os

FILE_NAME = "Audit Schedule - Internal - LPA.xlsx"
LOG_SHEET = "Audit_Log"

# -----------------------------
# MASTER LISTS (FROM YOUR FILE)
# -----------------------------

AUDITORS = [
    "Freddie Gamble", "Anthony Wall", "Tim Kass", "Bryan Profit",
    "Reggie Coleman", "Miguel Frias", "Jacoby Alston",
    "Art DiFilippo", "Jarvis Adelabu", "Brett Meyer"
]

AREAS = ["Maintenance", "Carbon", "Cast House", "Potline", "Environmental"]

LOCATIONS = [
    "PL 136 / Cruce cleaner","PL Room 1 West","PL Room 2 West","PL Room 3 West","PL Room 4 West",
    "PL Room 1 East","PL Room 2 East","PL Room 3 East","PL Room 4 East",
    "PL Services/ Ore Unloading","PL 138 Repair/Rebuild",
    "CB Coke Tank","CB Pitch Tanks","CB Green Mill","CB Rod Shop","CB Bake Oven North","CB Bake Oven South",
    "CH Moldshop","CH Laboratory","CH Shipping/Scales",
    "CH Casting Pits/Billet Inspection Line","CH Dross Room/Furnaces/Sow Line",
    "CH Homogen Furnaces/Pre-Heat/Cooling",
    "MT Crane Shop","MT Fabrication Shop","MT Rebuild Shop","MT Mobile Shop",
    "MT Utilities Shop","MT Compressor Room","MT 044 Warehouse",
    "ENV 161 East Baghouse","ENV 161 West Baghouse",
    "ENV 162 East Baghouse","ENV 162 West Baghouse",
    "ENV Waste Fluid Area","ENV 138 CAA Building","ENV 261G Bake Oven Scrubber"
]

AUDIT_TYPES = ["HK", "Safe Obs", "PPE", "LOTO", "Mobile"]

# -----------------------------
# CREATE MASTER LOG IF MISSING
# -----------------------------

def initialize_log():
    if not os.path.exists(FILE_NAME):
        print("Excel file not found.")
        return

    try:
        df = pd.read_excel(FILE_NAME, sheet_name=LOG_SHEET, engine="openpyxl")
    except:
        df = pd.DataFrame(columns=[
            "Audit ID", "Date Entered", "Week Of",
            "Auditor Name", "Area", "Location",
            "Audit Type", "Score (%)", "Completed"
        ])
        with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=LOG_SHEET, index=False)

# -----------------------------
# ADD NEW AUDIT
# -----------------------------

def add_audit(auditor, area, location, audit_type, week_of, score, completed):
    # VALIDATION
    if auditor not in AUDITORS:
        raise ValueError("Invalid auditor name")

    if area not in AREAS:
        raise ValueError("Invalid area")

    if location not in LOCATIONS:
        raise ValueError("Invalid location")

    if audit_type not in AUDIT_TYPES:
        raise ValueError("Invalid audit type")

    if not (1 <= score <= 100):
        raise ValueError("Score must be between 1-100")

    if completed.lower() != "c":
        raise ValueError("Completed must be 'C' or 'c'")

    df = pd.read_excel(FILE_NAME, sheet_name=LOG_SHEET, engine="openpyxl")

    next_id = len(df) + 1

    new_row = {
        "Audit ID": next_id,
        "Date Entered": datetime.now().strftime("%Y-%m-%d"),
        "Week Of": week_of,
        "Auditor Name": auditor,
        "Area": area,
        "Location": location,
        "Audit Type": audit_type,
        "Score (%)": score,
        "Completed": "C"
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=LOG_SHEET, index=False)

    print("✅ Audit added successfully!")

# -----------------------------
# VIEW HISTORY
# -----------------------------

def view_history():
    df = pd.read_excel(FILE_NAME, sheet_name=LOG_SHEET, engine="openpyxl")

    df["Week Of"] = df["Week Of"].apply(lambda x: f"Week of: {x}")

    print("\n===== AUDIT HISTORY =====\n")
    print(df.to_string(index=False))

# -----------------------------
# SIMPLE USER INPUT SYSTEM
# -----------------------------

def run():
    initialize_log()

    while True:
        print("\n1. Add New Audit")
        print("2. View Audit History")
        print("3. Exit")

        choice = input("Select option: ")

        if choice == "1":
            print("\n--- New Audit ---")

            auditor = input("Auditor Name: ")
            area = input("Area: ")
            location = input("Location: ")
            audit_type = input("Audit Type: ")
            week_of = input("Week of (ex: 6/1/26 - 6/5/26): ")
            score = int(input("Score (1-100): "))
            completed = input("Completed (C/c): ")

            try:
                add_audit(auditor, area, location, audit_type, week_of, score, completed)
            except Exception as e:
                print(f"❌ Error: {e}")

        elif choice == "2":
            view_history()

        elif choice == "3":
            break

        else:
            print("Invalid choice")

# -----------------------------
# RUN PROGRAM
# -----------------------------

if __name__ == "__main__":
    run()
