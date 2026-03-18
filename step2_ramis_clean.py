# ==========================================================
# RAMIS CLEAN SHEET FINAL ROTATION ENGINE (FINAL VERSION)
# ==========================================================

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
import os


# ----------------------------------------------------------
# CONFIG (USE ENV VARIABLES FROM STREAMLIT)
# ----------------------------------------------------------

INPUT_FILE = os.environ.get("CONNECT_FILE") or "input/connecting_flights.xlsx"
OUTPUT_FILE = "output/RAMIS_ROTATION_FINAL.xlsx"

CONFIG = {
    "SEASON_START": "2025-10-26",
    "SEASON_END": "2026-03-28"
}


# ----------------------------------------------------------
# VALIDATE INPUT FILE
# ----------------------------------------------------------

if not INPUT_FILE or not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(f"Missing file: {INPUT_FILE}")


# ----------------------------------------------------------
# LOAD FILE
# ----------------------------------------------------------

df = pd.read_excel(INPUT_FILE, dtype=str)
df.columns = df.columns.str.strip()

print("FILE LOADED:", INPUT_FILE)
print("TOTAL INPUT ROWS:", len(df))


# ----------------------------------------------------------
# CLEAN DATA
# ----------------------------------------------------------

df = df.dropna(subset=["Flight ID"])

exclude_list = [
    "ARRTBA","DEPTBA","RESARR","RESDEP",
    "MLE","MLED","DMLE","DMLED"
]

df = df[~df["Flight ID"].str.upper().isin(exclude_list)].copy()

df["Scheduled Day"] = df["Scheduled Day"].str.strip().str.upper()
df["Type"] = df["Type"].str.strip().str.upper()

df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
df["End Date"] = pd.to_datetime(df["End Date"], errors="coerce")

# CLEAN TIME FORMAT
df["Scheduled Time"] = df["Scheduled Time"].astype(str).str.split(".").str[0]

df["Scheduled Time"] = pd.to_datetime(
    df["Scheduled Time"],
    format="%H:%M:%S",
    errors="coerce"
)

# AIRLINE PREFIX (FIRST 2 CHARACTERS - LETTER OR NUMBER)
df["Prefix"] = df["Flight ID"].str[:2]


# ----------------------------------------------------------
# SEASON FILTER
# ----------------------------------------------------------

SEASON_START = pd.to_datetime(CONFIG["SEASON_START"])
SEASON_END   = pd.to_datetime(CONFIG["SEASON_END"])

df = df[
    (df["Start Date"] <= SEASON_END) &
    (df["End Date"] >= SEASON_START)
].copy()

print("AFTER FILTER:", len(df))


# ----------------------------------------------------------
# FLIGHT ID BUILDER
# ----------------------------------------------------------

def build_flight_id(arr_id, dep_id):

    arr_prefix = ''.join(filter(str.isalpha, arr_id))
    arr_num = arr_id.replace(arr_prefix, "")
    dep_num = dep_id.replace(arr_prefix, "")

    # CASE: 6E1133 + 6E1133-4 → 6E1134
    if "-" in dep_id:
        base, suffix = dep_id.split("-")
        return f"{arr_prefix}{suffix}"

    # CASE: same base
    if dep_id.startswith(arr_id):
        return f"{arr_id}-{dep_id[len(arr_id):]}"

    return f"{arr_id}-{dep_num}"


# ----------------------------------------------------------
# MATCHING ENGINE (ARR + DEP PAIRING)
# ----------------------------------------------------------

output_rows = []

for prefix, group in df.groupby(["Prefix"]):

    arrivals = group[group["Type"] == "ARRIVAL"].copy()
    departures = group[group["Type"] == "DEPARTURE"].copy()

    arrivals = arrivals.sort_values(["Start Date", "Scheduled Time"])
    departures = departures.sort_values(["Start Date", "Scheduled Time"])

    for _, arr in arrivals.iterrows():

        for _, dep in departures.iterrows():
            
        # --- STA / STD VALIDATION ---
        if pd.isna(arr["Scheduled Time"]) or pd.isna(dep["Scheduled Time"]):
            continue
        
        # STD must be after STA
        if dep["Scheduled Time"] <= arr["Scheduled Time"]:
            continue
        
        # --- DATE OVERLAP LOGIC (FIXED) ---
        if dep["Start Date"] > arr["End Date"]:
            continue
        
        if dep["End Date"] < arr["Start Date"]:
            continue

            output_rows.append({
                "AIRLINE": prefix,
                "DAYS OF OPS": day,
                "FLT NO": build_flight_id(arr["Flight ID"], dep["Flight ID"]),
                "STA": arr["Scheduled Time"].strftime("%H:%M"),
                "STD": dep["Scheduled Time"].strftime("%H:%M"),
                "EFFECTIVE": arr["Start Date"].strftime("%d.%m.%y") + " - " + arr["End Date"].strftime("%d.%m.%y")
            })

            break


print("TOTAL RAMIS RECORDS:", len(output_rows))


# ----------------------------------------------------------
# BUILD OUTPUT FILE
# ----------------------------------------------------------

wb = Workbook()
ws = wb.active
ws.title = "RAMIS CLEAN SHEET"

headers = ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]

for col, header in enumerate(headers, start=1):
    ws.cell(row=1, column=col).value = header
    ws.cell(row=1, column=col).font = Font(bold=True)

for r_idx, row in enumerate(output_rows):
    for c_idx, val in enumerate(row.values(), start=1):
        ws.cell(row=r_idx + 2, column=c_idx).value = val


# ----------------------------------------------------------
# SAVE OUTPUT
# ----------------------------------------------------------

os.makedirs("output", exist_ok=True)

wb.save(OUTPUT_FILE)

print("ARR COUNT:", len(df[df["Type"] == "ARRIVAL"]))
print("DEP COUNT:", len(df[df["Type"] == "DEPARTURE"]))

print("STEP 2 COMPLETE")
print("OUTPUT FILE:", OUTPUT_FILE)
