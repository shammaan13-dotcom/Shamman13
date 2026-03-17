# ==========================================================
# RAMIS CLEAN SHEET FINAL ROTATION ENGINE (LOCAL VERSION)
# ==========================================================

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
import os


# ----------------------------------------------------------
# CONFIG
# ----------------------------------------------------------

INPUT_FILE = "connecting.xlsx"
OUTPUT_FILE = "output/RAMIS_ROTATION_FINAL.xlsx"

CONFIG = {
    "SEASON_START": "2025-10-26",
    "SEASON_END": "2026-03-28",
    "ARR_START": "04:45:00",
    "ARR_END": "15:31:00",
    "DEP_START": "09:00:00",
    "DEP_END": "23:59:00",
    "LOOKBACK_DAYS": 7
}


# ----------------------------------------------------------
# LOAD FILE
# ----------------------------------------------------------

if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(f"Missing file: {INPUT_FILE}")

df = pd.read_excel(INPUT_FILE, dtype=str)
df.columns = df.columns.str.strip()


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

df["Scheduled Time"] = df["Scheduled Time"].astype(str).str.split(".").str[0]

df["Scheduled Time"] = pd.to_datetime(
    df["Scheduled Time"],
    format="%H:%M:%S",
    errors="coerce"
)

df["Prefix"] = df["Flight ID"].str[:2]


# ----------------------------------------------------------
# SEASON FILTER
# ----------------------------------------------------------

SEASON_START = pd.to_datetime(CONFIG["SEASON_START"])
SEASON_END = pd.to_datetime(CONFIG["SEASON_END"])

df = df[
    (df["Start Date"] <= SEASON_END) &
    (df["End Date"] >= SEASON_START)
].copy()


# ----------------------------------------------------------
# FLIGHT ID BUILDER
# ----------------------------------------------------------

def build_flight_id(arr_id, dep_id):

    arr_prefix=''.join(filter(str.isalpha,arr_id))
    arr_num=arr_id.replace(arr_prefix,"")
    dep_num=dep_id.replace(arr_prefix,"")

    if dep_id.startswith(arr_id):
        return f"{arr_id}-{dep_id[len(arr_id):]}"

    return f"{arr_id}-{dep_num}"


# ----------------------------------------------------------
# MATCHING ENGINE
# ----------------------------------------------------------

output_rows = []
used_flights = set()

for (prefix, day), group in df.groupby(["Prefix","Scheduled Day"]):

    arrivals = group[group["Type"]=="ARRIVAL"].copy()
    departures = group[group["Type"]=="DEPARTURE"].copy()

    arrivals = arrivals.sort_values(["Start Date","Scheduled Time"])
    departures = departures.sort_values(["Start Date","Scheduled Time"])

    for _, arr in arrivals.iterrows():

        for _, dep in departures.iterrows():

            if dep["Scheduled Time"] <= arr["Scheduled Time"]:
                continue

            if arr["Start Date"] != dep["Start Date"]:
                continue

            if arr["End Date"] != dep["End Date"]:
                continue

            output_rows.append({
                "AIRLINE": prefix,
                "DAYS OF OPS": day,
                "FLT NO": build_flight_id(arr["Flight ID"],dep["Flight ID"]),
                "STA": arr["Scheduled Time"].strftime("%H:%M"),
                "STD": dep["Scheduled Time"].strftime("%H:%M"),
                "EFFECTIVE": arr["Start Date"].strftime("%d.%m.%y")
                + " - " +
                arr["End Date"].strftime("%d.%m.%y")
            })

            used_flights.add(arr["Flight ID"])
            used_flights.add(dep["Flight ID"])
            break


# ----------------------------------------------------------
# FINAL DF
# ----------------------------------------------------------

ramis_final = pd.DataFrame(output_rows)

ramis_final = ramis_final.drop_duplicates().reset_index(drop=True)


# ----------------------------------------------------------
# BUILD OUTPUT FILE
# ----------------------------------------------------------

wb = Workbook()
ws = wb.active
ws.title = "RAMIS CLEAN SHEET"

headers = ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]

for col,header in enumerate(headers,start=1):
    ws.cell(row=1,column=col).value = header
    ws.cell(row=1,column=col).font = Font(bold=True)

for r_idx, row in ramis_final.iterrows():
    for c_idx, val in enumerate(row, start=1):
        ws.cell(row=r_idx+2, column=c_idx).value = val


# ----------------------------------------------------------
# SAVE FILE
# ----------------------------------------------------------

os.makedirs("output", exist_ok=True)

wb.save(OUTPUT_FILE)

print(f"✅ RAMIS clean file saved: {OUTPUT_FILE}")
