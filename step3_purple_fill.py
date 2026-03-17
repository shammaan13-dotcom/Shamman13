# ============================================================
# STEP 3 — RAMIS PURPLE FILL (FINAL WORKING VERSION)
# ============================================================

from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import os


# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

RAMIS_FILE  = "output/RAMIS_ROTATION_FINAL.xlsx"
MASTER_FILE = "output/MACL_RAMIS_MATCHED.xlsx"
OUTPUT_FILE = "output/MASTER_MACL_UPDATED.xlsx"


# ------------------------------------------------------------
# VALIDATE FILES
# ------------------------------------------------------------

if not os.path.exists(RAMIS_FILE):
    raise FileNotFoundError(f"Missing RAMIS file: {RAMIS_FILE}")

if not os.path.exists(MASTER_FILE):
    raise FileNotFoundError(f"Missing MASTER file: {MASTER_FILE}")


# ------------------------------------------------------------
# LOAD FILES
# ------------------------------------------------------------

ramis_wb = load_workbook(RAMIS_FILE, data_only=True)
master_wb = load_workbook(MASTER_FILE)

ramis_ws = ramis_wb.active
target_ws = master_wb.active

print("FILES LOADED")


# ------------------------------------------------------------
# NORMALIZE DAY
# ------------------------------------------------------------

def normalize_day(day):

    if not day:
        return None

    d = str(day).strip().upper()

    mapping = {
        "MON":"MONDAY","MONDAY":"MONDAY",
        "TUE":"TUESDAY","TUESDAY":"TUESDAY",
        "WED":"WEDNESDAY","WEDNESDAY":"WEDNESDAY",
        "THU":"THURSDAY","THURSDAY":"THURSDAY",
        "FRI":"FRIDAY","FRIDAY":"FRIDAY",
        "SAT":"SATURDAY","SATURDAY":"SATURDAY",
        "SUN":"SUNDAY","SUNDAY":"SUNDAY"
    }

    return mapping.get(d)


# ------------------------------------------------------------
# GROUP RAMIS DATA
# ------------------------------------------------------------

week_groups = defaultdict(list)

for row in ramis_ws.iter_rows(min_row=2, max_col=6, values_only=True):

    if not any(row):
        continue

    airline, day, flt, sta, std, eff = row

    weekday = normalize_day(day)

    if not weekday:
        continue

    week_groups[weekday].append(
        (airline, day, flt, sta, std, eff)
    )


print("TOTAL RAMIS RECORDS:", sum(len(v) for v in week_groups.values()))


# ------------------------------------------------------------
# CLEAR OLD PURPLE SECTION (I–N)
# ------------------------------------------------------------

for r in range(3, target_ws.max_row + 1):
    for c in range(9, 15):
        target_ws.cell(r, c).value = None


# ------------------------------------------------------------
# WRITE RAMIS DATA INTO MASTER (I–N)
# ------------------------------------------------------------

current_row = 3

WEEKDAYS = [
    "MONDAY","TUESDAY","WEDNESDAY",
    "THURSDAY","FRIDAY","SATURDAY","SUNDAY"
]

for day in WEEKDAYS:

    if day not in week_groups:
        continue

    for record in week_groups[day]:

        for offset, val in enumerate(record):
            target_ws.cell(current_row, 9 + offset).value = val

        current_row += 1


# ------------------------------------------------------------
# SAVE OUTPUT
# ------------------------------------------------------------

os.makedirs("output", exist_ok=True)

master_wb.save(OUTPUT_FILE)

print("STEP 3 COMPLETE — PURPLE FILLED")
print(f"OUTPUT: {OUTPUT_FILE}")
