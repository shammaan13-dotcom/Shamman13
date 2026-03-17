# ============================================================
# RAMIS PURPLE SIDE – PRODUCTION ENGINE (LOCAL VERSION)
# ============================================================

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from collections import defaultdict
from datetime import datetime
import os


# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

RAMIS_FILE = "output/RAMIS_ROTATION_FINAL.xlsx"
MASTER_FILE = "output/MACL_RAMIS_MATCHED.xlsx"
OUTPUT_FILE = "output/MASTER_MACL_UPDATED.xlsx"


# ------------------------------------------------------------
# VALIDATE FILES
# ------------------------------------------------------------

try:
    ramis_wb = load_workbook(RAMIS_FILE, data_only=True)
except Exception as e:
    raise Exception(f"Error loading RAMIS file: {e}")

try:
    master_wb = load_workbook(MASTER_FILE)
except Exception as e:
    raise Exception(f"Error loading MASTER file: {e}")


# ------------------------------------------------------------
# TIME WINDOWS
# ------------------------------------------------------------

ARR_START = datetime.strptime("04:45","%H:%M").time()
ARR_END   = datetime.strptime("15:45","%H:%M").time()

DEP_START = datetime.strptime("09:00","%H:%M").time()
DEP_END   = datetime.strptime("23:59","%H:%M").time()


# ------------------------------------------------------------
# TIME PARSER
# ------------------------------------------------------------

def parse_time(t):

    if not t:
        return None

    try:
        return datetime.strptime(str(t), "%H:%M").time()
    except:
        return None


# ------------------------------------------------------------
# MIDNIGHT FIX
# ------------------------------------------------------------

def correct_midnight_std(std):

    if std and (std.hour == 0 or std.hour == 1):
        return datetime.strptime("23:59","%H:%M").time()

    return std


# ------------------------------------------------------------
# VALIDATION
# ------------------------------------------------------------

def validate_flight(flt, sta, std):

    sta_t = parse_time(sta)
    std_t = parse_time(std)
    std_t = correct_midnight_std(std_t)

    arr_valid = sta_t and ARR_START <= sta_t <= ARR_END
    dep_valid = std_t and DEP_START <= std_t <= DEP_END

    if not arr_valid and not dep_valid:
        return None,None,None

    if arr_valid and dep_valid:
        if "-" not in flt:
            flt_out = flt + "-D" if not flt.endswith("D") else flt[:-1] + "-D"
        else:
            flt_out = flt

        return flt_out, sta, std_t.strftime("%H:%M")

    if arr_valid:
        return flt.split("-")[0], sta, ""

    if dep_valid:
        return flt.split("-")[-1], "", std_t.strftime("%H:%M")


# ------------------------------------------------------------
# NORMALIZE DAY
# ------------------------------------------------------------

def normalize_day(day):

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
# LOAD FILES (STREAMLIT SAFE)
# ------------------------------------------------------------

week_groups = defaultdict(list)

for row in ramis_ws.iter_rows(min_row=2, max_col=6, values_only=True):

    if not any(row):
        continue

    airline, day, flt, sta, std, eff = row

    weekday = normalize_day(day)
    if not weekday:
        continue

    validated = validate_flight(flt, sta, std)

    if validated[0] is None:
        continue

    new_flt, new_sta, new_std = validated

    week_groups[weekday].append((airline, day, new_flt, new_sta, new_std, eff))

# 👇 ADD HERE
print("TOTAL RECORDS:", sum(len(v) for v in week_groups.values()))

# ------------------------------------------------------------
# SORT
# ------------------------------------------------------------

for day in week_groups:
    week_groups[day] = sorted(week_groups[day], key=lambda x: str(x[0]))


# ------------------------------------------------------------
# CLEAR PURPLE
# ------------------------------------------------------------

for r in range(3, target_ws.max_row + 1):
    for c in range(9, 15):
        target_ws.cell(r, c).value = None


# ------------------------------------------------------------
# WRITE DATA
# ------------------------------------------------------------

WEEKDAYS = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"]

current_row = 3

for i, day in enumerate(WEEKDAYS):

    if day not in week_groups:
        continue

    if i != 0:
        current_row += 1
        target_ws.cell(current_row,9).value = day
        current_row += 1

    for record in week_groups[day]:

        for offset, val in enumerate(record):
            target_ws.cell(current_row, 9 + offset).value = val

        current_row += 1


# ------------------------------------------------------------
# SAVE OUTPUT
# ------------------------------------------------------------

import os
os.makedirs("output", exist_ok=True)

master_wb.save(OUTPUT_FILE)

print(f"✅ Purple section updated: {OUTPUT_FILE}")
