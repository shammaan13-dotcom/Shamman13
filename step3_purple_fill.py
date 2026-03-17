# ============================================================
# STEP 3 — RAMIS PURPLE FILL (FINAL STABLE VERSION)
# ============================================================

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from collections import defaultdict
from datetime import datetime
import os


# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

RAMIS_FILE = os.getenv("RAMIS_FILE", "input/ramis.xlsx")
MACL_FILE  = os.getenv("MACL_FILE", "input/macl_master.xlsx")

OUTPUT_FILE = os.getenv(
    "OUTPUT_FILE",
    os.path.join("output", "MASTER_MACL_WINTER_vs_RAMIS_UPDATED.xlsx")
)

os.makedirs("output", exist_ok=True)


# ------------------------------------------------------------
# LOAD FILES
# ------------------------------------------------------------

ramis_wb = load_workbook(RAMIS_FILE, data_only=True)
master_wb = load_workbook(MACL_FILE)

ramis_ws = ramis_wb.active
target_ws = master_wb.active

print("Files loaded successfully")


# ------------------------------------------------------------
# STYLES
# ------------------------------------------------------------

purple_fill = PatternFill(start_color="D9CCE3", end_color="D9CCE3", fill_type="solid")
header_font = Font(bold=True)
center_align = Alignment(horizontal="center", vertical="center")


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
# GROUP RAMIS DATA (NO VALIDATION DROP)
# ------------------------------------------------------------

week_groups = defaultdict(list)

for row in ramis_ws.iter_rows(min_row=2, max_col=6, values_only=True):

    if not any(row):
        continue

    airline, day, flt, sta, std, eff = row

    weekday = normalize_day(day)
    if not weekday:
        continue

    # ✅ NO VALIDATION DROP — ALWAYS KEEP DATA
    week_groups[weekday].append((airline, day, flt, sta, std, eff))


# DEBUG
print("TOTAL RECORDS:", sum(len(v) for v in week_groups.values()))


# ------------------------------------------------------------
# SORT DATA
# ------------------------------------------------------------

for d in week_groups:
    week_groups[d] = sorted(week_groups[d], key=lambda x: str(x[0]).upper())


# ------------------------------------------------------------
# CLEAR OLD PURPLE SECTION (I–N)
# ------------------------------------------------------------

for r in range(3, target_ws.max_row + 1):
    for c in range(9, 15):
        target_ws.cell(row=r, column=c).value = None


# ------------------------------------------------------------
# WRITE PURPLE SECTION
# ------------------------------------------------------------

WEEKDAYS = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"]

row_ptr = 3

for day in WEEKDAYS:

    if day not in week_groups:
        continue

    # Day Header
    target_ws.cell(row=row_ptr, column=9).value = day

    for c in range(9, 15):
        cell = target_ws.cell(row=row_ptr, column=c)
        cell.fill = purple_fill
        cell.font = header_font
        cell.alignment = center_align

    row_ptr += 1

    # Column Headers
    headers = ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]

    for i, h in enumerate(headers):
        cell = target_ws.cell(row=row_ptr, column=9+i)
        cell.value = h
        cell.fill = purple_fill
        cell.font = header_font
        cell.alignment = center_align

    row_ptr += 1

    # Data Rows
    for record in week_groups[day]:
        for i, val in enumerate(record):
            target_ws.cell(row=row_ptr, column=9+i).value = val
        row_ptr += 1

    row_ptr += 1


# ------------------------------------------------------------
# SAVE OUTPUT
# ------------------------------------------------------------

master_wb.save(OUTPUT_FILE)

print("STEP 3 COMPLETE — PURPLE FILLED")
