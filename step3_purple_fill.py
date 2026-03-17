# ============================================================
# STEP 3 — READ RAMIS FROM SAME SHEET (I–N) AND FORMAT PURPLE
# ============================================================

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from collections import defaultdict
import os


# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

MACL_FILE = os.getenv("MACL_FILE", "input/macl_master.xlsx")

OUTPUT_FILE = os.getenv(
    "OUTPUT_FILE",
    os.path.join("output", "MASTER_MACL_WINTER_vs_RAMIS_UPDATED.xlsx")
)

os.makedirs("output", exist_ok=True)


# ------------------------------------------------------------
# LOAD FILE
# ------------------------------------------------------------

wb = load_workbook(MACL_FILE)
ws = wb.active

print("File loaded")


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
# STEP 3 FINAL FIXED (STORE → CLEAR → WRITE)
# ------------------------------------------------------------

from collections import defaultdict

week_groups = defaultdict(list)

# ------------------------------------------------------------
# STEP 1: STORE DATA FROM I–N
# ------------------------------------------------------------

for r in range(3, ws.max_row + 1):

    airline = ws.cell(r, 9).value
    day     = ws.cell(r, 10).value
    flt     = ws.cell(r, 11).value
    sta     = ws.cell(r, 12).value
    std     = ws.cell(r, 13).value
    eff     = ws.cell(r, 14).value

    if not airline:
        continue

    weekday = normalize_day(day)
    if not weekday:
        continue

    week_groups[weekday].append((airline, day, flt, sta, std, eff))


# DEBUG
print("TOTAL RECORDS:", sum(len(v) for v in week_groups.values()))


# ------------------------------------------------------------
# STEP 2: CLEAR OLD DATA (I–N)
# ------------------------------------------------------------

for r in range(3, ws.max_row + 1):
    for c in range(9, 15):
        ws.cell(r, c).value = None


# ------------------------------------------------------------
# STEP 3: WRITE BACK STRUCTURED DATA
# ------------------------------------------------------------

row_ptr = 3

WEEKDAYS = [
    "MONDAY","TUESDAY","WEDNESDAY",
    "THURSDAY","FRIDAY","SATURDAY","SUNDAY"
]

for day in WEEKDAYS:

    if day not in week_groups:
        continue

    # DAY HEADER
    ws.cell(row_ptr, 9).value = day
    row_ptr += 1

    # COLUMN HEADERS
    headers = ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]

    for i, h in enumerate(headers):
        ws.cell(row_ptr, 9+i).value = h

    row_ptr += 1

    # DATA ROWS
    for record in week_groups[day]:
        for i, val in enumerate(record):
            ws.cell(row_ptr, 9+i).value = val
        row_ptr += 1

    row_ptr += 1


# ------------------------------------------------------------
# SAVE
# ------------------------------------------------------------

wb.save(OUTPUT_FILE)

print("STEP 3 COMPLETE — PURPLE FORMATTED")
