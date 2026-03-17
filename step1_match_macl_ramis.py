# ============================================================
# RAMIS PURPLE SIDE – STREAMLIT PRODUCTION VERSION
# ============================================================

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from collections import defaultdict
from datetime import datetime
import os


# ------------------------------------------------------------
# CONFIG (STREAMLIT SAFE)
# ------------------------------------------------------------

RAMIS_FILE = os.getenv("RAMIS_FILE", "input/ramis.xlsx")
MACL_FILE  = os.getenv("MACL_FILE", "input/macl_master.xlsx")

OUTPUT_FILE = os.getenv(
    "OUTPUT_FILE",
    os.path.join("output", "MASTER_MACL_WINTER_vs_RAMIS_UPDATED.xlsx")
)

os.makedirs("output", exist_ok=True)


# ------------------------------------------------------------
# TMA TIME WINDOWS
# ------------------------------------------------------------

ARR_START = datetime.strptime("04:45","%H:%M").time()
ARR_END   = datetime.strptime("15:45","%H:%M").time()

DEP_START = datetime.strptime("09:00","%H:%M").time()
DEP_END   = datetime.strptime("23:59","%H:%M").time()


# ------------------------------------------------------------
# TIME PARSER
# ------------------------------------------------------------

def parse_time(t):
    if t is None:
        return None
    t = str(t).strip()
    if t == "" or t == "-":
        return None
    try:
        return datetime.strptime(t,"%H:%M").time()
    except:
        return None


def correct_midnight_std(std):
    if std is None:
        return None
    if std.hour in [0,1]:
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
        return None, None, None

    if arr_valid and dep_valid:
        if "-" not in flt:
            flt_out = flt + "-D"
        else:
            flt_out = flt
        return flt_out, sta, std_t.strftime("%H:%M")

    if arr_valid:
        return flt.split("-")[0], sta, ""

    if dep_valid:
        if "-" in flt:
            return flt.split("-")[1], "", std_t.strftime("%H:%M")
        return flt, "", std_t.strftime("%H:%M")


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
# LOAD FILES
# ------------------------------------------------------------

ramis_wb = load_workbook(RAMIS_FILE, data_only=True)
master_wb = load_workbook(MACL_FILE)

ramis_ws = ramis_wb.active
target_ws = master_wb.active


# ------------------------------------------------------------
# STYLES
# ------------------------------------------------------------

purple_fill = PatternFill(start_color="D9CCE3", end_color="D9CCE3", fill_type="solid")
header_font = Font(bold=True)
center_align = Alignment(horizontal="center", vertical="center")


# ------------------------------------------------------------
# GROUP DATA
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


# SORT
for d in week_groups:
    week_groups[d] = sorted(week_groups[d], key=lambda x: str(x[0]).upper())


# ------------------------------------------------------------
# CLEAR OLD SECTION
# ------------------------------------------------------------

for r in range(3, target_ws.max_row + 1):
    for c in range(9, 15):
        target_ws.cell(r, c).value = None


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
# SAVE
# ------------------------------------------------------------

master_wb.save(OUTPUT_FILE)

print("Purple section updated successfully")
