# ============================================================
# RAMIS → MACL RECONCILIATION ENGINE
# ============================================================

from openpyxl import load_workbook
from datetime import datetime
from google.colab import files
import re


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
# SPLIT FLIGHT PREFIX
# ------------------------------------------------------------

def split_flight(f):

    f = str(f).strip().upper()

    m = re.match(r"([A-Z0-9]+?)(\d+)$", f)

    if not m:
        return None, None, None

    prefix = m.group(1)
    number = m.group(2)

    return prefix, int(number), len(number)


# ------------------------------------------------------------
# FLIGHT MATCH FUNCTION
# ------------------------------------------------------------

def match_flight(macl, ramis):

    macl = str(macl).strip().upper()
    ramis = str(ramis).strip().upper()

    if macl == ramis:
        return True

    if "-" not in macl:
        return False

    left, right = macl.split("-")

    prefix, base, digits = split_flight(left)
    r_prefix, r_base, _ = split_flight(ramis)

    if prefix is None or r_prefix is None:
        return False

    if prefix != r_prefix:
        return False

    flights = []

    flights.append(f"{prefix}{base:0{digits}d}")

    try:

        r = int(right)

        if len(right) >= digits:
            second = r

        elif len(right) == 1:
            second = int(str(base)[:-1] + right)

        elif len(right) == 2:
            second = int(str(base)[:-2] + right)

        else:
            second = r

        flights.append(f"{prefix}{second:0{digits}d}")

    except:
        pass

    return ramis in flights


# ------------------------------------------------------------
# PARSE EFFECTIVE DATE RANGE
# ------------------------------------------------------------

def parse_eff_range(eff):

    try:

        s, e = [x.strip() for x in str(eff).split("-")]

        s = datetime.strptime(s,"%d.%m.%y")
        e = datetime.strptime(e,"%d.%m.%y")

        return s, e

    except:
        return None, None


# ------------------------------------------------------------
# FILE UPLOAD
# ------------------------------------------------------------

ramis_file = list(files.upload().keys())[0]
macl_file = list(files.upload().keys())[0]


# ------------------------------------------------------------
# LOAD FILES
# ------------------------------------------------------------

ramis_wb = load_workbook(ramis_file, data_only=True)
ramis_ws = ramis_wb.active

macl_wb = load_workbook(macl_file)
macl_ws = macl_wb.active


# ------------------------------------------------------------
# CLEAR PURPLE SECTION
# ------------------------------------------------------------

for r in range(3, macl_ws.max_row + 1):
    for c in range(9, 15):
        macl_ws.cell(r, c).value = None


# ------------------------------------------------------------
# MAIN MATCH LOOP
# ------------------------------------------------------------

for row in ramis_ws.iter_rows(min_row=2, max_col=6, values_only=True):

    airline, day, flt, sta, std, eff = row

    if not airline or not flt:
        continue

    if airline == "AIRLINE":
        continue

    airline = str(airline).strip().upper()
    day = normalize_day(day)
    flt = str(flt).strip().upper()

    r_start, r_end = parse_eff_range(eff)

    for r in range(3, macl_ws.max_row + 1):

        macl_air = macl_ws.cell(r,1).value
        macl_day = normalize_day(macl_ws.cell(r,2).value)
        macl_flt = macl_ws.cell(r,4).value
        macl_eff = macl_ws.cell(r,7).value

        if not macl_air or not macl_flt:
            continue

        macl_air = str(macl_air).strip().upper()

        if airline != macl_air:
            continue

        if day != macl_day:
            continue

        if not match_flight(macl_flt, flt):
            continue

        m_start, m_end = parse_eff_range(macl_eff)

        if m_start and r_start:
            if r_end < m_start or r_start > m_end:
                continue

        macl_ws.cell(r,9).value  = airline
        macl_ws.cell(r,10).value = day
        macl_ws.cell(r,11).value = flt
        macl_ws.cell(r,12).value = sta
        macl_ws.cell(r,13).value = std
        macl_ws.cell(r,14).value = eff

        break


# ------------------------------------------------------------
# SAVE OUTPUT
# ------------------------------------------------------------

output_file = "MACL_RAMIS_MATCHED.xlsx"

macl_wb.save(output_file)

files.download(output_file)
