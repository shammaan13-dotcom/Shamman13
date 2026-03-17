# ============================================================
# RAMIS → MACL RECONCILIATION ENGINE (STREAMLIT SAFE VERSION)
# ============================================================

from openpyxl import load_workbook
from datetime import datetime
import re
import os


# ------------------------------------------------------------
# CONFIG (STREAMLIT SAFE)
# ------------------------------------------------------------

RAMIS_FILE = os.getenv("RAMIS_FILE", "input/ramis.xlsx")
MACL_FILE  = os.getenv("MACL_FILE", "input/macl_master.xlsx")

OUTPUT_FILE = os.getenv(
    "OUTPUT_FILE",
    os.path.join("output", "MACL_RAMIS_MATCHED.xlsx")
)

os.makedirs("output", exist_ok=True)


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
# SPLIT FLIGHT
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
# MATCH FLIGHT
# ------------------------------------------------------------

def match_flight(macl, ramis):
    macl = str(macl).strip().upper() if macl else ""
    ramis = str(ramis).strip().upper() if ramis else ""

    if macl == ramis:
        return True

    parts = macl.split("-")

    if len(parts) != 2:
        return False

    left, right = parts

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
# PARSE EFFECTIVE DATE
# ------------------------------------------------------------

def parse_eff_range(eff):
    try:
        if not eff:
            return None, None

        parts = str(eff).split("-")

        if len(parts) != 2:
            return None, None

        s, e = parts

        s = datetime.strptime(s.strip(), "%d.%m.%y")
        e = datetime.strptime(e.strip(), "%d.%m.%y")

        return s, e

    except:
        return None, None


# ------------------------------------------------------------
# LOAD FILES
# ------------------------------------------------------------

try:
    ramis_wb = load_workbook(RAMIS_FILE, data_only=True)
except Exception as e:
    raise Exception(f"Error loading RAMIS file: {e}")

try:
    macl_wb = load_workbook(MACL_FILE)
except Exception as e:
    raise Exception(f"Error loading MACL file: {e}")

ramis_ws = ramis_wb.active
macl_ws = macl_wb.active

print("Files loaded successfully")


# ------------------------------------------------------------
# CLEAR RAMIS SECTION (COL 9–14)
# ------------------------------------------------------------

for r in range(3, macl_ws.max_row + 1):
    for c in range(9, 15):
        macl_ws.cell(r, c).value = None


# ------------------------------------------------------------
# MAIN MATCH LOOP (FIXED + RELIABLE WRITE)
# ------------------------------------------------------------

for row in ramis_ws.iter_rows(min_row=2, max_col=6, values_only=True):

    airline, day, flt, sta, std, eff = row

    if not airline or not flt:
        continue

    airline = str(airline).strip().upper()
    day = normalize_day(day)
    flt = str(flt).strip().upper()

    r_start, r_end = parse_eff_range(eff)

    matched = False

    for r in range(3, macl_ws.max_row + 1):

        macl_air = macl_ws.cell(r, 1).value
        macl_day = normalize_day(macl_ws.cell(r, 2).value)
        macl_flt = macl_ws.cell(r, 4).value
        macl_eff = macl_ws.cell(r, 7).value

        if not macl_air or not macl_flt:
            continue

        macl_air = str(macl_air).strip().upper()
        macl_flt_str = str(macl_flt).strip().upper()

        # --- RELAXED MATCH (CRITICAL FIX) ---
        if airline != macl_air:
            continue

        if day != macl_day:
            continue

        # primary match
        if match_flight(macl_flt_str, flt):
            pass
        # fallback match (handles EK658-9 vs EK6589)
        elif macl_flt_str.replace("-", "") == flt:
            pass
        else:
            continue

        # date check (safe)
        m_start, m_end = parse_eff_range(macl_eff)

        if m_start and r_start:
            if r_end < m_start or r_start > m_end:
                continue

        # --- WRITE (FORCE FILL) ---
        macl_ws.cell(r, 9).value  = airline
        macl_ws.cell(r, 10).value = day
        macl_ws.cell(r, 11).value = flt
        macl_ws.cell(r, 12).value = sta
        macl_ws.cell(r, 13).value = std
        macl_ws.cell(r, 14).value = eff

        matched = True
        break

    # OPTIONAL DEBUG (can remove later)
    if not matched:
        print(f"NO MATCH → {airline} {flt} {day}")


# ------------------------------------------------------------
# SAVE OUTPUT
# ------------------------------------------------------------

macl_wb.save(OUTPUT_FILE)

print(f"STEP 1 COMPLETE: {OUTPUT_FILE}")
