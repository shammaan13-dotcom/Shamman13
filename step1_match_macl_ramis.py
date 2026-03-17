# ============================================================
# STEP 1 — MACL → RAMIS COPY (FINAL STABLE VERSION)
# ============================================================

from openpyxl import load_workbook
import os


# ------------------------------------------------------------
# CONFIG (STREAMLIT SAFE)
# ------------------------------------------------------------

MACL_FILE = os.getenv("MACL_FILE", "input/macl_master.xlsx")

OUTPUT_FILE = os.getenv(
    "OUTPUT_FILE",
    os.path.join("output", "MACL_RAMIS_MATCHED.xlsx")
)

os.makedirs("output", exist_ok=True)


# ------------------------------------------------------------
# LOAD FILE
# ------------------------------------------------------------

try:
    wb = load_workbook(MACL_FILE)
except Exception as e:
    raise Exception(f"Error loading MACL file: {e}")

ws = wb.active

print("File loaded successfully")


# ------------------------------------------------------------
# CLEAR RAMIS SECTION (I–N)
# ------------------------------------------------------------

for r in range(3, ws.max_row + 1):
    for c in range(9, 15):  # I to N
        ws.cell(row=r, column=c).value = None


# ------------------------------------------------------------
# COPY MACL (A–G) → RAMIS (I–N)
# ------------------------------------------------------------

for r in range(3, ws.max_row + 1):

    airline = ws.cell(row=r, column=1).value
    day     = ws.cell(row=r, column=2).value
    flt     = ws.cell(row=r, column=4).value
    sta     = ws.cell(row=r, column=5).value
    std     = ws.cell(row=r, column=6).value
    eff     = ws.cell(row=r, column=7).value

    # skip empty rows
    if airline is None or str(airline).strip() == "":
        continue

    ws.cell(row=r, column=9).value  = airline   # I
    ws.cell(row=r, column=10).value = day       # J
    ws.cell(row=r, column=11).value = flt       # K
    ws.cell(row=r, column=12).value = sta       # L
    ws.cell(row=r, column=13).value = std       # M
    ws.cell(row=r, column=14).value = eff       # N


print("RAMIS section filled successfully")


# ------------------------------------------------------------
# SAVE OUTPUT
# ------------------------------------------------------------

try:
    wb.save(OUTPUT_FILE)
except Exception as e:
    raise Exception(f"Error saving output file: {e}")

print(f"STEP 1 COMPLETE: {OUTPUT_FILE}")
