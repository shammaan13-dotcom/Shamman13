import pandas as pd
from datetime import datetime, timedelta, time
from google.colab import files
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------
# UPLOAD FILE
# ------------------------------------------------

uploaded = files.upload()
input_file = list(uploaded.keys())[0]

df = pd.read_excel(input_file, header=None)
out = df.copy()

# ------------------------------------------------
# COLUMN INDEX
# ------------------------------------------------

MACL_AIRLINE = 0
MACL_DAY = 1
MACL_FLT = 3
MACL_STA = 4
MACL_STD = 5
MACL_EFF = 6

RAMIS_AIRLINE = 8
RAMIS_DAY = 9
RAMIS_FLT = 10
RAMIS_STA = 11
RAMIS_STD = 12
RAMIS_EFF = 13

RAMIS_START = 8

NEW_FLT = 14
NEW_STA = 15
NEW_STD = 16
NEW_EFF = 17
STATUS  = 18
COMMENT = 19

# ------------------------------------------------
# DAY ORDER
# ------------------------------------------------

days = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"]

def prev_day(d):
    if d not in days:
        return d
    return days[(days.index(d)-1) % 7]

# ------------------------------------------------
# TIME FUNCTIONS
# ------------------------------------------------

def to_time(v):
    try:
        return datetime.strptime(str(v), "%H:%M").time()
    except:
        return None

def base_flt(f):
    return str(f).split("-")[0]

STA_START = time(4,45)
STA_END   = time(15,45)

STD_START = time(9,0)
STD_END   = time(23,59)

def sta_allowed(t):
    return t and STA_START <= t <= STA_END

def std_allowed(t):
    return t and STD_START <= t <= STD_END

# ------------------------------------------------
# BUILD RAMIS LOOKUP
# ------------------------------------------------

ramis_rows = {}

for i,r in df.iterrows():

    airline = str(r[RAMIS_AIRLINE]).strip().upper()
    day     = str(r[RAMIS_DAY]).strip().upper()
    flt     = str(r[RAMIS_FLT]).strip().upper()
    sta     = to_time(r[RAMIS_STA])
    std     = to_time(r[RAMIS_STD])

    if std == time(23,59):
        day = prev_day(day)

    if airline and day and flt and flt != "NAN":

        if airline not in ramis_rows:
            ramis_rows[airline] = []

        ramis_rows[airline].append((day,flt,sta,std,i))

used_ramis = set()

# ------------------------------------------------
# MATCH MACL WITH RAMIS
# ------------------------------------------------

for i,r in df.iterrows():

    airline  = str(r[MACL_AIRLINE]).strip().upper()
    macl_day = str(r[MACL_DAY]).strip().upper()
    macl_flt = str(r[MACL_FLT]).strip().upper()

    macl_sta = to_time(r[MACL_STA])
    macl_std = to_time(r[MACL_STD])

    macl_norm = macl_flt.replace("-","")
    macl_base = base_flt(macl_flt)

    found = False

    for r_day,r_flt,r_sta,r_std,r_i in ramis_rows.get(airline,[]):

        if r_i in used_ramis:
            continue

        if macl_day != r_day:
            continue

        ramis_norm = r_flt.replace("-","")
        ramis_base = base_flt(r_flt)

        # exact match
        if macl_norm == ramis_norm:
            out.iloc[i,RAMIS_START:] = df.iloc[r_i,RAMIS_START:]
            used_ramis.add(r_i)
            found = True
            break

        # +/-1 rule
        try:
            macl_num = int(''.join(filter(str.isdigit,macl_base)))
            ramis_num = int(''.join(filter(str.isdigit,ramis_base)))

            if abs(macl_num - ramis_num) == 1 and macl_std == r_std:
                out.iloc[i,RAMIS_START:] = df.iloc[r_i,RAMIS_START:]
                used_ramis.add(r_i)
                found = True
                break
        except:
            pass

        # base match
        if macl_base == ramis_base:

            sta_ok = macl_sta == r_sta if sta_allowed(macl_sta) else True
            std_ok = macl_std == r_std if std_allowed(macl_std) else True

            if sta_ok and std_ok:
                out.iloc[i,RAMIS_START:] = df.iloc[r_i,RAMIS_START:]
                used_ramis.add(r_i)
                found = True
                break

    if not found:
        out.iloc[i,RAMIS_START:] = ""

# ------------------------------------------------
# RECONCILIATION
# ------------------------------------------------

for i,r in out.iterrows():

    macl_air  = str(r[MACL_AIRLINE]).upper()
    ramis_air = str(r[RAMIS_AIRLINE]).upper()

    macl_flt = str(r[MACL_FLT])
    macl_sta = str(r[MACL_STA])
    macl_std = str(r[MACL_STD])

    ramis_sta = str(r[RAMIS_STA])
    ramis_std = str(r[RAMIS_STD])

    macl_eff = str(r[MACL_EFF])
    ramis_eff = str(r[RAMIS_EFF])

    if "CARGO" in macl_air:

        for c in range(8,18):
            out.iloc[i,c] = ""

        out.iloc[i,STATUS] = "IGNORED"
        out.iloc[i,COMMENT] = "CARGO FLIGHT IGNORED"
        continue

    expired = False

    try:
        end_date = macl_eff.split("-")[-1].strip()
        end_date = datetime.strptime(end_date,"%d.%m.%y")

        if datetime.today() > end_date + timedelta(days=7):
            expired = True
    except:
        pass

    if expired:
        out.iloc[i,STATUS] = "EXPIRED"
        out.iloc[i,COMMENT] = "OUT OF DATE FLIGHT"
        continue

    macl_sta_time = to_time(macl_sta)
    macl_std_time = to_time(macl_std)

    sta_in_range = macl_sta_time and sta_allowed(macl_sta_time)
    std_in_range = macl_std_time and std_allowed(macl_std_time)

    if ramis_air == "" or ramis_air == "NAN":

        out.iloc[i,NEW_FLT] = macl_flt
        out.iloc[i,NEW_EFF] = macl_eff

        out.iloc[i,NEW_STA] = macl_sta if sta_in_range else ""
        out.iloc[i,NEW_STD] = macl_std if std_in_range else ""

        out.iloc[i,STATUS] = "NEW FLIGHT"
        out.iloc[i,COMMENT] = "NEW FLIGHT"

        continue

    changes = []

    if sta_in_range and macl_sta != ramis_sta:
        out.iloc[i,NEW_STA] = macl_sta
        changes.append("STA CHANGE")

    if std_in_range and macl_std != ramis_std:
        out.iloc[i,NEW_STD] = macl_std
        changes.append("STD CHANGE")

    if macl_eff != ramis_eff:
        out.iloc[i,NEW_EFF] = macl_eff
        changes.append("EFFECTIVE CHANGE")

    if changes:
        out.iloc[i,STATUS] = "CHANGE"
        out.iloc[i,COMMENT] = ", ".join(changes)
    else:
        out.iloc[i,STATUS] = "NO CHANGE"
        out.iloc[i,COMMENT] = "-"

# ------------------------------------------------
# EXPORT
# ------------------------------------------------

output_file = "ALIGNED_FLIGHTS_OUTPUT.xlsx"
out.to_excel(output_file,index=False,header=False,startrow=2)

wb = load_workbook(output_file)
ws = wb.active

# ------------------------------------------------
# INSERT EMPTY + RED SEPARATOR ROWS
# ------------------------------------------------

red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

max_row = ws.max_row

row = 3

while row <= max_row:

    macl_val = ws.cell(row=row, column=1).value
    ramis_val = ws.cell(row=row, column=9).value

    # detect empty row between data
    if macl_val in [None,"","nan","NAN"] and ramis_val in [None,"","nan","NAN"]:

        # keep the empty row

        # insert one more empty row
        ws.insert_rows(row+1)

        # insert red row
        ws.insert_rows(row+2)

        for col in range(1,21):
            ws.cell(row=row+2, column=col).fill = red_fill

        row += 3
        max_row += 2

    else:
        row += 1

ws.freeze_panes = "A3"

header_row = 2
last_col = ws.max_column
last_col_letter = get_column_letter(last_col)

ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{header_row}"

for col in ws.columns:

    max_length = 0
    column = col[0].column
    column_letter = get_column_letter(column)

    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length,len(str(cell.value)))
        except:
            pass

    ws.column_dimensions[column_letter].width = max_length + 2

    # ------------------------------------------------
# LOAD WORKBOOK
# ------------------------------------------------

wb = load_workbook(output_file)
ws = wb.active

ws.freeze_panes = "A3"

center = Alignment(horizontal="center",vertical="center")
bold   = Font(bold=True)

thin = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

blue      = PatternFill(start_color="4F81BD",end_color="4F81BD",fill_type="solid")
lightblue = PatternFill(start_color="D9E1F2",end_color="D9E1F2",fill_type="solid")
purple    = PatternFill(start_color="D9CCE3",end_color="D9CCE3",fill_type="solid")
green     = PatternFill(start_color="C6E0B4",end_color="C6E0B4",fill_type="solid")
pink      = PatternFill(start_color="F4CCCC",end_color="F4CCCC",fill_type="solid")

# ------------------------------------------------
# MACL HEADER
# ------------------------------------------------

ws.merge_cells("A1:G1")
ws["A1"]="MACL WINTER SCHEDULE"
ws["A1"].alignment=center
ws["A1"].font=bold
ws["A1"].fill=blue

macl_headers=["AIRLINE","DAY","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]

for c,h in enumerate(macl_headers,1):

    cell=ws.cell(row=2,column=c,value=h)
    cell.fill=lightblue
    cell.alignment=center
    cell.font=bold
    cell.border=thin

# ------------------------------------------------
# RAMIS HEADER
# ------------------------------------------------

ws.merge_cells("I1:N1")
ws["I1"]="RAMIS SCHEDULE TIME"
ws["I1"].alignment=center
ws["I1"].font=bold
ws["I1"].fill=purple

ramis_headers=["AIRLINE","DAY","FLT NO","STA","STD","EFFECTIVE"]

for c,h in enumerate(ramis_headers,9):

    cell=ws.cell(row=2,column=c,value=h)
    cell.fill=purple
    cell.alignment=center
    cell.font=bold
    cell.border=thin

# ------------------------------------------------
# RAMIS REVISED
# ------------------------------------------------

ws.merge_cells("O1:R1")
ws["O1"]="RAMIS REVISED TIME"
ws["O1"].alignment=center
ws["O1"].font=bold
ws["O1"].fill=green

rev_headers=["NEW FLT NO","NEW STA","NEW STD","NEW EFFECTIVE"]

for c,h in enumerate(rev_headers,15):

    cell=ws.cell(row=2,column=c,value=h)
    cell.fill=green
    cell.alignment=center
    cell.font=bold
    cell.border=thin

# ------------------------------------------------
# CHANGE SECTION
# ------------------------------------------------

ws.merge_cells("S1:T1")
ws["S1"]="CHANGES"
ws["S1"].alignment=center
ws["S1"].font=bold
ws["S1"].fill=pink

change_headers=["STATUS","COMMENTS"]

for c,h in enumerate(change_headers,19):

    cell=ws.cell(row=2,column=c,value=h)
    cell.fill=pink
    cell.alignment=center
    cell.font=bold
    cell.border=thin

# ------------------------------------------------
# FILTERS
# ------------------------------------------------

header_row = 2
last_col = ws.max_column
last_col_letter = get_column_letter(last_col)

ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{header_row}"

# ------------------------------------------------
# BORDERS
# ------------------------------------------------

hairline = Border(
    left=Side(style="hair"),
    right=Side(style="hair"),
    top=Side(style="hair"),
    bottom=Side(style="hair")
)

max_row = ws.max_row
max_col = ws.max_column

for row in range(3, max_row + 1):
    for col in range(1, max_col + 1):
        ws.cell(row=row, column=col).border = hairline

# ------------------------------------------------
# COLOR RULES
# ------------------------------------------------

sta_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
std_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
eff_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")

grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
cargo_fill = PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")
ignored_fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")

alt_green = PatternFill(start_color="F3FAF3", end_color="F3FAF3", fill_type="solid")
empty_macl_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

for row in range(3, max_row + 1):

    if row % 2 == 0:
        for col in range(1,21):
            ws.cell(row=row, column=col).fill = alt_green

    airline = str(ws.cell(row=row, column=1).value).upper()
    status  = str(ws.cell(row=row, column=19).value).upper()
    comment = str(ws.cell(row=row, column=20).value).upper()

    if "CARGO" in airline:

        for col in range(1,21):
            ws.cell(row=row, column=col).fill = cargo_fill

        continue

    if status == "EXPIRED":

        for col in range(1,21):
            ws.cell(row=row, column=col).fill = grey_fill

        continue

    if status == "IGNORED":

        for col in range(1,21):
            ws.cell(row=row, column=col).fill = ignored_fill

        continue

    if "STA CHANGE" in comment:
        ws.cell(row=row, column=16).fill = sta_fill

    if "STD CHANGE" in comment:
        ws.cell(row=row, column=17).fill = std_fill

    if "EFFECTIVE CHANGE" in comment:

        ws.cell(row=row, column=7).fill = eff_fill
        ws.cell(row=row, column=14).fill = eff_fill
        ws.cell(row=row, column=18).fill = eff_fill

# ------------------------------------------------
# AUTO WIDTH
# ------------------------------------------------

    adjusted_width = max_length + 2
    ws.column_dimensions[column_letter].width = adjusted_width

from openpyxl.utils import get_column_letter

for col in ws.columns:

    max_length = 0
    column = col[0].column
    column_letter = get_column_letter(column)

    for cell in col:

        try:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass

    adjusted_width = max_length + 2
    ws.column_dimensions[column_letter].width = adjusted_width

# ------------------------------------------------
# SAVE FILE
# ------------------------------------------------

output_file = "ALIGNED_FLIGHTS_OUTPUT.xlsx"

wb.save(output_file)

print("File saved successfully")

# ------------------------------------------------
# DOWNLOAD FILE
# ------------------------------------------------

from google.colab import files
files.download(output_file)
