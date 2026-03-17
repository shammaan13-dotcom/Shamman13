# ==========================================================
# RAMIS CLEAN SHEET FINAL ROTATION ENGINE
# STRICT DATE MATCH + TIME WINDOW + NO DUPLICATION
# ==========================================================

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

def process(macl_df, ramis_df):

    # ------------------------------------------------
    # USER CONFIGURATION
    # ------------------------------------------------

    CONFIG = {
        "SEASON_START": "2025-10-26",
        "SEASON_END": "2026-03-28",
        "ARR_START": "04:45:00",
        "ARR_END": "15:36:00",
        "DEP_START": "09:00:00",
        "DEP_END": "23:59:00",
        "LOOKBACK_DAYS": 1
    }

    # your full logic continues here (INDENTED)

# ----------------------------------------------------------
# AIRLINE MASTER
# ----------------------------------------------------------

AIRLINE_MASTER = {
"3U":"SICHUAN AIR","4Y":"EUROWINGS","6E":"INDIGO","8D":"FITS AIR",
"AI":"AIR INDIA","AK":"AIR ASIA","AZ":"ITA AIRWAYS","B4":"FLY BEOND",
"BA":"BRITISH AIRWAYS","BS":"US-BANGLA AIRLINES","C6":"CENTRUM AIR",
"DE":"CONDOR","EK":"EMIRATES","EY":"ETIHAD AIRWAYS","FD":"THAI AIRASIA",
"FZ":"FLY DUBAI","G9":"AIR ARABIA","GF":"GULF AIR","H4":"HISKY",
"HB":"GREATER BAY AIRLINES","HX":"HONG KONG AIRLINES","J2":"AZERBAIJAN AIRLINES",
"JD":"BEIJING CAPITAL AIRLINES","KC":"AIR ASTANA","MF":"XIAMEN AIR",
"MH":"MALAYSIA AIRLINES","MP":"AFCOM CARGO","MU":"CHINA EASTHERN AIRLINES",
"NO":"NEOS SPA","OD":"BATIK AIR","OQ":"CHONGQING AIRLINES","OS":"AUSTRIAN AIRLINES",
"PG":"BANGKOK AIRWAYS","Q2":"MALDIVIAN","QR":"QATAR AIRWAYS",
"SQ":"SINGAPORE AIRLINES","SU":"AEROFLOT","SV":"SAUDI ARABIAN AIRLINES",
"TK":"TURKISH AIRLINES","UL":"SRILANKAN AIRLINES","VS":"VIRGIN ATLANTIC",
"WK":"EDELWEISS AIR","ZF":"AZUR AIR RUSSIA"
}


# ----------------------------------------------------------
# FIXED PAIR RULES
# ----------------------------------------------------------

FIXED_PAIRS = {
"QR":{"670":"671","672":"673","674":"675","676":"677"},
"EK":{"656":"657","658":"659","660":"661","652":"652D","653":"653D"},
"NO":{"110":"141"},
"UL":{"101":"102"},
"6E":{"1131":"1134","1129":"1130","1127":"1128","1133":"1132"}
}


# ----------------------------------------------------------
# FORCED MATCHING RULES (EDIT HERE)
# KEY = STA flight
# DEP = STD flight
# ----------------------------------------------------------

FORCED_MATCH_RULES = {

"6E1127":{
"DEP":"6E1128",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"6E1129":{
"DEP":"6E1130",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"6E1131":{
"DEP":"6E1134",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"6E1133":{
"DEP":"6E1132",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"UL101":{
"DEP":"UL102",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"UL103":{
"DEP":"UL104",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"QR674":{
"DEP":"QR675",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"QR676":{
"DEP":"QR677",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"QR670":{
"DEP":"QR671",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"EK652":{
"DEP":"EK652D",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"EK653":{
"DEP":"EK653D",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"FZ1207":{
"DEP":"FZ1208",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"FZ1026":{
"DEP":"FZ1026D",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"FZ1353":{
"DEP":"FZ1354",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"FZ1560":{
"DEP":"FZ15670",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
},

"FZ1025":{
"DEP":"FZ1025D",
"STA_WINDOW":("04:45","15:45"),
"STD_WINDOW":("09:00","23:59")
}

}


# ----------------------------------------------------------
# UPLOAD FILE
# ----------------------------------------------------------

print("Upload RAW Connecting Flight File")

upload = files.upload()
raw_file = list(upload.keys())[0]

df = pd.read_excel(raw_file, dtype=str)
df.columns = df.columns.str.strip()


# ----------------------------------------------------------
# CLEAN DATA
# ----------------------------------------------------------

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
df["Number"] = df["Flight ID"].str.extract(r"(\d+[A-Za-z]?)")


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
# TIME WINDOWS
# ----------------------------------------------------------

ARR_START = pd.to_datetime(CONFIG["ARR_START"]).time()
ARR_END = pd.to_datetime(CONFIG["ARR_END"]).time()

DEP_START = pd.to_datetime(CONFIG["DEP_START"]).time()
DEP_END = pd.to_datetime(CONFIG["DEP_END"]).time()


# ----------------------------------------------------------
# FLIGHT ID BUILDER
# ----------------------------------------------------------

def build_flight_id(arr_id, dep_id):

    arr_prefix=''.join(filter(str.isalpha,arr_id))
    arr_num=arr_id.replace(arr_prefix,"")
    dep_num=dep_id.replace(arr_prefix,"")

    if dep_id.startswith(arr_id):
        suffix=dep_id[len(arr_id):]
        return f"{arr_id}-{suffix}"

    if arr_num.isdigit() and dep_num.isdigit():

        max_len=max(len(arr_num),len(dep_num))
        arr_num_p=arr_num.zfill(max_len)
        dep_num_p=dep_num.zfill(max_len)

        diff_index=None

        for i in range(max_len):
            if arr_num_p[i]!=dep_num_p[i]:
                diff_index=i
                break

        if diff_index is None:
            return arr_id

        changed_part=dep_num_p[diff_index:]

        return f"{arr_id}-{changed_part}"

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

    # ------------------------------------------------------
    # FORCED MATCH LOGIC
    # ------------------------------------------------------

    for _, arr in arrivals.iterrows():

        arr_id = arr["Flight ID"]

        if arr_id in FORCED_MATCH_RULES:

            rule = FORCED_MATCH_RULES[arr_id]

            dep_id = rule["DEP"]

            sta_time = arr["Scheduled Time"].time()

            sta_ok = (
                pd.to_datetime(rule["STA_WINDOW"][0]).time()
                <= sta_time <=
                pd.to_datetime(rule["STA_WINDOW"][1]).time()
            )

            sta_display = arr["Scheduled Time"].strftime("%H:%M") if sta_ok else ""

            std_display = ""

            dep_match = departures[departures["Flight ID"] == dep_id]

            if not dep_match.empty:

                dep = dep_match.iloc[0]
                dep_time = dep["Scheduled Time"].time()

                std_ok = (
                    pd.to_datetime(rule["STD_WINDOW"][0]).time()
                    <= dep_time <=
                    pd.to_datetime(rule["STD_WINDOW"][1]).time()
                )

                if std_ok:
                    std_display = dep["Scheduled Time"].strftime("%H:%M")

            output_rows.append({
                "AIRLINE": AIRLINE_MASTER.get(prefix,prefix),
                "DAYS OF OPS": day,
                "FLT NO": build_flight_id(arr_id,dep_id),
                "STA": sta_display,
                "STD": std_display,
                "EFFECTIVE": arr["Start Date"].strftime("%d.%m.%y")
                + " - " +
                arr["End Date"].strftime("%d.%m.%y")
            })

            used_flights.add((arr_id,day,arr["Start Date"],arr["End Date"]))
            used_flights.add((dep_id,day,arr["Start Date"],arr["End Date"]))

    # ------------------------------------------------------
    # GENERIC MATCHING
    # ------------------------------------------------------

    for _, arr in arrivals.iterrows():

        if (arr["Flight ID"],day,arr["Start Date"],arr["End Date"]) in used_flights:
            continue

        for _, dep in departures.iterrows():

            if (dep["Flight ID"],day,dep["Start Date"],dep["End Date"]) in used_flights:
                continue

            if dep["Scheduled Time"] <= arr["Scheduled Time"]:
                continue

            if arr["Start Date"] != dep["Start Date"]:
                continue

            if arr["End Date"] != dep["End Date"]:
                continue

            output_rows.append({

                "AIRLINE": AIRLINE_MASTER.get(prefix,prefix),
                "DAYS OF OPS": day,
                "FLT NO": build_flight_id(arr["Flight ID"],dep["Flight ID"]),
                "STA": arr["Scheduled Time"].strftime("%H:%M"),
                "STD": dep["Scheduled Time"].strftime("%H:%M"),
                "EFFECTIVE": arr["Start Date"].strftime("%d.%m.%y")
                + " - " +
                arr["End Date"].strftime("%d.%m.%y")

            })

            used_flights.add((arr["Flight ID"],day,arr["Start Date"],arr["End Date"]))
            used_flights.add((dep["Flight ID"],day,dep["Start Date"],dep["End Date"]))
            break


# ----------------------------------------------------------
# UNMATCHED ARRIVALS
# ----------------------------------------------------------

    for _, arr in arrivals.iterrows():

        key = (arr["Flight ID"],day)

        if key in used_flights:
            continue

        output_rows.append({
            "AIRLINE": AIRLINE_MASTER.get(prefix,prefix),
            "DAYS OF OPS": day,
            "FLT NO": arr["Flight ID"],
            "STA": arr["Scheduled Time"].strftime("%H:%M"),
            "STD": "",
            "EFFECTIVE": arr["Start Date"].strftime("%d.%m.%y")
            + " - " +
            arr["End Date"].strftime("%d.%m.%y")
        })


# ----------------------------------------------------------
# UNMATCHED DEPARTURES
# ----------------------------------------------------------

    for _, dep in departures.iterrows():

        key = (dep["Flight ID"],day)

        if key in used_flights:
            continue

        output_rows.append({
            "AIRLINE": AIRLINE_MASTER.get(prefix,prefix),
            "DAYS OF OPS": day,
            "FLT NO": dep["Flight ID"],
            "STA": "",
            "STD": dep["Scheduled Time"].strftime("%H:%M"),
            "EFFECTIVE": dep["Start Date"].strftime("%d.%m.%y")
            + " - " +
            dep["End Date"].strftime("%d.%m.%y")
        })


# ----------------------------------------------------------
# FINAL DATAFRAME
# ----------------------------------------------------------

ramis_final = pd.DataFrame(output_rows)

# ----------------------------------------------------------
# REMOVE ARR/DEP ROWS WHEN ROTATION EXISTS
# ----------------------------------------------------------

pairs = ramis_final[ramis_final["FLT NO"].str.contains("-", na=False)]

clean_rows = []

for _, row in ramis_final.iterrows():

    if "-" in str(row["FLT NO"]):
        clean_rows.append(row)
        continue

    airline = row["AIRLINE"]
    day = row["DAYS OF OPS"]
    sta = row["STA"]
    std = row["STD"]
    eff = row["EFFECTIVE"]

    match = pairs[
        (pairs["AIRLINE"] == airline) &
        (pairs["DAYS OF OPS"] == day) &
        (pairs["EFFECTIVE"] == eff) &
        ((pairs["STA"] == sta) | (pairs["STD"] == std))
    ]

    if match.empty:
        clean_rows.append(row)

ramis_final = pd.DataFrame(clean_rows)

# ----------------------------------------------------------
# REMOVE DUPLICATE ROTATIONS
# ----------------------------------------------------------

ramis_final = ramis_final.drop_duplicates(
    subset=["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]
).reset_index(drop=True)

# ----------------------------------------------------------
# MERGE CONTINUOUS DATE RANGES
# ----------------------------------------------------------

ramis_final["START_DATE"] = ramis_final["EFFECTIVE"].str.split(" - ").str[0]
ramis_final["END_DATE"] = ramis_final["EFFECTIVE"].str.split(" - ").str[1]

ramis_final["START_DATE"] = pd.to_datetime(ramis_final["START_DATE"], format="%d.%m.%y")
ramis_final["END_DATE"] = pd.to_datetime(ramis_final["END_DATE"], format="%d.%m.%y")

ramis_final = ramis_final.sort_values(
    ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","START_DATE"]
)

merged_rows = []

for key, group in ramis_final.groupby(
    ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD"]
):

    group = group.sort_values("START_DATE")

    current_start = None
    current_end = None

    for _, row in group.iterrows():

        if current_start is None:
            current_start = row["START_DATE"]
            current_end = row["END_DATE"]
            continue

        if row["START_DATE"] == current_end + pd.Timedelta(days=1):
            current_end = row["END_DATE"]
        else:
            merged_rows.append((*key,current_start,current_end))
            current_start = row["START_DATE"]
            current_end = row["END_DATE"]

    merged_rows.append((*key,current_start,current_end))

merged_df = pd.DataFrame(
    merged_rows,
    columns=["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","START_DATE","END_DATE"]
)

merged_df["EFFECTIVE"] = (
    merged_df["START_DATE"].dt.strftime("%d.%m.%y")
    + " - " +
    merged_df["END_DATE"].dt.strftime("%d.%m.%y")
)

ramis_final = merged_df[["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]]

ramis_final = merged_df.copy()

ramis_final = ramis_final.sort_values(
    ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","START_DATE"],
    ascending=[True,True,True,True,True,False]
)

ramis_final = ramis_final.drop_duplicates(
    subset=["AIRLINE","DAYS OF OPS","FLT NO","STA","STD"],
    keep="first"
)

ramis_final["EFFECTIVE"] = (
    ramis_final["START_DATE"].dt.strftime("%d.%m.%y")
    + " - " +
    ramis_final["END_DATE"].dt.strftime("%d.%m.%y")
)

ramis_final = ramis_final[["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]]

# ----------------------------------------------------------
# DEBUG: SHOW UNPAIRED FLIGHTS
# ----------------------------------------------------------

missing_arrivals = []
missing_departures = []

for (prefix, day), group in df.groupby(["Prefix","Scheduled Day"]):

    arrivals = group[group["Type"]=="ARRIVAL"]
    departures = group[group["Type"]=="DEPARTURE"]

    for _, arr in arrivals.iterrows():

        key = (arr["Flight ID"], day)

        if key not in used_flights:
            missing_arrivals.append(
    f'{arr["Flight ID"]} {day} {arr["Start Date"].date()}-{arr["End Date"].date()}'
)

    for _, dep in departures.iterrows():

        key = (dep["Flight ID"], day)

        if key not in used_flights:
            missing_departures.append(
    f'{dep["Flight ID"]} {day} {dep["Start Date"].date()}-{dep["End Date"].date()}'
)

print("---- UNMATCHED ARRIVALS ----")
print(sorted(set(missing_arrivals)))

print("---- UNMATCHED DEPARTURES ----")
print(sorted(set(missing_departures)))

# ----------------------------------------------------------
# REMOVE FLIGHTS EXPIRED MORE THAN LOOKBACK DAYS
# ----------------------------------------------------------

today = pd.Timestamp.today().normalize()

cutoff_date = today - pd.Timedelta(days=CONFIG["LOOKBACK_DAYS"])

ramis_final["END_DATE_CHECK"] = (
    ramis_final["EFFECTIVE"].str.split(" - ").str[1]
)

ramis_final["END_DATE_CHECK"] = pd.to_datetime(
    ramis_final["END_DATE_CHECK"],
    format="%d.%m.%y",
    errors="coerce"
)

ramis_final = ramis_final[
    ramis_final["END_DATE_CHECK"] >= cutoff_date
].copy()

ramis_final.drop(columns=["END_DATE_CHECK"], inplace=True)



# ----------------------------------------------------------
# BUILD OUTPUT FILE
# ----------------------------------------------------------

wb = Workbook()
ws = wb.active
ws.title = "RAMIS CLEAN SHEET"

headers = ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]

weekday_order = [
"MONDAY","TUESDAY","WEDNESDAY",
"THURSDAY","FRIDAY","SATURDAY","SUNDAY"
]

row = 1

for day in weekday_order:

    ws.cell(row=row,column=1).value = day
    ws.cell(row=row,column=1).font = Font(bold=True)

    row += 2

    for col,header in enumerate(headers,start=1):

        ws.cell(row=row,column=col).value = header
        ws.cell(row=row,column=col).font = Font(bold=True)

    row += 1

    day_df = ramis_final[ramis_final["DAYS OF OPS"]==day]

    for _,r in day_df.iterrows():

        values=[r["AIRLINE"],r["DAYS OF OPS"],r["FLT NO"],r["STA"],r["STD"],r["EFFECTIVE"]]

        for c,v in enumerate(values,start=1):
            ws.cell(row=row,column=c).value = v

        row += 1

    row += 2


# ----------------------------------------------------------
# SAVE FILE
# ----------------------------------------------------------

output_file = "RAMIS_ROTATION_FINAL.xlsx"

from io import BytesIO

output = BytesIO()
macl_wb.save(output)
output.seek(0)

return output
