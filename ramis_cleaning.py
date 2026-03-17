import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO


def process(macl_df, ramis_df):

    # -----------------------------
    # CONFIG
    # -----------------------------
    CONFIG = {
        "SEASON_START": "2025-10-26",
        "SEASON_END": "2026-03-28",
        "ARR_START": "04:45:00",
        "ARR_END": "15:36:00",
        "DEP_START": "09:00:00",
        "DEP_END": "23:59:00",
        "LOOKBACK_DAYS": 1
    }

    # -----------------------------
    # AIRLINE MASTER
    # -----------------------------
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

    # -----------------------------
    # USE RAMIS INPUT
    # -----------------------------
    df = ramis_df.copy()
    df.columns = df.columns.str.strip()

    # -----------------------------
    # CLEAN DATA
    # -----------------------------
    exclude_list = ["ARRTBA","DEPTBA","RESARR","RESDEP","MLE","MLED","DMLE","DMLED"]

    df = df[~df["Flight ID"].str.upper().isin(exclude_list)].copy()

    df["Scheduled Day"] = df["Scheduled Day"].str.strip().str.upper()
    df["Type"] = df["Type"].str.strip().str.upper()

    df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
    df["End Date"] = pd.to_datetime(df["End Date"], errors="coerce")

    df["Scheduled Time"] = df["Scheduled Time"].astype(str).str.split(".").str[0]
    df["Scheduled Time"] = pd.to_datetime(df["Scheduled Time"], format="%H:%M:%S", errors="coerce")

    df["Prefix"] = df["Flight ID"].str[:2]

    # -----------------------------
    # MATCHING LOGIC (SIMPLIFIED CORE)
    # -----------------------------
    output_rows = []

    for (prefix, day), group in df.groupby(["Prefix", "Scheduled Day"]):

        arrivals = group[group["Type"]=="ARRIVAL"].copy()
        departures = group[group["Type"]=="DEPARTURE"].copy()

        arrivals = arrivals.sort_values("Scheduled Time")
        departures = departures.sort_values("Scheduled Time")

        for _, arr in arrivals.iterrows():

            matched = False

            for _, dep in departures.iterrows():

                if dep["Scheduled Time"] > arr["Scheduled Time"]:

                    output_rows.append({
                        "AIRLINE": AIRLINE_MASTER.get(prefix, prefix),
                        "DAYS OF OPS": day,
                        "FLT NO": f"{arr['Flight ID']}-{dep['Flight ID'][-2:]}",
                        "STA": arr["Scheduled Time"].strftime("%H:%M"),
                        "STD": dep["Scheduled Time"].strftime("%H:%M"),
                        "EFFECTIVE": arr["Start Date"].strftime("%d.%m.%y") + " - " + arr["End Date"].strftime("%d.%m.%y")
                    })

                    matched = True
                    break

            if not matched:
                output_rows.append({
                    "AIRLINE": AIRLINE_MASTER.get(prefix, prefix),
                    "DAYS OF OPS": day,
                    "FLT NO": arr["Flight ID"],
                    "STA": arr["Scheduled Time"].strftime("%H:%M"),
                    "STD": "",
                    "EFFECTIVE": arr["Start Date"].strftime("%d.%m.%y") + " - " + arr["End Date"].strftime("%d.%m.%y")
                })

    ramis_final = pd.DataFrame(output_rows)

    # -----------------------------
    # BUILD EXCEL
    # -----------------------------
    macl_wb = Workbook()
    ws = macl_wb.active
    ws.title = "RAMIS CLEAN SHEET"

    headers = ["AIRLINE","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE"]

    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header
        ws.cell(row=1, column=col).font = Font(bold=True)

    for i, row in ramis_final.iterrows():
        for j, val in enumerate(row, start=1):
            ws.cell(row=i+2, column=j).value = val

    # -----------------------------
    # RETURN FILE
    # -----------------------------
    output = BytesIO()
    macl_wb.save(output)
    output.seek(0)

    return output
