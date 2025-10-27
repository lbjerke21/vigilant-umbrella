import streamlit as st
import pandas as pd
import io
import csv
from datetime import datetime, timedelta
from openpyxl import load_workbook

st.set_page_config(page_title="CTI Sheet -> Meta Import Files", page_icon="📄", layout="centered")
st.title("📄 CTI Sheet -> Meta Import Files")
st.write("Upload your CTI Excel file and generate a single Meta Import CSV (Step1 + Step2).")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

# ---------- Helpers ----------

def _norm(name: str) -> str:
    return "".join(name.lower().split())

def get_sheet(wb, wanted: str):
    wanted_norm = _norm(wanted)
    for name in wb.sheetnames:
        if _norm(name) == wanted_norm:
            return wb[name]
    raise KeyError(f"Worksheet '{wanted}' not found. Available: {wb.sheetnames}")

def convert_template(template_name, region):
    if pd.isna(template_name):
        return ""
    mapping = {
        "UCaaS|Link Basic Auto-Attendant": "_AA_Easy",
        "UCaaS|Link Premium Auto-Attendant": "_AA_Premium",
        "UCaaS|Link Lite": "_Lite",
        "UCaaS|Link Standard": "_STD",
        "UCaaS|Link Complete": "_Complete",
        "UCaaS|Link Complete (HIPPA)": "Complete_HIPPA",
        "UCaaS|Link Complete (No Voicemail)": "_Complete_NoVM",
        "UCaaS|Link Complete ContactCenter Agent": "_Complete",
        "UCaaS|Link Complete ContactCenter Manager": "_Complete",
    }
    s = str(template_name).strip()
    return f"{region}{mapping[s]}" if s in mapping else ""

def mac_trusted_until_str():
    """4 weeks out at 11:59:59 pm, formatted m/d/yy h:mm:ss am"""
    dt = (datetime.now() + timedelta(weeks=4)).replace(hour=23, minute=59, second=59, microsecond=0)
    h12 = dt.hour % 12 or 12
    ampm = "am" if dt.hour < 12 else "pm"
    yy = dt.strftime("%y")
    return f"{dt.month}/{dt.day}/{yy} {h12}:{dt.minute:02d}:{dt.second:02d} {ampm}"

# ---------- Main ----------

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    try:
        user_details_ws = get_sheet(wb, "User details")
        eng_ws          = get_sheet(wb, "Engineering")
        call_flow_ws    = get_sheet(wb, "Call flow")
    except KeyError as e:
        st.error(str(e)); st.stop()

    st.caption(f"✅ Found sheets: {wb.sheetnames}")

    customer_name = user_details_ws["B3"].value
    region = (eng_ws["C4"].value or "").strip()  # CH or LV

    st.success(f"Loaded file for **{customer_name}** (Region: {region})")

    user_df = pd.read_excel(uploaded_file, sheet_name=user_details_ws.title, header=None)

    # Column indexes (0-based)
    COL_NAME = 0       # A
    COL_PHONE = 1      # B
    COL_CALLING = 3    # D
    COL_EXT = 4        # E
    COL_EMAIL = 5      # F
    COL_ACCT_TYPE = 7  # H
    COL_DEPT = 8       # I
    COL_TZ = 9         # J (per-user timezone)
    COL_TEMPLATE = 12  # M (shifted due to new J)
    COL_MAC = 13       # N
    COL_MLHG = 14      # O

    START_ROW = 8  # Excel row 9

    # =========================
    # Build BG (width=12)
    # =========================
    BG_COLS = 12
    def pad12(values): return (values + [""] * max(0, BG_COLS - len(values)))[:BG_COLS]
    bg_template = f"{region} BG"

    numbers = []
    for cell in user_details_ws["B9":"B100"]:
        for c in cell:
            if c.value:
                numbers.append(str(c.value).strip())
    for cell in call_flow_ws["D17":"D27"]:
        for c in cell:
            if c.value:
                numbers.append(str(c.value).strip())
    numbers = [n for n in dict.fromkeys(numbers) if n]

    departments, seen_depts = [], set()
    for cell in user_details_ws["I9":"I100"]:
        for c in cell:
            if c.value:
                d = str(c.value).strip()
                if d and d not in seen_depts:
                    seen_depts.add(d); departments.append(d)

    bg_rows = []
    bg_rows.append(pad12(["#"]))
    bg_rows.append(pad12(["#"]))
    bg_rows.append(pad12(["#"]))
    bg_rows.append(pad12(["#Business Groups"]))
    bg_rows.append(pad12(["Business Group"]))
    bg_rows.append(pad12([
        "MetaSphere CFS","MetaSphere EAS","Business Group","Template","CFS Persistent Profile",
        "Local CNAM name","Music On Hold Service - Subscribed","Music On Hold Service - class of service",
        "Music On Hold Service - limit concurrent calls","Music On Hold Service - maximum concurrent calls",
        "Music On Hold Service - Service Level","Music On Hold Service - Application Server",
    ]))
    bg_rows.append(pad12([
        "CommandLink","CommandLink_vEAS_LV",customer_name,bg_template,bg_template,"",
        "TRUE","0","", "16","Enhanced","EAS Voicemail",
    ]))
    bg_rows.append(pad12([""]))
    bg_rows.append(pad12([""]))
    bg_rows.append(pad12(["#BG Number Blocks"]))
    bg_rows.append(pad12(["Business Group Number Block"]))
    bg_rows.append(pad12([
        "MetaSphere CFS","Business Group","First Phone Number","Block size","CFS Subscriber Group",
    ]))
    for num in numbers:
        bg_rows.append(pad12(["CommandLink",customer_name,num,"1","Standard Subscribers"]))
    bg_rows.append(pad12([""]))
    bg_rows.append(pad12([""]))
    bg_rows.append(pad12([""]))
    bg_rows.append(pad12(["Department"]))
    bg_rows.append(pad12(["MetaSphere CFS","MetaSphere EAS","Business Group","Name"]))
    for dept in departments:
        bg_rows.append(pad12(["CommandLink","CommandLink_vEAS_LV",customer_name,dept]))

    # =========================
    # Build Seats/Devices/Exts/MLHG (width=28)
    # =========================
    SEATS_COLS = 28
    def pad27(values): return (values + [""] * max(0, SEATS_COLS - len(values)))[:SEATS_COLS]

    sub_rows = []
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#BG Subscriber"]))
    sub_rows.append(pad27(["Subscriber"]))
    sub_rows.append(pad27([
        "MetaSphere CFS","MetaSphere EAS","Phone number","Template","Business Group (CFS)","Business Group (EAS)",
        "CFS Subscriber Group","Name (CFS)","Name (EAS)","PIN (CFS)","PIN (EAS)","EAS Preferred Language",
        "EAS Customer Group","EAS Password","Business Group Administration - account type (CFS)",
        "Business Group Administration - account type (EAS)","Line State Monitoring - Subscribed",
        "Calling Name Delivery - local name (BG subscriber)","Account Email","Timezone (CFS)","Timezone (EAS)",
        "Calling party number","Charge number","Calling party number for emergency calls","Department (CFS)",
        "Department (EAS)","Calling Name Delivery - use local name for intra-BG calls only",
    ]))

    for i in range(START_ROW, len(user_df)):
        name        = user_df.iloc[i, COL_NAME]
        phone       = user_df.iloc[i, COL_PHONE]
        calling     = user_df.iloc[i, COL_CALLING]
        email       = user_df.iloc[i, COL_EMAIL]
        account_type= user_df.iloc[i, COL_ACCT_TYPE]
        department  = user_df.il_
