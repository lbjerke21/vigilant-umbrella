import streamlit as st
import pandas as pd
import io
import csv
from datetime import datetime, timedelta
from openpyxl import load_workbook

st.set_page_config(page_title="UCaaS CSV Generator", page_icon="📄", layout="centered")
st.title("📄 UCaaS CSV Generator")
st.write("Upload your UCaaS Excel file and generate the two formatted CSVs automatically.")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])


# ---------- Helpers ----------

def _norm(name: str) -> str:
    """Normalize a sheet name for case/space-insensitive matching."""
    return "".join(name.lower().split())

def get_sheet(wb, wanted: str):
    """Return worksheet even if the name capitalization or spacing differs."""
    wanted_norm = _norm(wanted)
    for name in wb.sheetnames:
        if _norm(name) == wanted_norm:
            return wb[name]
    raise KeyError(f"Worksheet '{wanted}' not found. Available: {wb.sheetnames}")

def convert_template(template_name, region):
    """Convert UCaaS template names based on mapping rules."""
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
    if s in mapping:
        return f"{region}{mapping[s]}"
    return ""

def mac_trusted_until_str():
    """m/d/YYYY  11:59:59 PM (two spaces before time), 4 weeks out."""
    dt = datetime.now() + timedelta(weeks=4)
    return f"{dt.month}/{dt.day}/{dt.year}  11:59:59 PM"


# ---------- Main ----------

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    try:
        user_details_ws = get_sheet(wb, "User details")
        cmdlink_ws      = get_sheet(wb, "CommandLink")
        call_flow_ws    = get_sheet(wb, "Call flow")
    except KeyError as e:
        st.error(str(e))
        st.stop()

    st.caption(f"✅ Found sheets: {wb.sheetnames}")

    # Key metadata
    customer_name = user_details_ws["B3"].value
    region = (cmdlink_ws["C4"].value or "").strip()  # CH or LV
    timezone = cmdlink_ws["C5"].value

    st.success(f"Loaded file for **{customer_name}** (Region: {region}, Timezone: {timezone})")

    # Read the User details sheet as a DataFrame for row-wise parsing (no header row)
    user_df = pd.read_excel(uploaded_file, sheet_name=user_details_ws.title, header=None)

    # Convenience indices (0-based) for columns on User details:
    COL_NAME = 0       # A
    COL_PHONE = 1      # B
    COL_CALLING = 3    # D (Calling party number)
    COL_EXT = 4        # E (Intercom code)
    COL_EMAIL = 5      # F
    COL_ACCT_TYPE = 7  # H
    COL_DEPT = 8       # I
    COL_TEMPLATE = 11  # L
    COL_MAC = 12       # M
    COL_MLHG = 13      # N

    START_ROW = 8  # Excel row 9

    # =========================
    # Build BG CSV (width=12) |
    # =========================
    BG_COLS = 12
    def pad12(values): return (values + [""] * max(0, BG_COLS - len(values)))[:BG_COLS]
    bg_template = f"{region} BG"

    # Numbers from User details!D9+ and Call flow!D17:D27
    numbers = []
    for cell in user_details_ws["D9":"D100"]:
        for c in cell:
            if c.value:
                numbers.append(str(c.value).strip())
    for cell in call_flow_ws["D17":"D27"]:
        for c in cell:
            if c.value:
                numbers.append(str(c.value).strip())

    # Unique departments from User details!I9+
    departments = []
    seen_depts = set()
    for cell in user_details_ws["I9":"I100"]:
        for c in cell:
            if c.value:
                d = str(c.value).strip()
                if d and d not in seen_depts:
                    seen_depts.add(d)
                    departments.append(d)

    bg_rows = []
    # Top comment/header lines
    bg_rows.append(pad12(["#"]))
    bg_rows.append(pad12(["#"]))
    bg_rows.append(pad12(["#"]))
    # Business Groups
    bg_rows.append(pad12(["#Business Groups"]))
    bg_rows.append(pad12(["Business Group"]))
    bg_rows.append(pad12([
        "MetaSphere CFS",
        "MetaSphere EAS",
        "Business Group",
        "Template",
        "CFS Persistent Profile",
        "Local CNAM name",
        "Music On Hold Service - Subscribed",
        "Music On Hold Service - class of service",
        "Music On Hold Service - limit concurrent calls",
        "Music On Hold Service - maximum concurrent calls",
        "Music On Hold Service - Service Level",
        "Music On Hold Service - Application Server",
    ]))
    bg_rows.append(pad12([
        "CommandLink",
        "CommandLink_vEAS_LV",
        customer_name,
        bg_template,
        bg_template,
        "",
        "TRUE",
        "0",
        "TRUE",
        "16",
        "Enhanced",
        "EAS Voicemail",
    ]))
    # Spacers
    bg_rows.append(pad12([""]))
    bg_rows.append(pad12([""]))
    # Number Blocks
    bg_rows.append(pad12(["#BG Number Blocks"]))
    bg_rows.append(pad12(["Business Group Number Block"]))
    bg_rows.append(pad12([
        "MetaSphere CFS",
        "Business Group",
        "First Phone Number",
        "Block size",
        "CFS Subscriber Group",
    ]))
    for num in numbers:
        bg_rows.append(pad12([
            "CommandLink",
            customer_name,
            num,
            "1",
            "Standard Subscribers",
        ]))
    # Spacers
    bg_rows.append(pad12([""]))
    bg_rows.append(pad12([""]))
    bg_rows.append(pad12([""]))
    # Department
    bg_rows.append(pad12(["Department"]))
    bg_rows.append(pad12([
        "MetaSphere CFS",
        "MetaSphere EAS",
        "Business Group",
        "Name",
    ]))
    for dept in departments:
        bg_rows.append(pad12([
            "CommandLink",
            "CommandLink_vEAS_LV",
            customer_name,
            dept,
        ]))

    # Write BG CSV to memory
    bg_buffer = io.StringIO()
    csv.writer(bg_buffer, lineterminator="\n").writerows(bg_rows)
    bg_filename = f"BG-NumberBlock-Departments-{customer_name}.csv"


    # =================================
    # Build Seats/Devices/Exts/MLHG CSV
    # with exact sections (width=27)
    # =================================
    SEATS_COLS = 27
    def pad27(values): return (values + [""] * max(0, SEATS_COLS - len(values)))[:SEATS_COLS]

    # ---- Subscribers ----
    sub_rows = []
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#BG Subscriber"]))
    sub_rows._
