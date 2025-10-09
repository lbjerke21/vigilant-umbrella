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

    # Numbers from User details!B9+ (phones) and Call flow!D17:D27 (pilot numbers)
    numbers = []
    for cell in user_details_ws["B9":"B100"]:  # fixed to B9
        for c in cell:
            if c.value:
                numbers.append(str(c.value).strip())
    for cell in call_flow_ws["D17":"D27"]:
        for c in cell:
            if c.value:
                numbers.append(str(c.value).strip())
    # de-dupe while preserving order
    numbers = [n for n in dict.fromkeys(numbers) if n]

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
        "",      # limit concurrent calls now blank
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
    SEATS_COLS = 28
    def pad27(values): return (values + [""] * max(0, SEATS_COLS - len(values)))[:SEATS_COLS]

    # ---- Subscribers ----
    sub_rows = []
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#"]))
    sub_rows.append(pad27(["#BG Subscriber"]))
    sub_rows.append(pad27(["Subscriber"]))
    sub_rows.append(pad27([
        "MetaSphere CFS",
        "MetaSphere EAS",
        "Phone number",
        "Template",
        "Business Group (CFS)",
        "Business Group (EAS)",
        "CFS Subscriber Group",
        "Name (CFS)",
        "Name (EAS)",
        "PIN (CFS)",
        "PIN (EAS)",
        "EAS Preferred Language",
        "EAS Customer Group",
        "EAS Password",
        "Business Group Administration - account type (CFS)",
        "Business Group Administration - account type (EAS)",
        "Line State Monitoring - Subscribed",
        "Calling Name Delivery - local name (BG subscriber)",
        "Account Email",
        "Timezone (CFS)",
        "Timezone (EAS)",
        "Calling party number",
        "Charge number",
        "Calling party number for emergency calls",
        "Department (CFS)",
        "Department (EAS)",
        "Calling Name Delivery - use local name for intra-BG calls only",
    ]))

    for i in range(START_ROW, len(user_df)):
        name = user_df.iloc[i, 0]
        phone = user_df.iloc[i, 1]
        calling = user_df.iloc[i, 3]
        email = user_df.iloc[i, 5]
        account_type = user_df.iloc[i, 7]
        department = user_df.iloc[i, 8]
        template_raw = user_df.iloc[i, 11]

        # Skip blank phone or reserved/none templates
        if pd.isna(phone) or str(template_raw).strip() in ["None", "Reserve Number", "None | Reserve Number"]:
            continue

        template = convert_template(template_raw, region)
        is_aa = template in [f"{region}_AA_Easy", f"{region}_AA_Premium"]

        line_state_monitor = "" if is_aa else "TRUE"
        calling_name_delivery = "" if is_aa else ("" if pd.isna(name) else str(name))
        intra_bg_calls = "" if is_aa else "TRUE"
        acct_value = "Administrator" if str(account_type) in ["Location Admin", "Company Admin"] else "Normal"

        sub_rows.append(pad27([
            "CommandLink",
            "CommandLink_vEAS_LV",
            str(phone),
            template,
            customer_name,
            customer_name,
            "Standard Subscribers",
            "" if pd.isna(name) else str(name),
            "" if pd.isna(name) else str(name),
            "",
            "",
            "eng",
            "defaultGroup",
            "",
            acct_value,
            acct_value,
            line_state_monitor,
            calling_name_delivery,
            "" if pd.isna(email) else str(email),
            timezone,
            timezone,
            "" if pd.isna(calling) else str(calling),
            str(phone),
            str(phone),
            "" if pd.isna(department) else str(department),
            "" if pd.isna(department) else str(department),
            intra_bg_calls,
        ]))

    # Spacer after Subscribers
    sub_rows.append(pad27([""]))

    # ---- Managed Device ----
    sub_rows.append(pad27(["#Managed Device"]))
    sub_rows.append(pad27(["Managed Device"]))
    sub_rows.append(pad27([
        "MetaSphere CFS",
        "Business Group",
        "MAC address",
        "Assigned to user",
        "User directory number",
        "MAC trusted until",
        "Device version",
        "Device model",
        "Description",
    ]))

    for i in range(START_ROW, len(user_df)):
        phone = user_df.iloc[i, 1]
        mac = user_df.iloc[i, 12]
        if pd.isna(phone) or pd.isna(mac) or str(mac).strip() == "":
            continue
        sub_rows.append(pad27([
            "CommandLink",
            customer_name,
            str(mac),
            "TRUE",
            str(phone),
            mac_trusted_until_str(),
            "2",
            "Determined by Endpoint Pack",
            "",
        ]))

    # Spacer
    sub_rows.append(pad27([""]))
    sub_rows.append(pad27([""]))
    sub_rows.append(pad27([""]))

    # ---- Intercom Code Range ----
    sub_rows.append(pad27(["#Intercom Code Range"]))
    sub_rows.append(pad27(["Intercom Code Range"]))
    sub_rows.append(pad27([
        "MetaSphere CFS",
        "MetaSphere EAS",
        "Business Group",
        "First Code",
        "Last Code",
        "First Directory Number",
    ]))

    for i in range(START_ROW, len(user_df)):
        phone = user_df.iloc[i, 1]
        ext = user_df.iloc[i, 4]
        if pd.isna(phone) or pd.isna(ext) or str(ext).strip() == "":
            continue
        sub_rows.append(pad27([
            "CommandLink",
            "CommandLink_vEAS_LV",
            customer_name,
            str(ext),
            str(ext),
            str(phone),
        ]))

    # Spacer lines
    sub_rows.append(pad27([""]))
    sub_rows.append(pad27([""]))
    sub_rows.append(pad27([""]))
    sub_rows.append(pad27([""]))

    # ---- MLHGs ---- (start at row 17, Hunt on no-answer=FALSE, normalize Ring All)
    sub_rows.append(pad27(["#MLHGs"]))
    sub_rows.append(pad27(["MLHG"]))
    sub_rows.append(pad27([
        "MetaSphere CFS",
        "Business Group",
        "MLHG Name",
        "Members;Directory number;Login/logout supported",
        "Distribution algorithm",
        "Hunt on no-answer",
    ]))

    for r in range(17, 28):  # Excel rows 17..27 (skip header on row 16)
        mlg_name = call_flow_ws[f"B{r}"].value
        if not mlg_name:
            continue
        dist_alg = call_flow_ws[f"C{r}"].value
        dist_alg_clean = "" if pd.isna(dist_alg) else str(dist_alg).strip()
        if dist_alg_clean == "Ring All":
            dist_alg_clean = "Ring all"

        members = []
        for i in range(START_ROW, len(user_df)):
            if str(user_df.iloc[i, 13]).strip() == str(mlg_name).strip():
                num = user_df.iloc[i, 1]
                if pd.notna(num):
                    members.append(f"{{'{str(num)}';'FALSE'}}")

        sub_rows.append(pad27([
            "CommandLink",
            customer_name,
            str(mlg_name),
            ";".join(members),
            dist_alg_clean,
            "FALSE",  # per example file
        ]))

    # Spacer lines
    sub_rows.append(pad27([""]))
    sub_rows.append(pad27([""]))

    # ---- MLHG Pilot ---- (start at row 17 to skip header row)
    sub_rows.append(pad27(["#MLHG Pilot"]))
    sub_rows.append(pad27(["MLHG Pilot Number"]))
    sub_rows.append(pad27([
        "MetaSphere CFS",
        "MetaSphere EAS",
        "Business Group (CFS)",
        "MLHG Name",
        "Phone number",
        "Template",
        "Name (EAS)",
        "Name (CFS)",
        "PIN (EAS)",
        "EAS Password",
        "EAS Customer Group",
    ]))

    for r in range(17, 28):
        mlg_name = call_flow_ws[f"B{r}"].value
        phone_number = call_flow_ws[f"D{r}"].value
        pilot_vm = call_flow_ws[f"H{r}"].value
        if not mlg_name or not phone_number:
            continue

        pilot_template = (
            f"{region}_MLHG_Pilot"
            if str(pilot_vm).strip().lower() == "yes"
            else f"{region}_MLHG_Pilot_NoVM"
        )

        sub_rows.append(pad27([
            "CommandLink",
            "CommandLink_vEAS_LV",
            customer_name,
            str(mlg_name),
            str(phone_number),
            pilot_template,
            f"{mlg_name} Pilot",
            f"{mlg_name} Pilot",
            "*",
            "*",
            "defaultGroup",
        ]))

    # Write Seats CSV to memory
    seats_buffer = io.StringIO()
    csv.writer(seats_buffer, lineterminator="\n").writerows(sub_rows)
    seats_filename = f"Seats-Devices-Exts-MLHG-{customer_name}.csv"

    # ---------- Downloads ----------
    st.download_button(
        label=f"⬇️ Download {bg_filename}",
        data=bg_buffer.getvalue(),
        file_name=bg_filename,
        mime="text/csv",
    )
    st.download_button(
        label=f"⬇️ Download {seats_filename}",
        data=seats_buffer.getvalue(),
        file_name=seats_filename,
        mime="text/csv",
    )

else:
    st.info("Please upload an Excel file to begin.")




