import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
from openpyxl import load_workbook

st.set_page_config(page_title="UCaaS CSV Generator", page_icon="📄", layout="centered")

st.title("📄 UCaaS CSV Generator")
st.write("Upload your UCaaS Excel file and generate the two formatted CSVs automatically.")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])


# --- Helper functions ---

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
        "UCaaS|Link Complete ContactCenter Manager": "_Complete"
    }
    for key, suffix in mapping.items():
        if template_name.strip() == key:
            return f"{region}{suffix}"
    return ""


# --- Main app logic ---

if uploaded_file:
    # Read workbook and normalize sheet names
    wb = load_workbook(uploaded_file, data_only=True)
    try:
        user_details = get_sheet(wb, "User details")
        cmdlink      = get_sheet(wb, "CommandLink")
        call_flow    = get_sheet(wb, "Call flow")
    except KeyError as e:
        st.error(str(e))
        st.stop()

    st.caption(f"✅ Found sheets: {wb.sheetnames}")

    # Extract key metadata
    customer_name = user_details["B3"].value
    region = cmdlink["C4"].value  # CH or LV
    timezone = cmdlink["C5"].value

    st.success(f"Loaded file for **{customer_name}** (Region: {region}, Timezone: {timezone})")

    # --- BG CSV ---
    bg_rows = []
    business_group = customer_name
    bg_template = f"{region} BG"
    numbers = []

    # Collect numbers from User details!D9+ and Call flow!D17:D27
    for cell in user_details["D9":"D100"]:
        for c in cell:
            if c.value:
                numbers.append(c.value)
    for cell in call_flow["D17":"D27"]:
        for c in cell:
            if c.value:
                numbers.append(c.value)

    departments = set()
    for cell in user_details["I9":"I100"]:
        for c in cell:
            if c.value:
                departments.add(c.value)

    for num in numbers:
        bg_rows.append({
            "MetaSphere CFS": "CommandLink",
            "MetaSphere EAS": "CommandLink_vEAS_LV",
            "Business Group": business_group,
            "Template": bg_template,
            "CFS Persistent Profile": bg_template,
            "Local CNAM name": "",
            "Music On Hold Service - Subscribed": "TRUE",
            "Music On Hold Service - class of service": 0,
            "Music On Hold Service - limit concurrent calls": "TRUE",
            "Music On Hold Service - maximum concurrent calls": 16,
            "Music On Hold Service - Service Level": "Enhanced",
            "Music On Hold Service - Application Server": "EAS Voicemail",
            "First Phone number": num,
            "Block size": 1,
            "CFS Subscriber Group": "Standard Subscribers"
        })

    # Add departments
    for dept in departments:
        bg_rows.append({"Department Name": dept})

    bg_df = pd.DataFrame(bg_rows)

    # --- Seats CSV ---
    seat_rows = []
    managed_device_rows = []
    intercom_rows = []
    mlg_rows = []
    mlg_pilot_rows = []

    df = pd.read_excel(uploaded_file, sheet_name=get_sheet(wb, "User details").title, header=None)
    start_row = 8  # Row 9 in Excel
    for i in range(start_row, len(df)):
        name = df.iloc[i, 0]
        phone = df.iloc[i, 1]
        ext = df.iloc[i, 4]
        calling = df.iloc[i, 3]
        email = df.iloc[i, 5]
        account_type = df.iloc[i, 7]
        department = df.iloc[i, 8]
        template_raw = df.iloc[i, 11]
        mac = df.iloc[i, 12]

        if pd.isna(phone) or str(template_raw).strip() in ["None", "Reserve Number", "None | Reserve Number"]:
            continue

        template = convert_template(str(template_raw), region)
        line_state_monitor = "" if template in [f"{region}_AA_Easy", f"{region}_AA_Premium"] else "TRUE"
        calling_name_delivery = "" if template in [f"{region}_AA_Easy", f"{region}_AA_Premium"] else name
        intra_bg_calls = "" if template in [f"{region}_AA_Easy", f"{region}_AA_Premium"] else "TRUE"
        account_type_value = "Administrator" if str(account_type) in ["Location Admin", "Company Admin"] else "Normal"

        seat_rows.append({
            "MetaSphere CFS": "CommandLink",
            "MetaSphere EAS": "CommandLink_vEAS_LV",
            "Phone Number": phone,
            "Template": template,
            "Business Group (CFS)": customer_name,
            "Business Group (EAS)": customer_name,
            "CFS Subscriber Group": "Standard Subscribers",
            "Name (CFS)": name,
            "Name (EAS)": name,
            "PIN (CFS)": "",
            "PIN (EAS)": "",
            "EAS Preferred Language": "eng",
            "EAS Password": "",
            "Business Group Administration - account type (CFS)": account_type_value,
            "Business Group Administration - account type (EAS)": account_type_value,
            "Line State Monitoring - Subscribed": line_state_monitor,
            "Calling Name Delivery - local name (BG subscriber)": calling_name_delivery,
            "Account Email": email,
            "Timezone (CFS)": timezone,
            "Timezone (EAS)": timezone,
            "Calling party number": calling,
            "Charge number": phone,
            "Calling party number for emergency calls": phone,
            "Department (CFS)": department,
            "Department (EAS)": department,
            "Calling Name Delivery - use local name for intra-BG calls only": intra_bg_calls
        })

        managed_device_rows.append({
            "MAC address": mac,
            "Assigned to user": "TRUE",
            "User directory number": phone,
            "MAC trusted until": (datetime.now() + timedelta(weeks=4)).strftime("%-m/%-d/%Y  11:59:59 PM"),
            "Device version": 2,
            "Device model": "Determined by Endpoint Pack",
            "Description": ""
        })

        intercom_rows.append({
            "First Code": ext,
            "Last Code": ext,
            "First Directory Number": phone
        })

    # --- MLHG Section ---
    for i in range(16, 27):
        mlg_name = call_flow[f"B{i}"].value
        if not mlg_name:
            continue
        members = []
        dist_alg = call_flow[f"C{i}"].value
        phone_number = call_flow[f"D{i}"].value
        pilot_vm = call_flow[f"H{i}"].value
        if not phone_number:
            continue
        mlg_users = []
        for j in range(start_row, len(df)):
            if str(df.iloc[j, 13]) == str(mlg_name):
                num = df.iloc[j, 1]
                if num:
                    mlg_users.append(f"{{'{num}';'FALSE'}}")
        member_field = ";".join(mlg_users)
        mlg_rows.append({
            "MLHG Name": mlg_name,
            "Members;Directory number;Login/logout supported": member_field,
            "Distribution algorithm": dist_alg,
            "Hunt on no-answer": "No"
        })
        pilot_template = f"{region}_MLHG_Pilot" if str(pilot_vm).strip().lower() == "yes" else f"{region}_MLHG_Pilot_NoVM"
        mlg_pilot_rows.append({
            "Phone number": phone_number,
            "Template": pilot_template,
            "Name (EAS)": f"{mlg_name} Pilot",
            "Name (CFS)": f"{mlg_name} Pilot",
            "PIN (EAS)": "*",
            "EAS Password": "*"
        })

    seats_df = pd.concat([
        pd.DataFrame(seat_rows),
        pd.DataFrame(managed_device_rows),
        pd.DataFrame(intercom_rows),
        pd.DataFrame(mlg_rows),
        pd.DataFrame(mlg_pilot_rows)
    ], axis=1)

    # Prepare CSVs in memory
    bg_buffer = io.StringIO()
    seat_buffer = io.StringIO()
    bg_df.to_csv(bg_buffer, index=False)
    seats_df.to_csv(seat_buffer, index=False)

    bg_filename = f"BG-NumberBlock-Departments-{customer_name}.csv"
    seat_filename = f"Seats-Devices-Exts-MLHG-{customer_name}.csv"

    st.download_button(
        label=f"⬇️ Download {bg_filename}",
        data=bg_buffer.getvalue(),
        file_name=bg_filename,
        mime="text/csv"
    )

    st.download_button(
        label=f"⬇️ Download {seat_filename}",
        data=seat_buffer.getvalue(),
        file_name=seat_filename,
        mime="text/csv"
    )

else:
    st.info("Please upload an Excel file to begin.")
