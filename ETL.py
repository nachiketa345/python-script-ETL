"""
HSEQ ETL Pipeline
=================
Extract  : Downloads MS Forms response Excel from SharePoint/OneDrive
Transform: Maps columns, routes by Category (Normal / Accident)
Load     : Outputs Maximo-ready CSV flat files

Requirements:
    pip install pandas openpyxl msal requests
"""

import os
import io
import re
import unicodedata
import pandas as pd
import requests
from msal import ConfidentialClientApplication
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIG — fill these in before running
# ─────────────────────────────────────────────
TENANT_ID = "YOUR_TENANT_ID"
CLIENT_ID = "YOUR_CLIENT_ID"
CLIENT_SECRET = "YOUR_CLIENT_SECRET"

SHAREPOINT_SITE = "yourorg.sharepoint.com"
SHAREPOINT_SITE_NAME = "YourSiteName"
FILE_PATH = "/Shared Documents/Forms/CVL Event Reporting.xlsx"

# Fixed values loaded into every row
FIXED_STATUS = "NEW"
FIXED_CLASS = "HSEQ"
FIXED_ORGID = "TFW"

OUTPUT_DIR = "./output"

LOCAL_TEST = True
LOCAL_INPUT_PATH = r"C:\Nachiketa\HSEQ RESPONSE EXCEL\CVL Event Reporting(1-13).xlsx"
LOCAL_OUTPUT_DIR = r"C:\Users\2152355\OneDrive - Cognizant\Desktop\OUTPUT ETL"

# ─────────────────────────────────────────────
# COLUMN MAPPINGS
# ─────────────────────────────────────────────

# KA_IRINCSOURCE_C — first non-empty wins from these 3 L2 columns
# Note: if all 3 are blank (e.g. user selected AIW IM but didn't fill L2)
# we fall back to 'Area you work within AIW IM' as an alternative
IRINCSOURCE_C_COLS = [
    "Who do you work for L2? (AIW IM)",
    "Who do you work for L2? (Transport\n  for Wales (TfW)",
    "Who do you work for L2? (Trans)",
    "Area you work within AIW IM",   # fallback if L2 cols are all blank
]

# KA_IRASSOCIATED_C — first non-empty wins from these 2
IRASSOCIATED_C_COLS = [
    "Event Associated with AIW Infrastructure Manager (IM)",
    # typo 'Assiciated' is in the form itself
    "Event Assiciated with AIW Transformation Project (Trans)",
]
CATEGORY_SUBCOL_MAP = {
    "CLOSE CALL": "Close Call Type",
    "INCIDENT": "Incident",
    "SERVICE STRIKE": "If, Service Strike, please select option below",
    "ASSAULT": "Assault Type",
    "OTHER": "Other",
    "ACCIDENT": "Nature of Accident (Select most serious):",
}

# Core normal mapping (excluding computed fields handled separately)
NORMAL_MAP = {
    "Email": "KA_FBEMAIL",
    "Raised by": "KA_RAISEDBY",
    "Who do you work for?": "KA_IRINCSOURCE",
    # KA_IRINCSOURCE_C  → computed from 3 cols (see get_irincsource_c)
    "Area you work within AIW IM": "KA_IRINCCONNECTED_C",
    "Date and Time of Event - Please input as follows (dd/mm/yyy 00:00)": "REPORTDATE",
    "Category": "KA_IRINCCATEGORY",
    # KA_IRINCCATEGORY_C → computed from Category (see get_irinccategory_c)
    "Was a subcontractor involved?": "KA_SUBCONYN",
    "Who is the Subcontractor?": "KA_SUBCONTRACTOR",
    "Event Associated with/caused by?": "KA_IRASSOCIATED",
    # KA_IRASSOCIATED_C → computed from 2 cols (see get_irassociated_c)
    "Location Type": "KA_IRLOCTYPE",
    "Side Of Line": "KA_SIDEOFLINE",
    "ELR": "KA_ELR",
    # KA_ELR, KA_MILES, KA_YARDS → all computed from ELR field (see parse_elr)
    "What 3 Words": "KA_IRWHAT3WORDS",
    "Latitude and Longtitude": "KA_LATITUDE",
    "Description of Event - Please DO NOT enter personal Information here": "KA_IRLONG_DESCRIPTION",
    "Immediate Action Taken": "KA_IRACTIONTAKEN",
    "What could have happened?": "KA_IRCONSEQUENCE",       # NEW
    "Did the immediate action resolve the issue?": "KA_ISSUERESOLVED",       # NEW
    "Do you require a receipt of this Event Record": "KA_REQFB",
}

# Accident-only extra mapping
ACCIDENT_EXTRA_MAP = {
    "Gender": "KA_RGENDER",
    "Age": "KA_RAGE",
    "Time into Shift (Hours)": "KA_RTIMEINTOSHIFT",
    "Time of Accident (00:00)": "KA_RTIMEOFACCIDENT",
    "Was the Accident reported immediately?": "KA_RARIMMEDIATELY",
    "Name of Injured person, if applicable": "KA_IPNAME",
    "Address of Injured person, if applicable (Enter Unknown if Unknown)": "KA_IPADDRESS",
    "Worker or Non-Worker": "KA_WORKER",
    "Length of time in Role (Number only)": "KA_RTIMEINROLE",
    "Role Measure": "KA_RTIMEINROLETYPE",
    "Any witnesses?": "KA_WITNESS",
    "Witness Name(s):": "KA_WITNESSNAME",
    "First aid given at scene?": "KA_FAGIVEN",
    "Details of First Aid Received": "KA_FADETAILS",
    "Injured Person taken directly to hospital?": "KA_IPHOSPITAL",
    "If Yes, provide details:": "KA_IPHOSPDETAIL",
    "Injury sustained": "KA_NATUREOFINJURY",
}

ACCIDENT_MAP = {**NORMAL_MAP, **ACCIDENT_EXTRA_MAP}

# Body part columns — names in the Excel ARE the Maximo attribute names
BODY_PART_COLS = [
    "KA_BPHEAD", "KA_BPCHEST", "KA_BPSHOULDERLEFT", "KA_BPEAR", "KA_BPABDOMEN",
    "KA_BPSHOULDERRIGHT", "KA_BPEYE", "KA_BPUPPERBACK", "KA_BPELBOWLEFT",
    "KA_BPFACIALBONES", "KA_BPLOWERBACK", "KA_BPELBOWRIGHT", "KA_BPNOSE",
    "KA_BPHIPLEFT", "KA_BPWRISTLEFT", "KA_BPMOUTH", "KA_BPHIPRIGHT", "KA_BPWRISTRIGHT",
    "KA_BPNECK", "KA_BPUPPERLEGLEFT", "KA_BPHANDRIGHT", "KA_BPHANDLEFT",
    "KA_BPUPPERLEGRIGHT", "KA_BPFINGER", "KA_BPKNEELEFT", "KA_BPKNEERIGHT",
    "KA_BPLOWERLEGLEFT", "KA_BPLOWERLEGRIGHT", "KA_BPANKLELEFT", "KA_BPANKLERIGHT",
    "KA_BPFOOTLEFT", "KA_BPFOOTRIGHT",
]

# Text fields that need encoding cleanup (Â, extra spaces, junk chars)
TEXT_CLEAN_FIELDS = {
    "KA_IRLONG_DESCRIPTION", "KA_IRACTIONTAKEN", "KA_IRCONSEQUENCE",
    "KA_IRINCSOURCE_C", "KA_IRINCCATEGORY_C", "KA_IRINCCONNECTED_C",
    "KA_SUBCONTRACTOR", "KA_IPNAME", "KA_IPADDRESS", "KA_WITNESSNAME",
    "KA_FADETAILS", "KA_IPHOSPDETAIL",
}

# Final ordered output columns
NORMAL_COLS = [
    "STATUS", "CLASS", "ORGID",
    "KA_FBEMAIL", "KA_RAISEDBY", "KA_IRINCSOURCE", "KA_IRINCSOURCE_C",
    "KA_IRINCCONNECTED_C", "REPORTDATE", "KA_IRINCCATEGORY", "KA_IRINCCATEGORY_C",
    "KA_SUBCONYN", "KA_SUBCONTRACTOR", "KA_IRASSOCIATED", "KA_IRASSOCIATED_C",
    "KA_IRLOCTYPE", "KA_SIDEOFLINE", "KA_ELR", "KA_MILES", "KA_YARDS",
    "KA_IRWHAT3WORDS", "KA_LATITUDE", "KA_IRLONG_DESCRIPTION",
    "KA_IRACTIONTAKEN", "KA_IRCONSEQUENCE", "KA_ISSUERESOLVED", "KA_REQFB",
]

ACCIDENT_COLS = [
    "CLASS", "ORGID", "KA_RAISEDBY", "KA_IRINCSOURCE", "KA_IRINCCONNECTED_C",
    "KA_IRASSOCIATED", "KA_IRASSOCIATED_C", "KA_IRLONG_DESCRIPTION", "KA_IRACTIONTAKEN",
    "KA_IRCONSEQUENCE", "KA_ISSUERESOLVED", "STATUS", "DESCRIPTION", "KA_SIDEOFLINE",
    "KA_IRLOCTYPE", "KA_IRINCCATEGORY", "KA_IRINCSOURCE_C", "KA_IRINCCATEGORY_C",
    "KA_SUBCONYN", "KA_ELR", "KA_MILES", "KA_YARDS", "AFFECTEDPERSON", "KA_EVENTOWNER",
    "KA_FBEMAIL", "KA_IRWHAT3WORDS", "KA_LATITUDE", "KA_LONGITUDE", "KA_REQFB",
    "KA_RTIMEOFACCIDENT", "KA_SUBCONTRACTOR", "KA_RGENDER", "KA_RTIMEINTOSHIFT",
    "KA_RAGE", "KA_RTIMEINROLE", "KA_RARIMMEDIATELY", "KA_RTIMEINROLETYPE",
    "KA_IPNAME", "KA_WITNESS", "KA_WORKER", "KA_WITNESSNAME", "KA_IPADDRESS",
    "KA_FAGIVEN", "KA_FADETAILS", "KA_IPHOSPITAL", "KA_IPHOSPDETAIL",
    "KA_BPHEAD", "KA_BPCHEST", "KA_BPSHOULDERLEFT", "KA_BPEAR", "KA_BPABDOMEN",
    "KA_BPSHOULDERRIGHT", "KA_BPEYE", "KA_BPUPPERBACK", "KA_BPELBOWLEFT",
    "KA_BPFACIALBONES", "KA_BPLOWERBACK", "KA_BPELBOWRIGHT", "KA_BPNOSE",
    "KA_BPHIPLEFT", "KA_BPWRISTLEFT", "KA_BPMOUTH", "KA_BPHIPRIGHT", "KA_BPWRISTRIGHT",
    "KA_BPNECK", "KA_BPUPPERLEGLEFT", "KA_BPHANDRIGHT", "KA_BPHANDLEFT",
    "KA_BPUPPERLEGRIGHT", "KA_BPFINGER", "KA_BPKNEELEFT", "KA_BPKNEERIGHT",
    "KA_BPLOWERLEGLEFT", "KA_BPLOWERLEGRIGHT", "KA_BPANKLELEFT", "KA_BPANKLERIGHT",
    "KA_BPFOOTLEFT", "KA_BPFOOTRIGHT", "KA_NATUREOFINJURY",
]


# ─────────────────────────────────────────────
# TRANSFORMATION HELPERS
# ─────────────────────────────────────────────

def clean_text(val):
    """
    Clean encoding junk from text fields.
    - Normalise unicode (removes Â, Â°, and similar mojibake characters)
    - Strip leading/trailing whitespace
    - Collapse multiple internal spaces into one
    - Remove any remaining non-printable characters
    Example: 'Amey OLEÂ (Trans)' → 'Amey OLE (Trans)'
    """
    if not val:
        return val
    # Normalise to NFKD then re-encode to ASCII ignoring non-ASCII chars
    # This removes characters like Â which appear from Windows-1252 → UTF-8 mix
    normalised = unicodedata.normalize("NFKD", val)
    cleaned = normalised.encode("ascii", "ignore").decode("ascii")
    # Collapse extra spaces and strip
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned


def yes_no_to_yn(val):
    """
    Convert Yes/YES → 'Y', No/NO → 'N', anything else → ''
    Used for KA_SUBCONYN and KA_ISSUERESOLVED.
    """
    v = str(val).strip().upper()
    if v == "YES":
        return "Y"
    if v == "NO":
        return "N"
    return ""


def get_irincsource_c(row):
    """
    KA_IRINCSOURCE_C: return first non-empty value from the 3 L2 work-for columns.
    Also runs clean_text to remove any encoding junk.
    """
    for col in IRINCSOURCE_C_COLS:
        val = row.get(col, "")
        if val and str(val).strip().lower() not in ("", "nan"):
            return clean_text(str(val).strip())
    return ""


def get_irassociated_c(row):
    """
    KA_IRASSOCIATED_C: collect values from both associated columns.
      - 'Event Associated with AIW Infrastructure Manager (IM)'
      - 'Event Assiciated with AIW Transformation Project (Trans)'
    If only one has a value → use that.
    If both have values     → concatenate with ' | ' separator.
    If neither              → return ''
    """
    vals = []
    for col in IRASSOCIATED_C_COLS:
        val = row.get(col, "")
        if val and str(val).strip().lower() not in ("", "nan"):
            vals.append(str(val).strip())
    return " | ".join(vals)


def get_irinccategory_c(row):
    """
    KA_IRINCCATEGORY_C: use the Category field to decide which sub-category
    column to read.
    """
    category = str(row.get("Category", "")).strip().upper()
    src_col = CATEGORY_SUBCOL_MAP.get(category, "")
    if not src_col:
        return ""
    val = row.get(src_col, "")
    return str(val).strip() if val and str(val).strip().lower() != "nan" else ""


def parse_elr(raw):
    """
    Parse ELR field — extracts only the ELR code (first uppercase word).
    The ELR field format is: 'CAR (1.0000 to 24.0880)'
    So KA_ELR = 'CAR' — everything after the first word is range info, ignored.
    Miles and yards come from the separate 'Enter Miles and Yards or Chains' field.
    """
    if not raw or str(raw).strip().lower() in ("", "nan"):
        return ""
    raw_str = str(raw).strip()
    # Extract first word of uppercase letters (2–4 chars) before space or bracket
    elr_match = re.match(r'^([A-Z]{2,4})\b', raw_str)
    return elr_match.group(1) if elr_match else ""


def parse_miles_yards(raw):
    """
    Parse 'Enter Miles and Yards or Chains' field into (miles, yards).

    Handles formats:
      '20m0977yds'       → miles='20', yards='977'  (leading zeros stripped)
      '21m0043yds'       → miles='21', yards='43'
      '20m977yards'      → miles='20', yards='977'
      '23m31ch'          → miles='23', yards='682'  (chains * 22)
      '23miles 31ch'     → same
      '23miles 31chains' → same
      Blank / unreadable → miles='', yards=''
    """
    if not raw or str(raw).strip().lower() in ("", "nan"):
        return "", ""

    dist = str(raw).strip().lower()
    dist = re.sub(r'yards?|yds?', 'yds', dist)
    dist = re.sub(r'miles?',       'm',   dist)
    dist = re.sub(r'chains?',      'ch',  dist)
    dist = re.sub(r'\s+',          '',    dist)

    miles = ""
    yards = ""

    m_match = re.search(r'(\d+)m',            dist)
    yds_match = re.search(r'(\d+)yds',          dist)
    ch_match = re.search(r'(\d+(?:\.\d+)?)ch', dist)

    if m_match:
        miles = m_match.group(1)
    if yds_match:
        yards = str(int(yds_match.group(1)))       # strip leading zeros
    elif ch_match:
        # convert chains → yards
        yards = str(int(float(ch_match.group(1)) * 22))

    return miles, yards


def transform_date(val):
    """
    Standardise date/time to Maximo ISO format: YYYY-MM-DDTHH:MM:SS
    Handles inputs like:
      '08/04/2026 01:30' → '2026-04-08T01:30:00'
      '04/04/26 23:30'   → '2026-04-04T23:30:00'  (2-digit year)
      '08/04/2026'       → '2026-04-08T00:00:00'  (no time)
      '4-7-26 13:10:08'  → '2026-04-07T13:10:08'  (dashed format)
    """
    if pd.isnull(val) or str(val).strip() in ("", "nan"):
        return ""
    try:
        parsed = pd.to_datetime(val, dayfirst=True)
        # If year parsed as 1926 due to 2-digit year, fix to 2000s
        if parsed.year < 2000:
            parsed = parsed.replace(year=parsed.year + 100)
        return parsed.strftime("%Y-%m-%dT%H:%M:%S")
    except Exception:
        return str(val)


def safe(row, col):
    """Get a cell value safely, returning empty string for NaN/None."""
    val = row.get(col, "")
    return "" if pd.isnull(val) or str(val).strip().lower() == "nan" else str(val).strip()


def get_raisedby(row):
    """
    If user chose to report anonymously → KA_RAISEDBY = 'Anonymous'
    Otherwise use the actual 'Raised by' field value.
    """
    anon = safe(row, "Would you like to report anonymously?").strip().upper()
    if anon == "YES":
        return "Anonymous"
    return safe(row, "Raised by")


def get_receipt_and_email(row):
    """
    Receipt = Yes → KA_REQFB=1, KA_FBEMAIL=email (last col preferred)
    Receipt = No  → both blank
    """
    receipt = safe(
        row, "Do you require a receipt of this Event Record").strip().upper()
    if receipt != "YES":
        return "", ""
    email_last = safe(row, "Email Address")
    email_first = safe(row, "Email")
    email = email_last if email_last else email_first
    return 1, email


def get_subcontractor(row):
    """
    Subcontractor = Yes → KA_SUBCONYN='Y', KA_SUBCONTRACTOR=name
    Subcontractor = No  → KA_SUBCONYN='N', KA_SUBCONTRACTOR=''
    """
    sub = safe(row, "Was a subcontractor involved?").strip().upper()
    yn = yes_no_to_yn(sub)
    if yn == "Y":
        return "Y", safe(row, "Who is the Subcontractor?")
    return yn, ""


def get_issueresolved(row):
    """
    Yes/YES → 'Y'
    No/NO   → 'N'
    Blank   → ''
    """
    val = safe(row, "Did the immediate action resolve the issue?")
    return yes_no_to_yn(val)


def parse_body_parts(row):
    """
    Each body part is its own column already named as the Maximo attribute.
    Non-empty → 1, blank → ""
    """
    return {col: (1 if safe(row, col) else "") for col in BODY_PART_COLS}


def apply_common_transforms(row, mapping):
    """
    Apply standard column mapping for a row.
    Skips computed fields and applies date formatting and text cleaning.
    """
    computed = {
        "KA_IRINCSOURCE_C", "KA_IRINCCATEGORY_C",
        "KA_IRASSOCIATED_C",                           # handled by get_irassociated_c
        "KA_ELR", "KA_MILES", "KA_YARDS",
        "KA_SUBCONYN", "KA_SUBCONTRACTOR",
        "KA_ISSUERESOLVED",
    }
    out = {}
    for src, tgt in mapping.items():
        if tgt in computed:
            continue
        val = safe(row, src)
        if tgt == "REPORTDATE":
            val = transform_date(val)
        elif tgt in TEXT_CLEAN_FIELDS:
            val = clean_text(val)
        out[tgt] = val
    return out


# ─────────────────────────────────────────────
# STEP 1 — EXTRACT: Download from SharePoint
# ─────────────────────────────────────────────

def get_access_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise Exception(f"Auth failed: {result.get('error_description')}")
    return result["access_token"]


def download_excel_from_sharepoint(token):
    headers = {"Authorization": f"Bearer {token}"}
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE}:/sites/{SHAREPOINT_SITE_NAME}"
    site_id = requests.get(site_url, headers=headers).json()["id"]
    drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
    drive_id = requests.get(drive_url, headers=headers).json()["id"]
    file_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{FILE_PATH}:/content"
    response = requests.get(file_url, headers=headers)
    if response.status_code != 200:
        raise Exception(
            f"File download failed: {response.status_code} {response.text}")
    print(f"[EXTRACT] Downloaded: {FILE_PATH}")
    return io.BytesIO(response.content)


# ─────────────────────────────────────────────
# STEP 2 — TRANSFORM
# ─────────────────────────────────────────────

def transform_row_normal(row):
    out = {"STATUS": FIXED_STATUS, "CLASS": FIXED_CLASS, "ORGID": FIXED_ORGID}

    out.update(apply_common_transforms(row, NORMAL_MAP))

    out["KA_RAISEDBY"] = get_raisedby(row)
    out["KA_REQFB"], out["KA_FBEMAIL"] = get_receipt_and_email(row)
    out["KA_SUBCONYN"], out["KA_SUBCONTRACTOR"] = get_subcontractor(row)
    out["KA_ISSUERESOLVED"] = get_issueresolved(row)
    out["KA_IRINCSOURCE_C"] = get_irincsource_c(row)
    out["KA_IRINCCATEGORY_C"] = get_irinccategory_c(row)
    out["KA_IRASSOCIATED_C"] = get_irassociated_c(row)
    out["KA_ELR"] = parse_elr(safe(row, "ELR"))
    out["KA_MILES"], out["KA_YARDS"] = parse_miles_yards(
        safe(row, "Enter Miles and Yards or Chains"))

    return out


def transform_row_accident(row):
    out = {"STATUS": FIXED_STATUS, "CLASS": FIXED_CLASS, "ORGID": FIXED_ORGID}

    out.update(apply_common_transforms(row, ACCIDENT_MAP))

    out["KA_RAISEDBY"] = get_raisedby(row)
    out["KA_REQFB"], out["KA_FBEMAIL"] = get_receipt_and_email(row)
    out["KA_SUBCONYN"], out["KA_SUBCONTRACTOR"] = get_subcontractor(row)
    out["KA_ISSUERESOLVED"] = get_issueresolved(row)
    out["KA_IRINCSOURCE_C"] = get_irincsource_c(row)
    out["KA_IRINCCATEGORY_C"] = get_irinccategory_c(row)
    out["KA_IRASSOCIATED_C"] = get_irassociated_c(row)
    out["KA_ELR"] = parse_elr(safe(row, "ELR"))
    out["KA_MILES"], out["KA_YARDS"] = parse_miles_yards(
        safe(row, "Enter Miles and Yards or Chains"))
    out.update(parse_body_parts(row))

    return out


def transform(df):
    normal_rows, accident_rows = [], []
    for _, row in df.iterrows():
        category = str(row.get("Category", "")).strip().upper()
        if category == "ACCIDENT":
            accident_rows.append(transform_row_accident(row))
        else:
            normal_rows.append(transform_row_normal(row))

    df_normal = pd.DataFrame(
        normal_rows,   columns=NORMAL_COLS) if normal_rows else pd.DataFrame(columns=NORMAL_COLS)
    df_accident = pd.DataFrame(
        accident_rows, columns=ACCIDENT_COLS) if accident_rows else pd.DataFrame(columns=ACCIDENT_COLS)

    print(f"[TRANSFORM] Normal rows  : {len(df_normal)}")
    print(f"[TRANSFORM] Accident rows: {len(df_accident)}")
    return df_normal, df_accident


# ─────────────────────────────────────────────
# STEP 3 — LOAD: Write output CSVs
# ─────────────────────────────────────────────

def run_pipeline():
    print("=" * 50)
    print("  HSEQ Maximo ETL Pipeline Starting...")
    print("=" * 50)

    if LOCAL_TEST:
        print(f"[MODE]    Local Test — reading from: {LOCAL_INPUT_PATH}")
        df = pd.read_excel(LOCAL_INPUT_PATH, sheet_name=0, dtype=str)
        output_dir = LOCAL_OUTPUT_DIR
    else:
        print("[MODE]    Production — downloading from SharePoint")
        token = get_access_token()
        file_obj = download_excel_from_sharepoint(token)
        df = pd.read_excel(file_obj, sheet_name=0, dtype=str)
        output_dir = OUTPUT_DIR

    df.columns = df.columns.str.strip().str.replace('\xa0', ' ', regex=False)
    print(f"[EXTRACT] Total rows: {len(df)}")

    df_normal, df_accident = transform(df)

    os.makedirs(output_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    es_header = "KA_EXTSYS,KA_HSEQOBJ,AddChange,EN\n"

    if df_normal.empty and df_accident.empty:
        print("[LOAD] No records found — no files created.")
        return

    if not df_normal.empty:
        normal_path = os.path.join(output_dir, f"normal_dataload_{ts}.csv")
        with open(normal_path, "w", newline="", encoding="utf-8") as f:
            f.write(es_header)
            df_normal.to_csv(f, index=False)
        print(f"[LOAD] Normal CSV   → {normal_path}")
    else:
        print("[LOAD] No Normal category records found — skipping normal file.")

    if not df_accident.empty:
        accident_path = os.path.join(output_dir, f"accident_dataload_{ts}.csv")
        with open(accident_path, "w", newline="", encoding="utf-8") as f:
            f.write(es_header)
            df_accident.to_csv(f, index=False)
        print(f"[LOAD] Accident CSV → {accident_path}")
    else:
        print("[LOAD] No Accident category records found — skipping accident file.")

    print("=" * 50)
    print("  Pipeline completed successfully!")
    print("=" * 50)


if __name__ == "__main__":
    run_pipeline()
