import os
import io
import re
import time
import unicodedata
import pandas as pd
import requests
from msal import ConfidentialClientApplication
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
TENANT_ID = "YOUR_TENANT_ID"
CLIENT_ID = "YOUR_CLIENT_ID"
CLIENT_SECRET = "YOUR_CLIENT_SECRET"
SHAREPOINT_SITE = "yourorg.sharepoint.com"
SHAREPOINT_SITE_NAME = "YourSiteName"
FILE_PATH = "/Shared Documents/Forms/CVL Event Reporting.xlsx"

FIXED_STATUS = "NEW"
FIXED_CLASS = "HSEQ"
FIXED_ORGID = "TFW"

LOCAL_TEST = True
WATCH_FOLDER = r"C:\Nachiketa\HSEQ RESPONSE EXCEL"
LOCAL_OUTPUT_DIR = r"C:\Users\2152355\OneDrive - Cognizant\Desktop\OUTPUT ETL"
OUTPUT_DIR = "./output"

# ─────────────────────────────────────────────
# COLUMN MAPPINGS
# ─────────────────────────────────────────────

# REPORTDATE source columns — only one will exist in any given Excel file
REPORTDATE_COLS = [
    "Date and Time of Event - Please input as follows (dd/mm/yyy 00:00)",
    "Date of Event",
]

# KA_IRINCCONNECTED_C — first non-empty wins from these 2
IRINCCONNECTED_C_COLS = [
    "Area you work within AIW IM",
    "Area you work within AIW Transformation Project (Trans)",
]

# KA_IRINCSOURCE_C — first non-empty wins from these 4
IRINCSOURCE_C_COLS = [
    "Who do you work for L2? (AIW IM)",
    "Who do you work for L2? (Transport\n  for Wales (TfW)",
    "Who do you work for L2? (Trans)",
    "Area you work within AIW IM",
]

# KA_IRASSOCIATED_C — collect both, join with ' | ' if both populated
IRASSOCIATED_C_COLS = [
    "Event Associated with AIW Infrastructure Manager (IM)",
    "Event Assiciated with AIW Transformation Project (Trans)",
]

# KA_IRINCCATEGORY_C — which sub-column to read based on Category
CATEGORY_SUBCOL_MAP = {
    "CLOSE CALL": "Close Call Type",
    "INCIDENT": "Incident",
    "SERVICE STRIKE": "If, Service Strike, please select option below",
    "ASSAULT": "Assault Type",
    "OTHER": "Other",
    "ACCIDENT": "Nature of Accident (Select most serious):",
}

# KA_IRLOCTYPE_C — which sub-column to read based on Location Type
IRLOCTYPE_C_MAP = {
    "DEPOT, SIDING, YARD OR COMPOUND": "Depot, Siding, Yard or Compound",
    "DEPOT": "Depot, Siding, Yard or Compound",
    "SIDING": "Depot, Siding, Yard or Compound",
    "YARD": "Depot, Siding, Yard or Compound",
    "COMPOUND": "Depot, Siding, Yard or Compound",
    "LEVEL CROSSING": "Level Crossing",
    "OFFICE": "Office",
    "SIGNAL BOX": "Signal Box",
    "STATION": "Station",
}

# Normal category mapping (computed fields excluded)
NORMAL_MAP = {
    "Email": "KA_FBEMAIL",
    "Raised by": "KA_RAISEDBY",
    "Who do you work for?": "KA_IRINCSOURCE",
    # REPORTDATE          -> computed via get_reportdate()
    # KA_IRINCCONNECTED_C -> computed via get_irincconnected_c()
    # KA_IRINCSOURCE_C    -> computed via get_irincsource_c()
    "Category": "KA_IRINCCATEGORY",
    # KA_IRINCCATEGORY_C  -> computed via get_irinccategory_c()
    "Was a subcontractor involved?": "KA_SUBCONYN",
    "Who is the Subcontractor?": "KA_SUBCONTRACTOR",
    # KA_SUBCONYN, KA_SUBCONTRACTOR -> computed via get_subcontractor()
    "Event Associated with/caused by?": "KA_IRASSOCIATED",
    # KA_IRASSOCIATED_C   -> computed via get_irassociated_c()
    "Location Type": "KA_IRLOCTYPE",
    # KA_IRLOCTYPE_C      -> computed via get_irloctype_c()
    "Side Of Line": "KA_SIDEOFLINE",
    "ELR": "KA_ELR",
    # KA_ELR, KA_MILES, KA_YARDS -> computed via parse_elr() / parse_miles_yards()
    "What 3 Words": "KA_IRWHAT3WORDS",
    "Latitude and Longtitude": "KA_LATITUDE",
    "Description of Event - Please DO NOT enter personal Information here": "KA_IRLONG_DESCRIPTION",
    "Immediate Action Taken": "KA_IRACTIONTAKEN",
    "What could have happened?": "KA_IRCONSEQUENCE",
    "Did the immediate action resolve the issue?": "KA_ISSUERESOLVED",
    # KA_ISSUERESOLVED    -> computed via get_issueresolved()
    "Do you require a receipt of this Event Record": "KA_REQFB",
    # KA_REQFB, KA_FBEMAIL -> computed via get_receipt_and_email()
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

# Body part columns — already named as Maximo attributes in the Excel
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

# Text fields that need encoding cleanup
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
    "KA_IRLOCTYPE", "KA_IRLOCTYPE_C", "KA_SIDEOFLINE", "KA_ELR", "KA_MILES", "KA_YARDS",
    "KA_IRWHAT3WORDS", "KA_LATITUDE", "KA_IRLONG_DESCRIPTION",
    "KA_IRACTIONTAKEN", "KA_IRCONSEQUENCE", "KA_ISSUERESOLVED", "KA_REQFB",
]

ACCIDENT_COLS = [
    "CLASS", "ORGID", "KA_RAISEDBY", "KA_IRINCSOURCE", "KA_IRINCCONNECTED_C",
    "KA_IRASSOCIATED", "KA_IRASSOCIATED_C", "KA_IRLONG_DESCRIPTION", "KA_IRACTIONTAKEN",
    "KA_IRCONSEQUENCE", "KA_ISSUERESOLVED", "STATUS", "DESCRIPTION", "KA_SIDEOFLINE",
    "KA_IRLOCTYPE", "KA_IRLOCTYPE_C", "KA_IRINCCATEGORY", "KA_IRINCSOURCE_C",
    "KA_IRINCCATEGORY_C", "KA_SUBCONYN", "KA_ELR", "KA_MILES", "KA_YARDS",
    "AFFECTEDPERSON", "KA_EVENTOWNER", "KA_FBEMAIL", "KA_IRWHAT3WORDS",
    "KA_LATITUDE", "KA_LONGITUDE", "KA_REQFB",
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
    if not val:
        return val
    normalised = unicodedata.normalize("NFKD", val)
    cleaned = normalised.encode("ascii", "ignore").decode("ascii")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def yes_no_to_yn(val):
    v = str(val).strip().upper()
    if v == "YES":
        return "Y"
    if v == "NO":
        return "N"
    return ""


def safe(row, col):
    val = row.get(col, "")
    if val is None:
        return ""
    if pd.isnull(val):
        return ""
    str_val = str(val).strip()
    if str_val.lower() == "nan":
        return ""
    return str_val


def transform_date(val):
    if val is None:
        return ""
    str_val = str(val).strip()
    if str_val == "" or str_val.lower() in ("nan", "none", "nat"):
        return ""
    str_val = re.sub(r"(\d{2})\.(\d{2})$", r"\1:\2", str_val)
    str_val = re.sub(r"\s(\d{2})(\d{2})$", r" \1:\2", str_val)
    try:
        parsed = pd.to_datetime(str_val, dayfirst=True)
        if parsed.year < 2000:
            parsed = parsed.replace(year=parsed.year + 100)
        return parsed.strftime("%Y-%m-%dT%H:%M:%S")
    except Exception:
        return ""


def get_reportdate(row):
    for col in REPORTDATE_COLS:
        val = safe(row, col)
        if val:
            return transform_date(val)
    return ""


def get_irincconnected_c(row):
    for col in IRINCCONNECTED_C_COLS:
        val = row.get(col, "")
        if val and str(val).strip().lower() not in ("", "nan"):
            return clean_text(str(val).strip())
    return ""


def get_irincsource_c(row):
    for col in IRINCSOURCE_C_COLS:
        val = row.get(col, "")
        if val and str(val).strip().lower() not in ("", "nan"):
            return clean_text(str(val).strip())
    return ""


def get_irinccategory_c(row):
    category = str(row.get("Category", "")).strip().upper()
    src_col = CATEGORY_SUBCOL_MAP.get(category, "")
    if not src_col:
        return ""
    val = row.get(src_col, "")
    return str(val).strip() if val and str(val).strip().lower() != "nan" else ""


def get_irassociated_c(row):
    vals = []
    for col in IRASSOCIATED_C_COLS:
        val = row.get(col, "")
        if val and str(val).strip().lower() not in ("", "nan"):
            vals.append(str(val).strip())
    return " | ".join(vals)


def get_irloctype_c(row):
    loc_type = safe(row, "Location Type").replace("\xa0", " ").strip().upper()
    src_col = IRLOCTYPE_C_MAP.get(loc_type, "")
    if not src_col:
        return ""
    val = safe(row, src_col)
    if not val:
        val = safe(row, src_col.replace(" ", "\xa0"))
    return val


def get_raisedby(row):
    anon = safe(row, "Would you like to report anonymously?").strip().upper()
    return "Anonymous" if anon == "YES" else safe(row, "Raised by")


def get_receipt_and_email(row):
    receipt = safe(
        row, "Do you require a receipt of this Event Record").strip().upper()
    if receipt == "YES":
        email = safe(row, "Email Address")
        if not email:
            email = safe(row, "Email")
        return 1, email
    if receipt == "NO":
        return 0, ""
    return "", ""


def get_subcontractor(row):
    yn = yes_no_to_yn(safe(row, "Was a subcontractor involved?"))
    return ("Y", safe(row, "Who is the Subcontractor?")) if yn == "Y" else (yn, "")


def get_issueresolved(row):
    return yes_no_to_yn(safe(row, "Did the immediate action resolve the issue?"))


def parse_body_parts(row):
    return {col: (1 if safe(row, col) else "") for col in BODY_PART_COLS}


def parse_elr(raw):
    if not raw or str(raw).strip().lower() in ("", "nan"):
        return ""
    elr_match = re.match(r"^([A-Z]{2,4})\b", str(raw).strip())
    return elr_match.group(1) if elr_match else ""


def parse_miles_yards(raw):
    if not raw or str(raw).strip().lower() in ("", "nan"):
        return "", ""
    dist = str(raw).strip().lower()
    dist = re.sub(r"chains?",           "ch",    dist)
    dist = re.sub(r"yards?",            "yds",   dist)
    dist = re.sub(r"yds?",              "yds",   dist)
    dist = re.sub(r"(\d)y(\b|(?=\d))", r"\1yds", dist)
    dist = re.sub(r"miles?",            "m",     dist)
    dist = re.sub(r"\s+",               "",      dist)
    miles = ""
    yards = ""
    m_match = re.search(r"(\d+)m",             dist)
    yds_match = re.search(r"(\d+(?:\.\d+)?)yds", dist)
    ch_match = re.search(r"(\d+(?:\.\d+)?)ch",  dist)
    if m_match:
        miles = m_match.group(1)
    if yds_match and ch_match:
        yards = str(round(float(yds_match.group(1)) +
                    float(ch_match.group(1)) * 22))
    elif yds_match:
        yards = str(int(float(yds_match.group(1))))
    elif ch_match:
        yards = str(round(float(ch_match.group(1)) * 22))
    return miles, yards


def apply_common_transforms(row, mapping):
    computed = {
        "REPORTDATE",
        "KA_IRINCCONNECTED_C",
        "KA_IRINCSOURCE_C", "KA_IRINCCATEGORY_C", "KA_IRASSOCIATED_C", "KA_IRLOCTYPE_C",
        "KA_ELR", "KA_MILES", "KA_YARDS",
        "KA_SUBCONYN", "KA_SUBCONTRACTOR",
        "KA_ISSUERESOLVED",
        "KA_REQFB", "KA_FBEMAIL",
    }
    out = {}
    for src, tgt in mapping.items():
        if tgt in computed:
            continue
        val = safe(row, src)
        if tgt in TEXT_CLEAN_FIELDS:
            val = clean_text(val)
        out[tgt] = val
    return out


# ─────────────────────────────────────────────
# STEP 1 — EXTRACT
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
    out["REPORTDATE"] = get_reportdate(row)
    out["KA_REQFB"], out["KA_FBEMAIL"] = get_receipt_and_email(row)
    out["KA_SUBCONYN"], out["KA_SUBCONTRACTOR"] = get_subcontractor(row)
    out["KA_ISSUERESOLVED"] = get_issueresolved(row)
    out["KA_IRINCCONNECTED_C"] = get_irincconnected_c(row)
    out["KA_IRINCSOURCE_C"] = get_irincsource_c(row)
    out["KA_IRINCCATEGORY_C"] = get_irinccategory_c(row)
    out["KA_IRASSOCIATED_C"] = get_irassociated_c(row)
    out["KA_IRLOCTYPE_C"] = get_irloctype_c(row)
    out["KA_ELR"] = parse_elr(safe(row, "ELR"))
    out["KA_MILES"], out["KA_YARDS"] = parse_miles_yards(
        safe(row, "Enter Miles and Yards or Chains"))
    return out


def transform_row_accident(row):
    out = {"STATUS": FIXED_STATUS, "CLASS": FIXED_CLASS, "ORGID": FIXED_ORGID}
    out.update(apply_common_transforms(row, ACCIDENT_MAP))
    out["KA_RAISEDBY"] = get_raisedby(row)
    out["REPORTDATE"] = get_reportdate(row)
    out["KA_REQFB"], out["KA_FBEMAIL"] = get_receipt_and_email(row)
    out["KA_SUBCONYN"], out["KA_SUBCONTRACTOR"] = get_subcontractor(row)
    out["KA_ISSUERESOLVED"] = get_issueresolved(row)
    out["KA_IRINCCONNECTED_C"] = get_irincconnected_c(row)
    out["KA_IRINCSOURCE_C"] = get_irincsource_c(row)
    out["KA_IRINCCATEGORY_C"] = get_irinccategory_c(row)
    out["KA_IRASSOCIATED_C"] = get_irassociated_c(row)
    out["KA_IRLOCTYPE_C"] = get_irloctype_c(row)
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
# STEP 3 — LOAD
# ─────────────────────────────────────────────

def save_outputs(df_normal, df_accident, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    es_header = "KA_EXTSYS,KA_HSEQOBJ,AddChange,EN\n"
    if df_normal.empty and df_accident.empty:
        print("[LOAD] No records found -- no files created.")
        return
    if not df_normal.empty:
        path = os.path.join(output_dir, f"normal_dataload_{ts}.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            f.write(es_header)
            df_normal.to_csv(f, index=False)
        print(f"[LOAD] Normal CSV   -> {path}")
    else:
        print("[LOAD] No Normal records -- skipping normal file.")
    if not df_accident.empty:
        path = os.path.join(output_dir, f"accident_dataload_{ts}.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            f.write(es_header)
            df_accident.to_csv(f, index=False)
        print(f"[LOAD] Accident CSV -> {path}")
    else:
        print("[LOAD] No Accident records -- skipping accident file.")


# ─────────────────────────────────────────────
# CORE PIPELINE
# ─────────────────────────────────────────────

def run_pipeline(file_source, output_dir):
    print("=" * 50)
    print(
        f"  Running pipeline on: {file_source if isinstance(file_source, str) else 'SharePoint file'}")
    print("=" * 50)
    df = pd.read_excel(
        file_source,
        sheet_name=0,
        dtype=str,
        keep_default_na=False,
        na_values=[]
    )
    df.columns = df.columns.str.strip().str.replace("\xa0", " ", regex=False)
    print(f"[EXTRACT] Total rows: {len(df)}")
    df_normal, df_accident = transform(df)
    save_outputs(df_normal, df_accident, output_dir)
    print("=" * 50)
    print("  Pipeline completed successfully!")
    print("=" * 50)


# ─────────────────────────────────────────────
# FOLDER WATCHER
# ─────────────────────────────────────────────

class ExcelFileHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self._processed = {}

    def on_created(self, event):
        self._process(event)

    def on_modified(self, event):
        self._process(event)

    def _process(self, event):
        print(
            f"[WATCHER] Event detected: {event.src_path} | is_directory: {event.is_directory}")
        if event.is_directory:
            return
        path = event.src_path
        if not path.lower().endswith((".xlsx", ".xls")):
            print(f"[WATCHER] Ignored non-Excel file: {path}")
            return

        now = time.time()
        last = self._processed.get(path, 0)
        if now - last < 10:
            print(f"[WATCHER] Skipping duplicate event for: {path}")
            return
        self._processed[path] = now

        print(f"[WATCHER] New Excel file detected: {path}")
        time.sleep(3)
        try:
            run_pipeline(path, LOCAL_OUTPUT_DIR)
        except Exception as e:
            print(f"[ERROR] Pipeline failed for {path}: {e}")


def start_watcher():
    print("=" * 50)
    print(f"  HSEQ ETL -- Watching folder for new files...")
    print(f"  Folder : {WATCH_FOLDER}")
    print(f"  Output : {LOCAL_OUTPUT_DIR}")
    print(f"  Press Ctrl+C to stop.")
    print("=" * 50)
    handler = ExcelFileHandler()
    observer = Observer()
    observer.schedule(handler, path=WATCH_FOLDER, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("[WATCHER] Stopped by user.")
        observer.stop()
    observer.join()


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":
    if LOCAL_TEST:
        start_watcher()
    else:
        token = get_access_token()
        file_obj = download_excel_from_sharepoint(token)
        run_pipeline(file_obj, OUTPUT_DIR)
