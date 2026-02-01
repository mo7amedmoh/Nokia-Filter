import pandas as pd
import os
from datetime import datetime, timedelta, timezone

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ===== Helpers =====
def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).upper().replace('\r',' ').replace('\n',' ').strip()

# ===== Global Cache =====
_cached_data = None
_last_fetch_time = None

# ===== Load Site Master Function =====
def get_live_data(force_refresh=False):
    global _cached_data, _last_fetch_time
    
    # Check cache (10 minutes timeout)
    now = datetime.now()
    if not force_refresh and _cached_data and _last_fetch_time:
        if (now - _last_fetch_time).total_seconds() < 600:
            return _cached_data

    MASTER_URL = 'https://docs.google.com/spreadsheets/d/1JgwNsrL8U81-HelF0HYvaBLwlaK7oIHw/export?format=csv'
    try:
        print(f"Fetching live data from Google Sheets... ({now})")
        site_master = pd.read_csv(MASTER_URL)
        site_master.columns = site_master.columns.str.strip()
        site_master["Site Code"] = site_master["Site Code"].astype(str).str.strip().apply(clean_text)
        site_master = site_master.drop_duplicates(subset=["Site Code"])
        site_master_dict = site_master.set_index("Site Code").to_dict("index")
        valid_sites = set(site_master_dict.keys())
        
        oz_list = sorted(
            site_master["OZ"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )

        # ===== Load Comments List =====
        COMMENTS_URL = 'https://docs.google.com/spreadsheets/d/1VQrXnYudk5P-kgOio_sPXweFGi5v90gH3hW3CMjo5pA/export?format=csv'
        comments_df = pd.read_csv(COMMENTS_URL)
        # Assume the comments are in the first column
        comments_list = [""] + comments_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()

        _cached_data = (site_master_dict, valid_sites, oz_list, comments_list)
        _last_fetch_time = now
        return _cached_data
    except Exception as e:
        print(f"Error loading live data: {e}")
        if _cached_data:
            return _cached_data
        return {}, set(), [], [""]

# Initial Load
site_master_dict, valid_sites, oz_list, comments_list = get_live_data()


# ===== Load Alarm Config =====
alarm_config = pd.read_excel(os.path.join(BASE_DIR, "../alarm_config.xlsx"), engine="openpyxl")
alarm_config["Alarm Text"] = alarm_config["Alarm Text"].astype(str).apply(clean_text)
alarm_category_dict = dict(zip(alarm_config["Alarm Text"], alarm_config["Category"]))
down_alarm_names = set(alarm_category_dict.keys())

# ===== Load ENV Rename =====
alarm_rename_df = pd.read_excel(os.path.join(BASE_DIR, "../alarm_rename.xlsx"), engine="openpyxl")
alarm_rename_dict = dict(zip(alarm_rename_df["Alarm Text"].apply(clean_text),
                             alarm_rename_df["Renamed Alarm"].astype(str).str.strip()))
# ===== Load ENV Criticality =====
alarm_rename_df["Alarm crtiticality"] = alarm_rename_df["Alarm crtiticality"].astype(str).str.upper().str.strip()

critical_rows = alarm_rename_df[alarm_rename_df["Alarm crtiticality"] == "CRITICAL"]
critical_env_alarms = set(critical_rows["Alarm Text"].apply(clean_text)) | \
                      set(critical_rows["Renamed Alarm"].astype(str).str.upper().str.strip())

# ===== Load HW-Rename (for Down alarms) =====
hw_rename_df = pd.read_excel(os.path.join(BASE_DIR, "../HW-Rename.xlsx"), engine="openpyxl")
hw_rename_dict = dict(zip(hw_rename_df["Supplementary Information"].apply(clean_text),
                          hw_rename_df["Output"].astype(str).str.strip()))

# ===== Regex & Tech Map =====
import re
SITE_CODE_REGEX = re.compile(r'(\d{4}(?:AL|DE|SI))')
TECH_MAP = {"2G_Down":"2G","3G_Down":"3G","4G_Down":"4G","5G_Down":"5G"}

# ===== Filter by Date Range =====
def filter_by_date_range(df, start_date_str, end_date_str, time_col="Alarm Time"):
    if time_col not in df.columns or not start_date_str or not end_date_str:
        return df
    df = df.copy()
    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    
    start_dt = pd.to_datetime(start_date_str)
    # Set end_dt to the end of the day or just as is
    end_dt = pd.to_datetime(end_date_str) + timedelta(days=1) - timedelta(seconds=1)
    
    mask = (df[time_col].notna()) & (df[time_col] >= start_dt) & (df[time_col] <= end_dt)
    return df[mask]

# ===== Extract Site Code =====
def extract_site_code(row):
    for col in ["Site Name","Name"]:
        val = clean_text(row.get(col,""))
        match = SITE_CODE_REGEX.search(val)
        if match:
            return match.group(1)
    return None


__all__ = [
    "site_master_dict",
    "valid_sites",
    "critical_env_alarms",
    "oz_list",
    "get_live_data",
    "comments_list"
]
