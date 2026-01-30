import pandas as pd
import os
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ===== Helpers =====
def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).upper().replace('\r',' ').replace('\n',' ').strip()

# ===== Load Site Master Function =====
def get_live_data():
    MASTER_URL = 'https://docs.google.com/spreadsheets/d/1JgwNsrL8U81-HelF0HYvaBLwlaK7oIHw/export?format=csv'
    try:
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
        return site_master_dict, valid_sites, oz_list
    except Exception as e:
        print(f"Error loading live data: {e}")
        return {}, set(), []

# Initial Load
site_master_dict, valid_sites, oz_list = get_live_data()


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

critical_env_alarms = set(
    alarm_rename_df[
        alarm_rename_df["Alarm crtiticality"] == "CRITICAL"
    ]["Alarm Text"].apply(clean_text)
)

# ===== Load HW-Rename (for Down alarms) =====
hw_rename_df = pd.read_excel(os.path.join(BASE_DIR, "../HW-Rename.xlsx"), engine="openpyxl")
hw_rename_dict = dict(zip(hw_rename_df["Supplementary Information"].apply(clean_text),
                          hw_rename_df["Output"].astype(str).str.strip()))

# ===== Regex & Tech Map =====
import re
SITE_CODE_REGEX = re.compile(r'(\d{4}(?:AL|DE|SI))')
TECH_MAP = {"2G_Down":"2G","3G_Down":"3G","4G_Down":"4G","5G_Down":"5G"}

# ===== Filter last 40 days =====
def filter_last_40_days(df, time_col="Alarm Time"):
    if time_col not in df.columns:
        return df
    df = df.copy()
    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    cutoff_date = datetime.now() - timedelta(days=40)
    return df[(df[time_col].notna()) & (df[time_col] >= cutoff_date)]

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
    "get_live_data"
]
