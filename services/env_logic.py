import pandas as pd
from services.loaders import clean_text, valid_sites, alarm_rename_dict, extract_site_code

def get_env_alarm_name(row):
    obj_class = clean_text(row.get("Object Class",""))
    alarm_text = clean_text(row.get("Alarm Text",""))
    user_info = str(row.get("User Additional Information","")).strip()
    supp_info = str(row.get("Supplementary Information","")).strip()

    if obj_class == "BSC":
        name = alarm_text
    elif alarm_text == "":
        name = user_info
    elif not supp_info or supp_info.upper()=="NAN":
        name = alarm_text
    else:
        name = supp_info

    return alarm_rename_dict.get(name.upper(), name).strip()

def build_env_dict(filepath):
    env_info = {}
    try:
        env_df = pd.read_excel(filepath, sheet_name="Environmental", engine="openpyxl")
        if env_df.empty:
            return env_info
        for _, row in env_df.iterrows():
            site_code = extract_site_code(row)
            if not site_code or site_code not in valid_sites:
                continue
            renamed_alarm = get_env_alarm_name(row)
            alarm_time = row.get("Alarm Time")
            if site_code not in env_info:
                env_info[site_code] = {"alarms":[], "times":[]}
            if renamed_alarm not in env_info[site_code]["alarms"]:
                env_info[site_code]["alarms"].append(renamed_alarm)
            if pd.notna(alarm_time):
                env_info[site_code]["times"].append(alarm_time)
    except Exception as e:
        print("⚠️ ENV skipped:", e)
    return env_info
