import pandas as pd
from services.loaders import clean_text, valid_sites, down_alarm_names, alarm_category_dict, hw_rename_dict, TECH_MAP, extract_site_code, filter_last_40_days

def build_down_description_per_tech(row, alarm_category):
    desc_per_tech = {}

    if alarm_category == "O&M":
        return desc_per_tech

    user_info = str(row.get("User Additional Information", "")).strip()
    diag_info = str(row.get("Diagnostic Info", "")).strip()
    supp_info = str(row.get("Supplementary Information", "")).strip()

    # ===== Cells =====
    cells_map = {}
    if "faulty_cells=" in user_info.lower():
        try:
            data = user_info.lower().split("faulty_cells=")[1].split(";")
            for item in data:
                if ":" in item:
                    tech_name, cells = item.split(":")
                    tech_name = tech_name.strip().upper()
                    count = len([c for c in cells.split(",") if c.strip()])
                    cells_map[tech_name] = cells_map.get(tech_name, 0) + count
        except:
            pass

    for tech, count in cells_map.items():
        desc_per_tech.setdefault(tech, [])
        desc_per_tech[tech].append(f"{tech} Cells: {count}")

    # ===== HW Alarm (مع rename) =====
    hw_items = set()
    if supp_info:
        hw_name = hw_rename_dict.get(supp_info.upper(), supp_info)
        hw_items.add(f"HW Alarm: {hw_name}")

    # ==== ربط HW بالـ techs اللي عندها cells ====
    for tech in cells_map.keys():
        for hw in hw_items:
            if hw not in desc_per_tech.setdefault(tech, []):
                desc_per_tech[tech].append(hw)

    # لو مفيش cells بس فيه HW
    if not cells_map and hw_items:
        sheet_tech = row.get("Sheet Tech", "").upper()
        for hw in hw_items:
            desc_per_tech.setdefault(sheet_tech, []).append(hw)

    # ==== HW من unitName= ====
    if "unitname=" in diag_info.lower():
        parts = diag_info.lower().split("unitname=")[1:]
        for part in parts:
            unit = part[:4].upper()  # أول 4 حروف
            hw_name = hw_rename_dict.get(supp_info.upper(), supp_info)
            hw_items.add(f"{hw_name} ({unit})")
        for tech in (list(cells_map.keys()) or [row.get("Sheet Tech","").upper()]):
            for hw in hw_items:
                if hw not in desc_per_tech.setdefault(tech, []):
                    desc_per_tech[tech].append(hw)

    return desc_per_tech


def build_down_dict(filepath):
    down_info = {}
    for sheet, tech in TECH_MAP.items():
        try:
            df = pd.read_excel(filepath, sheet_name=sheet, engine="openpyxl")
            if df.empty: continue
            df = filter_last_40_days(df)
            for _, row in df.iterrows():
                alarm_text = clean_text(row.get("Alarm Text",""))
                if not alarm_text or alarm_text not in down_alarm_names:
                    continue
                site_code = extract_site_code(row)
                if not site_code or site_code not in valid_sites:
                    continue

                if site_code not in down_info:
                    down_info[site_code] = {
                        "techs": set(),
                        "cells_only": set(),
                        "om_only": set(),
                        "partial_only": set(),
                        "descriptions": [],
                        "descriptions_per_tech": {},
                        "times": []
                    }

                cat = alarm_category_dict.get(alarm_text)

                if cat == "O&M":
                    down_info[site_code]["techs"].add(tech)
                    down_info[site_code]["om_only"].add(tech)
                    # Removing from cells_only if it was there to keep techs clean for display
                    down_info[site_code]["cells_only"].discard(tech)
                else:
                    # Only add to cells_only if not already down via O&M
                    # But we'll handle the 'Total Down' rule in summary.py logic
                    if tech not in down_info[site_code]["om_only"]:
                        down_info[site_code]["cells_only"].add(tech)
                    else:
                        # Even if it's O&M down, we note that it has cells down for Micro logic
                        down_info[site_code]["partial_only"].add(tech) # Using partial_only as a flag for O&M + Cells

                # ===== Build descriptions per tech مع HW rename =====
                desc_dict = build_down_description_per_tech(row, cat)
                for t, descs in desc_dict.items():
                    down_info[site_code]["descriptions_per_tech"].setdefault(t, [])
                    for d in descs:
                        if d not in down_info[site_code]["descriptions_per_tech"][t]:
                            down_info[site_code]["descriptions_per_tech"][t].append(d)

                # ===== Build old style descriptions =====
                for descs in desc_dict.values():
                    for d in descs:
                        if d not in down_info[site_code]["descriptions"]:
                            down_info[site_code]["descriptions"].append(d)

                # ===== Alarm Times =====
                alarm_time = row.get("Alarm Time")
                if pd.notna(alarm_time):
                    down_info[site_code]["times"].append(alarm_time)

        except Exception as e:
            print(f"⚠️ Skipped {sheet}: {e}")

    # Merge techs and cells_only
    for site_code, info in down_info.items():
        final_techs = list(info["techs"]) + [f"{t} Cells" for t in info["cells_only"]]
        down_info[site_code]["techs"] = final_techs
        down_info[site_code]["cells_only"] = list(info["cells_only"])

    return down_info
