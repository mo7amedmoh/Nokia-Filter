from services.loaders import site_master_dict, valid_sites, critical_env_alarms
from services.down_logic import build_down_dict
from services.env_logic import build_env_dict

import pandas as pd
from datetime import datetime
import os
import html


def build_summary(filepath, selected_oz=None):
    down_info = build_down_dict(filepath)
    env_info = build_env_dict(filepath)

    critical_env_table = []

    comments_list = [
        "", "Weather Issue", "Access Requested", "Access Blocked",
        "On way to site", "Working in site", "Spare parts required",
        "Spare Shortage", "Power Issue", "Cleared", "Access Tower H&S",
        "H&S case", "HDSL", "Planned Action", "Cascaded",
        "Shared PM Issue", "Theft and sabotage", "Owner PM Issue"
    ]

    def escape_but_allow_br(text):
        return html.escape(text).replace("&lt;br&gt;", "<br>")

    rows = []

    for site_code in valid_sites:
        master = site_master_dict.get(site_code, {})

        # ===== OZ Filter =====
        if selected_oz:
            if str(master.get("OZ", "")).strip().upper() != str(selected_oz).strip().upper():
                continue

        row = {
            "Site Code": site_code,
            "Site Name": master.get("Site Name", ""),
            "SC Office": master.get("SC Office", ""),
            "Down Alarm": "",
            "Down Alarm Description": "",
            "Alarm Time": "",
            "Down Type": "",
            "ENV Alarms": "",
            "ENV Alarm Time": "",
            "Comment": comments_list.copy(),
            "Duration": "",
            "_down_time": pd.Timestamp.max,
            "_env_time": pd.Timestamp.min
        }

        # ================= DOWN =================
        if site_code in down_info:
            site_down = down_info[site_code]

            techs_down = site_down.get("techs", [])
            om_only = set(site_down.get("om_only", []))
            cells_only = set(site_down.get("cells_only", []))

            row["Down Alarm"] = ", ".join(sorted(techs_down))

            if site_down.get("times"):
                down_times = pd.to_datetime(site_down["times"], errors="coerce").dropna()
                if not down_times.empty:
                    first_time = down_times.min()
                    row["Alarm Time"] = first_time.strftime("%Y-%m-%d %H:%M:%S")
                    row["_down_time"] = first_time

                    duration = datetime.now() - first_time
                    total_minutes = int(duration.total_seconds() // 60)
                    row["Duration"] = f"{total_minutes//60:02d}:{total_minutes%60:02d}"

            site_type = master.get("Site Type", "").upper()

            om_count = len(om_only)
            partial_count = len([t for t in techs_down if t not in om_only])

            if site_type == "MICRO":
                if om_count + partial_count >= 2 or (om_count >= 1 and partial_count >= 1):
                    row["Down Type"] = "Total"
                elif om_count + partial_count == 1:
                    row["Down Type"] = "Partial"

            elif site_type in ["PICO", "NANO"]:
                if om_count + partial_count >= 1:
                    row["Down Type"] = "Total"

            else:  # MACRO
                if om_count >= 3:
                    row["Down Type"] = "Total"
                elif om_count == 2 and partial_count >= 1:
                    row["Down Type"] = "Total"
                elif om_count + partial_count > 0:
                    row["Down Type"] = "Partial"

            if row["Down Type"] == "Total":
                row["Down Alarm Description"] = "Total Down"

            elif row["Down Type"] == "Partial":
                lines = []
                for tech, descs in site_down.get("descriptions_per_tech", {}).items():
                    for d in descs:
                        lines.append(f"{tech}: {escape_but_allow_br(d)}")
                row["Down Alarm Description"] = "<br>".join(dict.fromkeys(lines))

        # ================= ENV =================
        if site_code in env_info:
            site_type = master.get("Site Type", "").upper()

            env_alarms = []
            critical_env = []

            for alarm in env_info[site_code].get("alarms", []):
                clean_alarm = alarm.upper().strip()
                env_alarms.append(escape_but_allow_br(alarm))

                if (
                    clean_alarm in critical_env_alarms
                    and site_type not in ["MICRO", "PICO", "NANO"]
                ):
                    critical_env.append(escape_but_allow_br(alarm))

            row["ENV Alarms"] = " | ".join(env_alarms)

            if env_info[site_code].get("times"):
                env_times = pd.to_datetime(env_info[site_code]["times"], errors="coerce").dropna()
                if not env_times.empty:
                    last_time = env_times.max()
                    row["ENV Alarm Time"] = last_time.strftime("%Y-%m-%d %H:%M:%S")
                    row["_env_time"] = last_time

            if critical_env:
                critical_env_table.append({
                    "Site Code": site_code,
                    "Site Name": row["Site Name"],
                    "SC Office": row["SC Office"],
                    "ENV Alarm": " | ".join(critical_env),
                    "Site Type": site_type
                })

        rows.append(row)

    # ================= DataFrame =================
    df = pd.DataFrame(rows)
    df = df[(df["Down Alarm"] != "") | (df["ENV Alarms"] != "")].reset_index(drop=True)

    # ================= Badges =================
    for i, r in df.iterrows():
        master = site_master_dict.get(r["Site Code"], {})
        badges = []

        for bname, col in [("VIP", "VIP"), ("CEO", "CEO"), ("Router", "Router")]:
            if str(master.get(col, "")).strip().upper() in ["TRUE", "YES", "1"]:
                badges.append(f'<span class="badge badge-{bname.lower()} me-1">{bname}</span>')

        site_html = str(r["Site Name"]) + "<br>"

        if master.get("Backup time"):
            site_html += f'<span class="badge badge-bdt me-1">{master["Backup time"]} mins</span>'
        if master.get("Nodal Deg."):
            site_html += f'<span class="badge badge-nodal me-1">{master["Nodal Deg."]}</span>'
        if master.get("Power Source"):
            site_html += f'<span class="badge badge-power me-1">{master["Power Source"]}</span>'
        if master.get("Site Type"):
            site_html += f'<span class="badge badge-site me-1">{master["Site Type"]}</span>'

        site_html += " ".join(badges)
        df.at[i, "Site Name"] = site_html

    # ================= Dashboards =================
    dashboard = {}
    for office, g in df.groupby("SC Office"):
        dashboard[office] = {
            "Total Down": (g["Down Type"] == "Total").sum(),
            "Partial Down": (g["Down Type"] == "Partial").sum(),
            "ENV": g["ENV Alarms"].apply(lambda x: len(str(x).split("<br>")) if x else 0).sum()
        }

    dashboard_summary = {
        "Total Down Sites": (df["Down Type"] == "Total").sum(),
        "Total Partial Sites": (df["Down Type"] == "Partial").sum(),
        "Total Env Alarms": df["ENV Alarms"].apply(lambda x: len(str(x).split("<br>")) if x else 0).sum()
    }

    tables_down_env = df[df["Down Alarm"] != ""].sort_values("_down_time").to_dict("records")
    tables_env_only = df[(df["Down Alarm"] == "") & (df["ENV Alarms"] != "")] \
        .sort_values("_env_time", ascending=False).to_dict("records")

    excel_path = os.path.join("uploads", "Summary.xlsx")
    df.to_excel(excel_path, index=False)

    tech_labels = ["2G", "3G", "4G", "5G"]
    tech_counts = [0, 0, 0, 0]

    down_type_counts = {
        "Total": (df["Down Type"] == "Total").sum(),
        "Partial": (df["Down Type"] == "Partial").sum()
    }

    env_labels = []
    env_values = []

    return (
        df,
        dashboard,
        dashboard_summary,
        tables_down_env,
        tables_env_only,
        critical_env_table,
        tech_labels,
        tech_counts,
        down_type_counts,
        env_labels,
        env_values,
        excel_path
    )
