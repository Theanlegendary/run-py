import os
import json
import pandas as pd
from datetime import datetime
import generate_report
import generate_summary

with open("config.json", encoding="utf-8") as f:
    cfg = json.load(f)

src_xlsx = "test_detail.xlsx"
ref_path = "post_office_lookup.csv"
out_dir = "test_out"

# Generate report metadata
result = generate_report.generate_reports_from_data(
    src_xlsx, ref_path, out_dir, return_metadata=True, mode="wide"
)

# Apply Zone 5 filter just like bot.py does for /total zone5
zone_key = "zone5"
zone_filter = [h.upper() for h in cfg["total_zones"][zone_key]]

# Filter handle results
filtered_results = [
    hr for hr in result["handle_results"]
    if hr["handle"] in zone_filter
]

# Recalculate overall counts
overall = {"Pickup": 0, "Delivery": 0, "Pending": 0}
for hr in filtered_results:
    for k in overall:
        overall[k] += hr["handle_counts"].get(k, 0)

# Calculate day_date_counts and urgent_counts
zone_day_date_counts = {}
zone_urgent_counts = {}
today_date = datetime.now().date()

# Filter type_data by handles
zone_type_data = {}
for rn in ["Pickup", "Delivery", "Pending"]:
    df = result["type_data"].get(rn)
    if df is not None and not df.empty:
        filter_col = "POST OFFICE HANDLE"
        if filter_col in df.columns:
            zone_type_data[rn] = df[df[filter_col].isin(zone_filter)].copy()
        else:
            zone_type_data[rn] = df.copy()
    else:
        zone_type_data[rn] = pd.DataFrame()

for rn in ["Pickup", "Delivery", "Pending"]:
    df_z = zone_type_data.get(rn)
    if df_z is None or df_z.empty:
        continue
    date_col_z = result.get("cur_time_col") or (
        "CREATED DATE" if "CREATED DATE" in df_z.columns else
        "CURRENT TIME"  if "CURRENT TIME"  in df_z.columns else None
    )
    if date_col_z and date_col_z in df_z.columns:
        parsed_z = pd.to_datetime(df_z[date_col_z], dayfirst=True, format="mixed", errors="coerce")
        df_z = df_z.copy()
        df_z["_zdate"] = parsed_z.dt.date

    handle_col = "POST OFFICE HANDLE"
    if handle_col not in df_z.columns:
        continue

    for _, row_z in df_z.iterrows():
        h = str(row_z.get(handle_col, "")).strip().upper()
        if not h:
            continue
        # date counts
        d_val = row_z.get("_zdate") if "_zdate" in df_z.columns else None
        if d_val and not pd.isna(d_val):
            zone_day_date_counts.setdefault(h, {})
            zone_day_date_counts[h][d_val] = zone_day_date_counts[h].get(d_val, 0) + 1
        # urgent = overdue (created > 1 day ago)
        created_d = None
        if "CREATED DATE" in df_z.columns:
            cd = pd.to_datetime(row_z.get("CREATED DATE"), dayfirst=True, format="mixed", errors="coerce")
            if not pd.isna(cd):
                created_d = cd.date()
        if created_d and (today_date - created_d).days > 1:
            zone_urgent_counts[h] = zone_urgent_counts.get(h, 0) + 1

# Build summary image using generate_summary
img_buf = generate_summary.build_summary_image(
    filtered_results,
    overall,
    zone_label="ZONE5",
    day_date_counts=zone_day_date_counts if zone_day_date_counts else None,
    urgent_counts=zone_urgent_counts if zone_urgent_counts else None,
)

# Save image
out_png = "zone5_summary_test.png"
with open(out_png, "wb") as f_img:
    f_img.write(img_buf.getvalue())

# Build total Excel using generate_summary
total_xlsx = "zone5_total_test.xlsx"
zone_result = {**result, "handle_results": filtered_results, "overall_counts": overall, "type_data": zone_type_data}
generate_summary.build_total_excel(zone_result, total_xlsx)

print("Successfully generated Zone 5 summary image and total Excel!")
