import pandas as pd
import sys
import json
import os

file_path = sys.argv[1]

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

# ----------------------------
# SAME FILTER USED FOR CHARTS
# ----------------------------
def is_metadata_column(col: str) -> bool:
    col_low = col.lower()
    return (
        col_low == "id"
        or "start time" in col_low
        or "completion time" in col_low
        or "email" in col_low
        or col_low == "name"
        or "privacy" in col_low
        or "consent" in col_low
    )

# ----------------------------
# REMOVE METADATA FROM DATASET
# ----------------------------
df_clean = df.drop(columns=[
    col for col in df.columns if is_metadata_column(col)
], errors="ignore")

# ----------------------------
# JSON SAFE CONVERSION
# ----------------------------
def make_safe(v):
    if pd.isna(v):
        return ""
    if hasattr(v, "strftime"):  # timestamps
        return v.strftime("%Y-%m-%d %H:%M:%S")
    return str(v)

data = df_clean.applymap(make_safe).to_dict(orient="records")

# ----------------------------
# SAVE
# ----------------------------
os.makedirs("site", exist_ok=True)

with open("site/data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2)

print("Clean JSON exported:", df_clean.shape)
