import pandas as pd
import sys
import json
import os

file_path = sys.argv[1]

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

# ----------------------------
# REMOVE METADATA COLUMNS
# ----------------------------
def is_metadata(col):
    c = col.lower()
    return (
        c == "id"
        or "start time" in c
        or "completion time" in c
        or "email" in c
        or c == "name"
        or "privacy" in c
        or "consent" in c
    )

df = df.drop(columns=[c for c in df.columns if is_metadata(c)], errors="ignore")

# ----------------------------
# JSON SAFE CONVERSION
# ----------------------------
def safe(v):
    if pd.isna(v):
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    return str(v)

data = df.applymap(safe).to_dict(orient="records")

os.makedirs("site", exist_ok=True)

with open("site/data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2)

print("Exported clean dataset:", df.shape)
