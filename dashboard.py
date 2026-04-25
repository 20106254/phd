import pandas as pd
import sys
import json
import os

file_path = sys.argv[1]

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

df = df.drop(columns=[c for c in df.columns if "Column" in c], errors="ignore")

# ----------------------------
# CONVERT EVERYTHING TO JSON-SAFE
# ----------------------------
def make_json_safe(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp,)):
        return val.strftime("%Y-%m-%d %H:%M:%S")
    if hasattr(val, "item"):  # numpy types
        return val.item()
    return str(val)

data = df.applymap(make_json_safe).to_dict(orient="records")

# ----------------------------
# SAVE
# ----------------------------
os.makedirs("site", exist_ok=True)

with open("site/data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2)

print("Exported clean JSON:", len(data), "rows")
