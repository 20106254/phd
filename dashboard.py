import pandas as pd
import plotly.express as px
import os
import sys
from collections import Counter

# ----------------------------
# INPUT
# ----------------------------
if len(sys.argv) < 2:
    print("Usage: python dashboard.py <excel_file>")
    sys.exit(1)

file_path = sys.argv[1]

if not os.path.exists(file_path):
    print("File not found:", file_path)
    sys.exit(1)

# ----------------------------
# LOAD DATA
# ----------------------------
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

# Remove junk Excel columns
df = df.drop(columns=[c for c in df.columns if "Column" in c], errors="ignore")

print(f"Loaded dataset: {df.shape}")

# ----------------------------
# OUTPUT FOLDER
# ----------------------------
os.makedirs("site/charts", exist_ok=True)

# ----------------------------
# METADATA FILTER
# ----------------------------
def is_metadata_column(col: str) -> bool:
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

# ----------------------------
# CLEAN SERIES
# ----------------------------
def clean_series(s):
    s = s.dropna()
    if s.dtype == "object":
        s = s.astype(str).str.strip()
        s = s[(s != "") & (s.str.lower() != "nan")]
    return s

# ----------------------------
# MULTI-SELECT COUNTS
# ----------------------------
def split_counts(series):
    items = []
    for v in series:
        v = str(v).replace(",", ";")
        items.extend([x.strip() for x in v.split(";") if x.strip()])
    return pd.Series(Counter(items)).sort_values(ascending=False)

# ----------------------------
# STYLE
# ----------------------------
def style(fig):
    fig.update_layout(
        height=340,
        template="plotly_white",
        margin=dict(l=10, r=10, t=40, b=10)
    )
    return fig

# ----------------------------
# BUILD CHARTS
# ----------------------------
chart_files = []

for i, col in enumerate(df.columns):

    if is_metadata_column(col):
        print(f"Skipping metadata: {col}")
        continue

    series = clean_series(df[col])

    if len(series) == 0:
        print(f"Skipping empty: {col}")
        continue

    print(f"Processing: {col}")

    try:
        file_name = f"chart_{i}.html"
        path = f"site/charts/{file_name}"

        # ----------------------------
        # NUMERIC DETECTION (FIXED)
        # ----------------------------
        numeric = pd.to_numeric(series, errors="coerce")
        valid_numeric = numeric.dropna()

        numeric_ratio = len(valid_numeric) / len(series)

        # ----------------------------
        # MULTI-SELECT DETECTION
        # ----------------------------
        multi_ratio = series.astype(str).str.contains(";|,").mean()

        # ----------------------------
        # 1. BOX PLOT (RESTORED)
        # ----------------------------
        if numeric_ratio > 0.6 and len(valid_numeric) > 3:

            fig = px.box(
                y=valid_numeric,
                title=col
            )

        # ----------------------------
        # 2. MULTI-SELECT BAR
        # ----------------------------
        elif multi_ratio > 0.3:

            counts = split_counts(series).head(15)

            if counts.empty:
                continue

            fig = px.bar(
                counts,
                orientation="h",
                title=col
            )

        # ----------------------------
        # 3. NORMAL CATEGORICAL BAR
        # ----------------------------
        else:

            counts = series.value_counts().head(15)

            if counts.empty:
                continue

            fig = px.bar(
                counts,
                title=col
            )

        fig = style(fig)

        fig.write_html(
            path,
            include_plotlyjs="cdn",
            full_html=True
        )

        chart_files.append(file_name)

    except Exception as e:
        print(f"Error in {col}: {e}")

print(f"\nGenerated {len(chart_files)} charts")

# ----------------------------
# INDEX.HTML
# ----------------------------
html = """
<!DOCTYPE html>
<html>
<head>
    <title>Farm Survey Dashboard</title>

    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Arial;
            background: #f7f8fa;
            margin: 0;
            padding: 25px;
        }

        h1 {
            margin-bottom: 20px;
        }

        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
            gap: 15px;
        }

        iframe {
            width: 100%;
            height: 380px;
            border: none;
            border-radius: 10px;
            background: white;
            box-shadow: 0 2px 10px rgba(0,0,0,0.06);
        }
    </style>
</head>

<body>

<h1>Farm Survey Dashboard</h1>

<div class="grid">
"""

for f in chart_files:
    html += f'<iframe src="charts/{f}"></iframe>\n'

html += """
</div>

</body>
</html>
"""

os.makedirs("site", exist_ok=True)

with open("site/index.html", "w", encoding="utf-8") as f:
    f.write(html)

print("Dashboard created → site/index.html")
