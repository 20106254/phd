import pandas as pd
import plotly.express as px
import os
import sys
from collections import Counter

# ----------------------------
# INPUT FILE
# ----------------------------
if len(sys.argv) < 2:
    print("Usage: python dashboard.py <excel_file>")
    sys.exit(1)

file_path = sys.argv[1]

if not os.path.exists(file_path):
    print("File not found:", file_path)
    sys.exit(1)

df = pd.read_excel(file_path)

# Remove junk columns
df = df.drop(columns=[c for c in df.columns if "Column" in c], errors="ignore")

# ----------------------------
# OUTPUT STRUCTURE
# ----------------------------
os.makedirs("site/charts", exist_ok=True)

# ----------------------------
# HELPERS
# ----------------------------
def split_counts(series):
    items = []
    for v in series.dropna():
        v = str(v).replace(",", ";")
        items.extend([x.strip() for x in v.split(";") if x.strip()])
    return pd.Series(Counter(items)).sort_values(ascending=False)

def style(fig):
    fig.update_layout(
        height=350,
        template="plotly_white",
        margin=dict(l=10, r=10, t=40, b=10)
    )
    return fig

# ----------------------------
# GENERATE CHARTS
# ----------------------------
chart_files = []

for i, col in enumerate(df.columns):

    if col in ["Id", "Start time", "Completion time", "Email", "Name"]:
        continue

    series = df[col].dropna()

    if len(series) == 0:
        continue

    try:
        file_name = f"chart_{i}.html"
        path = f"site/charts/{file_name}"

        # MULTI-SELECT
        if series.astype(str).str.contains(";|,").mean() > 0.3:
            counts = split_counts(series).head(15)
            fig = px.bar(counts, orientation="h", title=col)

        else:
            numeric = pd.to_numeric(series, errors="coerce")

            if numeric.notna().mean() > 0.6:
                fig = px.strip(x=numeric, title=col)
            else:
                counts = series.value_counts().head(15)
                fig = px.bar(counts, title=col)

        fig = style(fig)

        fig.write_html(
            path,
            include_plotlyjs=False,
            full_html=True
        )

        chart_files.append(file_name)

    except Exception as e:
        print("Skipped:", col, e)

print(f"Generated {len(chart_files)} charts")

# ----------------------------
# BUILD INDEX.HTML
# ----------------------------
html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Farm Survey Dashboard</title>

    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Arial;
            margin: 0;
            padding: 30px;
            background: #f7f8fa;
        }}

        h1 {{
            margin-bottom: 20px;
        }}

        .grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 15px;
        }}

        iframe {{
            width: 100%;
            height: 380px;
            border: none;
            background: white;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        }}
    </style>
</head>

<body>

<h1>Farm Survey Dashboard</h1>

<div class="grid">
"""

for c in chart_files:
    html += f'<iframe src="charts/{c}"></iframe>\n'

html += """
</div>

</body>
</html>
"""

with open("site/index.html", "w", encoding="utf-8") as f:
    f.write(html)

print("Dashboard created successfully in /site")
