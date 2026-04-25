import pandas as pd
import plotly.express as px
import plotly.io as pio
from collections import Counter
import sys
import os

# ----------------------------
# INPUT FILE
# ----------------------------
if len(sys.argv) < 2:
    print("Usage: python dashboard.py <path_to_excel>")
    sys.exit(1)

file_path = sys.argv[1]

if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
    sys.exit(1)

# ----------------------------
# LOAD DATA
# ----------------------------
df = pd.read_excel(file_path)

# Remove junk columns
df = df.drop(columns=[c for c in df.columns if "Column" in c], errors="ignore")

print(f"\nLoaded {len(df.columns)} columns\n")

# ----------------------------
# HELPERS
# ----------------------------
def split_counts(series):
    """Handles multi-select fields (; or , separated)."""
    items = []
    for val in series.dropna():
        val = str(val).replace(",", ";")
        items.extend([x.strip() for x in val.split(";") if x.strip()])
    return pd.Series(Counter(items)).sort_values(ascending=False)

def style_fig(fig):
    """Clean premium look + consistent sizing"""
    fig.update_layout(
        height=320,
        template="plotly_white",
        margin=dict(l=10, r=10, t=40, b=10),
        showlegend=True,
        title_font=dict(size=15)
    )
    return fig

def render(fig):
    """Convert plotly figure to embeddable HTML"""
    return pio.to_html(
        fig,
        full_html=False,
        config={
            "displayModeBar": False,
            "responsive": True
        }
    )

# ----------------------------
# DASHBOARD STORAGE
# ----------------------------
all_charts = []

# ----------------------------
# AUTO-GENERATE CHARTS
# ----------------------------
for col in df.columns:

    # Skip metadata
    if col in ["Id", "Start time", "Completion time", "Email", "Name"]:
        continue

    series = df[col].dropna()

    if len(series) == 0:
        print(f"Skipping (empty): {col}")
        continue

    print(f"Processing: {col}")

    try:

        # ------------------------
        # MULTI-SELECT DETECTION
        # ------------------------
        multi_ratio = series.astype(str).str.contains(";|,").mean()

        if multi_ratio > 0.3:
            counts = split_counts(series).head(15)

            fig = px.bar(
                counts,
                orientation="h",
                title=col
            )

        else:

            # ------------------------
            # NUMERIC DETECTION
            # ------------------------
            numeric = pd.to_numeric(series, errors="coerce")

            numeric_ratio = numeric.notna().mean()

            if numeric_ratio > 0.6:
                fig = px.strip(
                    x=numeric,
                    title=col
                )

            # ------------------------
            # CATEGORICAL FALLBACK
            # ------------------------
            else:
                counts = series.value_counts().head(15)

                fig = px.bar(
                    counts,
                    title=col
                )

        fig = style_fig(fig)
        all_charts.append(render(fig))

    except Exception as e:
        print(f"ERROR in {col}: {e}")

# ----------------------------
# BUILD HTML DASHBOARD
# ----------------------------
html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Farm Survey Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>

    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial;
            background: #f7f8fa;
            margin: 0;
            padding: 30px;
            color: #222;
        }}

        .container {{
            max-width: 1200px;
            margin: auto;
        }}

        h1 {{
            font-size: 28px;
            margin-bottom: 20px;
        }}

        .grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
            gap: 18px;
        }}

        .card {{
            background: white;
            border-radius: 12px;
            padding: 12px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            overflow: hidden;
        }}
    </style>
</head>

<body>

<div class="container">

<h1>Farm Survey Dashboard</h1>

<div class="grid">
    {''.join(f'<div class="card">{c}</div>' for c in all_charts)}
</div>

</div>

</body>
</html>
"""

# ----------------------------
# SAVE OUTPUT
# ----------------------------
os.makedirs("site", exist_ok=True)

output_path = "site/index.html"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\nDashboard created: {output_path}")
print(f"Total charts generated: {len(all_charts)}")
