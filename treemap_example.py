from pathlib import Path

import pandas as pd
import plotly
import plotly.express as px

# Define excel file location
excel_file = Path(__file__).parent / "data.xlsx"

# Read Data from Excel
df = pd.read_excel(excel_file)

# Assign columns to variables
continent = df["Continent"]
country = df["Country"]
region = df["Region"]
sales = df["Sales"]
margin = df["Profit Margin %"]
remark = df["Remark"]

# Create Chart and store as figure [fig]
fig = px.treemap(
    df,
    path=[continent, region, country],
    values=sales,
    color=margin,
    color_continuous_scale=["red", "yellow", "green"],
    title="Sales/Profit Overview",
    hover_name=remark,
)

# Update/Change Layout
fig.update_layout(title_font_size=42, title_font_family="Arial")

# Define output path
output_path = Path(__file__).parent / "Treemap_Chart.html"

# Save Chart and Export to HTML
plotly.offline.plot(fig, filename=str(output_path))
