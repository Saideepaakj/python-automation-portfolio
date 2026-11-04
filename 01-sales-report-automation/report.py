
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

# Define the input and output file locations.
input_file = Path("sample_data") / "sales_sample.xlsx"
output_folder = Path("output")
output_folder.mkdir(parents=True, exist_ok=True)
output_file = output_folder / "sales_report.xlsx"

# Stop if the input file does not exist.
if not input_file.exists():
    print("Error: Sample data file not found.")
    print("Run python generate_sample_data.py first.")
    raise SystemExit

# Read the Excel file.
df = pd.read_excel(input_file)
print("Original number of rows:", len(df))

# Check that the required columns are available.
required_columns = [
    "Order_ID", "Order_Date", "Product",
    "Quantity", "Unit_Price", "Region"
]

missing_columns = [
    column for column in required_columns
    if column not in df.columns
]

if missing_columns:
    print("Error: Missing required columns:", missing_columns)
    raise SystemExit

# Remove extra spaces from text columns.
for column in ["Product", "Region", "Order_ID"]:
    df[column] = df[column].astype("string").str.strip()

# Standardize capitalization.
df["Product"] = df["Product"].str.title()
df["Region"] = df["Region"].str.title()

# Convert dates into a standard date format.
df["Order_Date"] = pd.to_datetime(
    df["Order_Date"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

# Convert quantity and unit price into numbers.
df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
df["Unit_Price"] = pd.to_numeric(df["Unit_Price"], errors="coerce")

# Remove duplicate rows.
df = df.drop_duplicates()

# Remove rows missing essential information.
df = df.dropna(
    subset=[
        "Order_ID", "Order_Date", "Product",
        "Quantity", "Unit_Price"
    ]
)

# Remove rows with invalid or non-positive quantities and prices.
df = df[(df["Quantity"] > 0) & (df["Unit_Price"] > 0)]

# Calculate the sales amount for each order.
df["Total_Sales"] = df["Quantity"] * df["Unit_Price"]

# Create a summary of sales by product.
product_summary = (
    df.groupby("Product", as_index=False)["Total_Sales"]
    .sum()
    .sort_values("Total_Sales", ascending=False)
)

# Create a summary of sales by month.
df["Month"] = df["Order_Date"].dt.to_period("M").astype(str)

monthly_summary = (
    df.groupby("Month", as_index=False)["Total_Sales"]
    .sum()
    .sort_values("Month")
)

# Save the cleaned data and both summaries into one Excel file.
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)

    product_summary.to_excel(
        writer, sheet_name="Summary", index=False, startrow=0
    )

    monthly_summary.to_excel(
        writer, sheet_name="Summary", index=False, startrow=len(product_summary) + 4
    )

# Add a bar chart to the Summary sheet.
workbook = load_workbook(output_file)
summary_sheet = workbook["Summary"]

# Add headings above the two summary tables.
summary_sheet["A1"] = "Sales by Product"
monthly_heading_row = len(product_summary) + 5
summary_sheet.cell(row=monthly_heading_row, column=4, value="Sales by Month")

# Create a bar chart using the product summary.
chart = BarChart()
chart.type = "bar"
chart.style = 10
chart.title = "Total Sales by Product"
chart.x_axis.title = "Total Sales"
chart.y_axis.title = "Product"
chart.height = 7
chart.width = 14

sales_data = Reference(
    summary_sheet,
    min_col=2,
    min_row=1,
    max_row=len(product_summary) + 1
)

product_names = Reference(
    summary_sheet,
    min_col=1,
    min_row=2,
    max_row=len(product_summary) + 1
)

chart.add_data(sales_data, titles_from_data=True)
chart.set_categories(product_names)

# Place the chart to the right of the summary tables.
summary_sheet.add_chart(chart, "G2")

# Save the final Excel report.
workbook.save(output_file)

print("\nData cleaning and report generation completed!")
print("Cleaned number of rows:", len(df))
print("Total sales:", round(df["Total_Sales"].sum(), 2))
print("Report saved to:", output_file.resolve())
print("\nSales by product:")
print(product_summary)
print("\nSales by month:")
print(monthly_summary)
