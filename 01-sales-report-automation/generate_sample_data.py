
import random
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd

# Use a fixed seed so the sample data is reproducible.
random.seed(42)

# Create the output folder if it does not exist.
output_folder = Path("sample_data")
output_folder.mkdir(parents=True, exist_ok=True)

# Fake products and regions.
products = ["Laptop", "Mouse", "Keyboard", "Monitor", "Headphones"]
regions = ["Chennai", "Bengaluru", "Mumbai", "Delhi", "Hyderabad"]

sales_data = []

# Generate 190 original fake sales records.
start_date = datetime(2026, 1, 1)

for i in range(1, 191):
    order_date = start_date + timedelta(days=random.randint(0, 300))
    product = random.choice(products)
    quantity = random.randint(1, 10)
    unit_price = random.choice([250, 500, 750, 1000, 1500, 2500])

    sales_data.append({
        "Order_ID": f"ORD{i:04d}",
        "Order_Date": order_date.strftime("%Y-%m-%d"),
        "Product": product,
        "Quantity": quantity,
        "Unit_Price": unit_price,
        "Region": random.choice(regions),
    })

# Convert the records into a pandas DataFrame.
df = pd.DataFrame(sales_data)

# Add inconsistent date formats.
for i in range(0, 30, 3):
    date_value = pd.to_datetime(df.loc[i, "Order_Date"])
    df.loc[i, "Order_Date"] = date_value.strftime("%d/%m/%Y")

for i in range(1, 30, 3):
    date_value = pd.to_datetime(df.loc[i, "Order_Date"])
    df.loc[i, "Order_Date"] = date_value.strftime("%d-%b-%Y")

# Add extra spaces and inconsistent capitalization.
for i in range(0, 20, 2):
    df.loc[i, "Product"] = f"  {str(df.loc[i, 'Product']).upper()}  "

for i in range(1, 20, 2):
    df.loc[i, "Product"] = str(df.loc[i, "Product"]).lower()

# Introduce missing values.
for column, row_numbers in {
    "Product": [25, 40],
    "Quantity": [30, 55],
    "Unit_Price": [35, 60],
    "Region": [45, 70],
    "Order_Date": [50],
}.items():
    df.loc[row_numbers, column] = None

# Convert numeric columns to text so they can contain invalid values.
df["Quantity"] = df["Quantity"].astype("object")
df["Unit_Price"] = df["Unit_Price"].astype("object")

# Allow these columns to contain numbers and text.
df["Quantity"] = df["Quantity"].astype("object")
df["Unit_Price"] = df["Unit_Price"].astype("object")

# Introduce invalid text in numeric columns.
df.loc[75, "Quantity"] = "five"
df.loc[80, "Quantity"] = "unknown"
df.loc[85, "Unit_Price"] = "not available"
df.loc[90, "Unit_Price"] = "unknown"

# Add 10 duplicate records to create a total of 200 rows.
duplicates = df.iloc[:10].copy()
df = pd.concat([df, duplicates], ignore_index=True)

# Save the fake data to an Excel file.
output_file = output_folder / "sales_sample.xlsx"
df.to_excel(output_file, index=False, engine="openpyxl")

print("Sample sales data generated successfully!")
print(f"File saved to: {output_file.resolve()}")
print(f"Total rows: {len(df)}")
print("\nFirst 5 rows:")
print(df.head())
