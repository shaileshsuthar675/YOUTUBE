import pandas as pd

data = [
    {"Date": "2026-01-01", "Product": "Laptop", "Region": "East", "Sales": 50000},
    {"Date": "2026-01-02", "Product": "Phone", "Region": "West", "Sales": 20000},
    {"Date": "2026-01-03", "Product": "Laptop", "Region": "East", "Sales": 60000},
    {"Date": "2026-01-04", "Product": "Headphones", "Region": "North", "Sales": 8000},
    {"Date": "2026-01-05", "Product": "Phone", "Region": "West", "Sales": None},      # Missing
    {"Date": "2026-01-06", "Product": "Laptop", "Region": "South", "Sales": "55,000"}, # Comma
    {"Date": "2026-01-07", "Product": "Phone", "Region": "East", "Sales": "18000"},    # String
    {"Date": None, "Product": None, "Region": None, "Sales": None},                    # Empty row
]

df = pd.DataFrame(data)
df.to_excel("./1. Automate Excel with Python (Beginner Project)/sales_raw.xlsx", index=False)
print("✅ sales_raw.xlsx created")