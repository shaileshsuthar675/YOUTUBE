import numpy as np
import pandas as pd

np.random.seed(7)

N_ORDERS = 12000
N_CUSTOMERS = 1800
N_PRODUCTS = 220

regions = ["North", "South", "East", "West", "Central"]
cities = ["Delhi", "Mumbai", "Bengaluru", "Hyderabad", "Pune", "Kolkata", "Chennai", "Jaipur"]
segments = ["Consumer", "Corporate", "SMB"]
categories = ["Electronics", "Grocery", "Fashion", "Home", "Beauty", "Sports"]
suppliers = ["Apex", "Zenith", "Nova", "Orion", "Pulse"]

# Products
products = pd.DataFrame({
    "product_id": np.arange(1001, 1001+N_PRODUCTS),
    "product_name": [f"Product_{i}" for i in range(N_PRODUCTS)],
    "category": np.random.choice(categories, N_PRODUCTS),
    "supplier": np.random.choice(suppliers, N_PRODUCTS),
    "cost": np.round(np.random.uniform(50, 3000, N_PRODUCTS), 2),
    "tax_rate": np.random.choice([0.0, 0.05, 0.12, 0.18], N_PRODUCTS, p=[0.1, 0.25, 0.35, 0.30])
})

# Customers
signup_dates = pd.to_datetime("2024-01-01") + pd.to_timedelta(np.random.randint(0, 730, N_CUSTOMERS), unit="D")
customers = pd.DataFrame({
    "customer_id": np.arange(50001, 50001+N_CUSTOMERS),
    "customer_name": [f"Customer_{i}" for i in range(N_CUSTOMERS)],
    "city": np.random.choice(cities, N_CUSTOMERS),
    "region": np.random.choice(regions, N_CUSTOMERS),
    "segment": np.random.choice(segments, N_CUSTOMERS, p=[0.55, 0.25, 0.20]),
    "signup_date": signup_dates
})

# Orders
order_dates = pd.to_datetime("2025-01-01") + pd.to_timedelta(np.random.randint(0, 400, N_ORDERS), unit="D")
orders = pd.DataFrame({
    "order_id": np.arange(9000001, 9000001+N_ORDERS),
    "order_date": order_dates,
    "customer_id": np.random.choice(customers["customer_id"], N_ORDERS),
    "product_id": np.random.choice(products["product_id"], N_ORDERS),
    "qty": np.random.randint(1, 8, N_ORDERS),
    "unit_price": np.round(np.random.uniform(80, 6000, N_ORDERS), 2),
    "discount_pct": np.round(np.random.choice([0, 5, 10, 15, 20, 30, 40], N_ORDERS, p=[0.35,0.18,0.16,0.12,0.10,0.06,0.03]), 2),
    "payment_mode": np.random.choice(["UPI","Card","COD","NetBanking"], N_ORDERS, p=[0.42,0.35,0.15,0.08])
})

# Add some dirty data intentionally
# 1) some negative qty
bad_idx = np.random.choice(orders.index, 25, replace=False)
orders.loc[bad_idx, "qty"] = -1
# 2) some weird region strings in customers
customers.loc[np.random.choice(customers.index, 20, replace=False), "region"] = " east "  # messy
# 3) some missing prices
orders.loc[np.random.choice(orders.index, 30, replace=False), "unit_price"] = np.nan

# Returns (~8%)
ret_mask = np.random.rand(N_ORDERS) < 0.08
returns = orders.loc[ret_mask, ["order_id","order_date"]].copy()
returns["return_date"] = returns["order_date"] + pd.to_timedelta(np.random.randint(2, 25, len(returns)), unit="D")
returns["return_reason"] = np.random.choice(
    ["Damaged", "Wrong Item", "Late Delivery", "Size Issue", "Changed Mind"],
    len(returns)
)

# Targets (month-region)
months = pd.date_range("2025-01-01", periods=14, freq="MS")
targets = []
for m in months:
    for r in regions:
        targets.append([m.strftime("%Y-%m"), r, float(np.random.randint(800000, 1800000))])
targets = pd.DataFrame(targets, columns=["month","region","target_sales"])

# Save
with pd.ExcelWriter("input_business_data.xlsx", engine="openpyxl") as w:
    orders.to_excel(w, sheet_name="Orders", index=False)
    products.to_excel(w, sheet_name="Products", index=False)
    customers.to_excel(w, sheet_name="Customers", index=False)
    returns.to_excel(w, sheet_name="Returns", index=False)
    targets.to_excel(w, sheet_name="Targets", index=False)

print("✅ Created: input_business_data.xlsx")

