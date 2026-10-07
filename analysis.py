import pandas as pd
import matplotlib.pyplot as plt


# Load data
df = pd.read_excel("online_retail_II.xlsx")


# Data cleaning
df = df.drop_duplicates()

df["Cancelled"] = (
    df["Invoice"]
    .astype(str)
    .str.strip()
    .str.startswith("C")
)

sales = df[
    (~df["Cancelled"]) &
    (df["Quantity"] > 0) &
    (df["Price"] > 0)
].copy()


# Sales calculations
sales["Revenue"] = sales["Quantity"] * sales["Price"]
sales["Month"] = sales["InvoiceDate"].dt.to_period("M").astype(str)


# KPIs
total_revenue = sales["Revenue"].sum()
total_orders = sales["Invoice"].nunique()
total_customers = sales["Customer ID"].nunique()
total_quantity = sales["Quantity"].sum()
average_order_value = total_revenue / total_orders


# Monthly sales analysis
monthly_sales = (
    sales.groupby("Month")
    .agg(
        Revenue=("Revenue", "sum"),
        Orders=("Invoice", "nunique"),
        Quantity=("Quantity", "sum")
    )
    .sort_index()
)


# Product analysis
product_sales = (
    sales.groupby("Description")
    .agg(
        Revenue=("Revenue", "sum"),
        Quantity=("Quantity", "sum"),
        Orders=("Invoice", "nunique")
    )
    .sort_values("Revenue", ascending=False)
)

top_products = product_sales.head(20)


# Country analysis
country_sales = (
    sales.groupby("Country")
    .agg(
        Revenue=("Revenue", "sum"),
        Orders=("Invoice", "nunique"),
        Customers=("Customer ID", "nunique")
    )
    .sort_values("Revenue", ascending=False)
)


# Customer analysis
customer_sales = (
    sales.dropna(subset=["Customer ID"])
    .groupby("Customer ID")
    .agg(
        Revenue=("Revenue", "sum"),
        Orders=("Invoice", "nunique"),
        Quantity=("Quantity", "sum"),
        LastPurchase=("InvoiceDate", "max")
    )
    .sort_values("Revenue", ascending=False)
)


# RFM analysis
reference_date = sales["InvoiceDate"].max() + pd.Timedelta(days=1)

rfm = (
    sales.dropna(subset=["Customer ID"])
    .groupby("Customer ID")
    .agg(
        Recency=("InvoiceDate", lambda x: (reference_date - x.max()).days),
        Frequency=("Invoice", "nunique"),
        Monetary=("Revenue", "sum")
    )
)


rfm["R_Score"] = (
    pd.qcut(
        rfm["Recency"].rank(method="first"),
        4,
        labels=[4, 3, 2, 1]
    )
    .astype(int)
)

rfm["F_Score"] = (
    pd.qcut(
        rfm["Frequency"].rank(method="first"),
        4,
        labels=[1, 2, 3, 4]
    )
    .astype(int)
)

rfm["M_Score"] = (
    pd.qcut(
        rfm["Monetary"].rank(method="first"),
        4,
        labels=[1, 2, 3, 4]
    )
    .astype(int)
)

rfm["RFM_Score"] = (
    rfm["R_Score"].astype(str) +
    rfm["F_Score"].astype(str) +
    rfm["M_Score"].astype(str)
)


# Cancellation analysis
cancelled = df[df["Cancelled"]].copy()

cancelled["Cancelled_Value"] = (
    cancelled["Quantity"].abs() *
    cancelled["Price"].abs()
)

cancellation_summary = pd.DataFrame({
    "Cancelled_Invoices": [cancelled["Invoice"].nunique()],
    "Cancelled_Items": [cancelled["Quantity"].abs().sum()],
    "Cancelled_Value": [cancelled["Cancelled_Value"].sum()]
})


# Export analysis to Excel
with pd.ExcelWriter(
    "online_retail_analysis.xlsx",
    engine="openpyxl"
) as writer:

    monthly_sales.to_excel(
        writer,
        sheet_name="Monthly Sales"
    )

    top_products.to_excel(
        writer,
        sheet_name="Top Products"
    )

    country_sales.to_excel(
        writer,
        sheet_name="Country Sales"
    )

    customer_sales.to_excel(
        writer,
        sheet_name="Customer Sales"
    )

    rfm.to_excel(
        writer,
        sheet_name="RFM Analysis"
    )

    cancellation_summary.to_excel(
        writer,
        sheet_name="Cancellations",
        index=False
    )

    for worksheet in writer.book.worksheets:
        for column in worksheet.columns:
            max_length = max(
                len(str(cell.value))
                if cell.value is not None
                else 0
                for cell in list(column)
            )

            worksheet.column_dimensions[
                column[0].column_letter].width = max_length + 2


# Monthly revenue chart
plt.figure(figsize=(12, 6))

plt.plot(
    monthly_sales.index,
    monthly_sales["Revenue"]
)

plt.title("Monthly Revenue")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.xticks(rotation=45)
plt.tight_layout()

plt.savefig("monthly_revenue.png")
plt.close()


# Top products chart
plt.figure(figsize=(10, 6))

top_products.head(10)["Revenue"].sort_values().plot(
    kind="barh"
)

plt.title("Top 10 Products by Revenue")
plt.xlabel("Revenue")
plt.tight_layout()

plt.savefig("top_products.png")
plt.close()


# Top countries chart
plt.figure(figsize=(10, 6))

country_sales.head(10)["Revenue"].sort_values().plot(
    kind="barh"
)

plt.title("Top 10 Countries by Revenue")
plt.xlabel("Revenue")
plt.tight_layout()

plt.savefig("top_countries.png")
plt.close()


# Final dashboard
fig = plt.figure(figsize=(18, 11))

fig.suptitle(
    "Online Retail Sales Analytics Dashboard",
    fontsize=22,
    fontweight="bold"
)


# KPI cards
kpis = [
    ("Total Revenue", f"{total_revenue:,.0f}"),
    ("Total Orders", f"{total_orders:,}"),
    ("Total Customers", f"{total_customers:,}"),
    ("Average Order Value", f"{average_order_value:,.2f}")
]

positions = [
    (0.08, 0.82),
    (0.30, 0.82),
    (0.52, 0.82),
    (0.74, 0.82)
]

for (title, value), (x, y) in zip(kpis, positions):

    fig.text(
        x,
        y,
        title,
        ha="center",
        fontsize=12
    )

    fig.text(
        x,
        y - 0.045,
        value,
        ha="center",
        fontsize=20,
        fontweight="bold"
    )


# Monthly revenue
ax1 = fig.add_axes([0.08, 0.47, 0.84, 0.25])

ax1.plot(
    monthly_sales.index,
    monthly_sales["Revenue"]
)

ax1.set_title("Monthly Revenue", fontsize=14)
ax1.set_xlabel("Month")
ax1.set_ylabel("Revenue")
ax1.tick_params(axis="x", rotation=45)


# Top products
ax2 = fig.add_axes([0.08, 0.08, 0.38, 0.28])

top_products.head(10)["Revenue"].sort_values().plot(
    kind="barh",
    ax=ax2
)

ax2.set_title("Top 10 Products by Revenue", fontsize=14)
ax2.set_xlabel("Revenue")


# Top countries
ax3 = fig.add_axes([0.55, 0.08, 0.38, 0.28])

country_sales.head(10)["Revenue"].sort_values().plot(
    kind="barh",
    ax=ax3
)

ax3.set_title("Top 10 Countries by Revenue", fontsize=14)
ax3.set_xlabel("Revenue")


plt.savefig(
    "online_retail_dashboard.png",
    dpi=200,
    bbox_inches="tight"
)

plt.close()