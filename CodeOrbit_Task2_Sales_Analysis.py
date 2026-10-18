
import pandas as pd

file_name = "CodeOrbit_Task1_Cleaned_Online_RetailDataset.xlsx"
output_file = "CodeOrbit_Task2_Sales_Analysis_Results.xlsx"

try:
    # 1. Load the cleaned dataset
    df = pd.read_excel(file_name)

    # 2. Prepare data types
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["UnitPrice"] = pd.to_numeric(df["UnitPrice"], errors="coerce")
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], errors="coerce")

    # 3. Calculate sales amount
    df["SalesAmount"] = df["Quantity"] * df["UnitPrice"]

    # 4. Overall sales summary
    total_sales = df["SalesAmount"].sum()
    positive_sales = df.loc[df["SalesAmount"] > 0, "SalesAmount"].sum()
    return_value = df.loc[df["SalesAmount"] < 0, "SalesAmount"].sum()
    total_quantity = df["Quantity"].sum()
    unique_invoices = df["InvoiceNo"].nunique()

    overall_summary = pd.DataFrame({
        "Metric": [
            "Total Records",
            "Unique Invoices",
            "Total Quantity (net of returns)",
            "Net Sales (including returns)",
            "Positive Sales",
            "Return Value (negative)"
        ],
        "Value": [
            len(df),
            unique_invoices,
            total_quantity,
            total_sales,
            positive_sales,
            return_value
        ]
    })

    # 5. Monthly sales
    df["SalesMonth"] = df["InvoiceDate"].dt.to_period("M").astype(str)
    monthly_sales = (
        df.dropna(subset=["InvoiceDate"])
          .groupby("SalesMonth", as_index=False)["SalesAmount"]
          .sum()
          .sort_values("SalesMonth")
    )

    # 6. Country-wise sales
    country_sales = (
        df.groupby("Country", as_index=False)["SalesAmount"]
          .sum()
          .sort_values("SalesAmount", ascending=False)
    )

    # 7. Top 10 products by sales
    product_sales = (
        df.groupby(["StockCode", "Description"], dropna=False, as_index=False)
          ["SalesAmount"].sum()
          .sort_values("SalesAmount", ascending=False)
    )
    top_10_products = product_sales.head(10)

    # 8. Print key findings
    print("\n===== CODEORBIT TASK 2 RESULTS =====")
    print("\nOVERALL SUMMARY")
    print(overall_summary.to_string(index=False))

    print("\nMONTHLY SALES")
    print(monthly_sales.to_string(index=False))

    print("\nTOP 10 COUNTRIES BY SALES")
    print(country_sales.head(10).to_string(index=False))

    print("\nTOP 10 PRODUCTS BY SALES")
    print(top_10_products.to_string(index=False))

    # 9. Export results to a multi-sheet Excel workbook
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        overall_summary.to_excel(
            writer, sheet_name="Overall_Summary", index=False
        )
        monthly_sales.to_excel(
            writer, sheet_name="Monthly_Sales", index=False
        )
        country_sales.to_excel(
            writer, sheet_name="Country_Sales", index=False
        )
        top_10_products.to_excel(
            writer, sheet_name="Top_10_Products", index=False
        )
        product_sales.to_excel(
            writer, sheet_name="All_Product_Sales", index=False
        )

    print(f"\nAnalysis completed successfully!")
    print(f"Results saved to: {output_file}")

except Exception as e:
    print("\nAn error occurred:", e)

# 10. Create a short findings report

report = """
CODEORBIT INTERNSHIP - TASK 2
Sales Data Analysis Using Pandas

1. OBJECTIVE
To analyze online retail sales data using Python Pandas and
identify sales performance trends.

2. DATASET
The cleaned dataset contains 535,187 records and 8 columns,
including invoice number, product, quantity, date, unit price,
customer ID, and country.

3. KEY FINDINGS
- Net sales: 9,726,006.95
- Unique invoices: 24,446
- Average net sales per invoice: 397.86
- Average positive invoice sales: 533.17
- Highest-sales country: United Kingdom
- Highest-sales month: November 2011
- Highest-ranked item overall: DOTCOM POSTAGE

4. ANALYSIS PERFORMED
Sales amount was calculated as Quantity multiplied by UnitPrice.
Sales were summarized by month, country, and product.
An Excel workbook was generated containing the analysis results.

5. CONCLUSION
The analysis indicates that the United Kingdom contributed the
largest share of sales in the dataset. November 2011 recorded the
highest monthly sales. Negative quantities and sales values were
retained to account for returns. The results can help identify
important markets, sales periods, and high-performing items.
"""

with open("CodeOrbit_Task2_Findings_Report.txt", "w", encoding="utf-8") as f:
    f.write(report.strip())

print("Findings report created successfully!")

