import pandas as pd
import sqlite3
import os

os.makedirs('outputs', exist_ok=True)

df = pd.read_csv('data/cleaned_retail_data.csv')
conn = sqlite3.connect('retail_database.db')
df.to_sql('sales', conn, if_exists='replace', index=False)

def run_query(query, description, filename):
    print(f"\n{description}")
    result = pd.read_sql_query(query, conn)
    print(result.to_string(index=False))
    result.to_csv(f'outputs/{filename}', index=False)
    return result

run_query("""
SELECT
    COUNT(*) AS Total_Clean_Rows,
    COUNT(DISTINCT CustomerID) AS Unique_Customers,
    COUNT(DISTINCT InvoiceNo) AS Unique_Orders,
    ROUND(SUM(Revenue), 2) AS Total_Revenue,
    ROUND(SUM(Revenue) / COUNT(DISTINCT InvoiceNo), 2) AS Avg_Order_Value
FROM sales;
""", "KPI Summary", "01_kpi_summary.csv")

run_query("""
SELECT
    SUBSTR(InvoiceDate, 1, 7) AS Month,
    ROUND(SUM(Revenue), 2) AS Monthly_Revenue,
    COUNT(DISTINCT InvoiceNo) AS Monthly_Orders,
    COUNT(DISTINCT CustomerID) AS Monthly_Customers
FROM sales
GROUP BY Month
ORDER BY Month;
""", "Monthly Revenue Trend", "02_monthly_trend.csv")

run_query("""
SELECT
    Description AS Product,
    SUM(Quantity) AS Units_Sold,
    ROUND(SUM(Revenue), 2) AS Total_Revenue
FROM sales
GROUP BY Description
ORDER BY Total_Revenue DESC
LIMIT 10;
""", "Top 10 Products", "03_top_products.csv")

run_query("""
SELECT
    Description AS Product,
    SUM(Quantity) AS Units_Sold,
    ROUND(SUM(Revenue), 2) AS Total_Revenue
FROM sales
GROUP BY Description
HAVING Total_Revenue > 0
ORDER BY Total_Revenue ASC
LIMIT 10;
""", "Bottom 10 Products", "04_bottom_products.csv")

run_query("""
SELECT
    Country,
    COUNT(DISTINCT CustomerID) AS Customers,
    COUNT(DISTINCT InvoiceNo) AS Orders,
    ROUND(SUM(Revenue), 2) AS Revenue,
    ROUND(SUM(Revenue) * 100.0 / (SELECT SUM(Revenue) FROM sales), 1) AS Revenue_Pct
FROM sales
GROUP BY Country
ORDER BY Revenue DESC;
""", "Revenue by Country", "05_revenue_by_country.csv")

run_query("""
SELECT
    Segment,
    COUNT(*) AS Customers,
    ROUND(AVG(Total_Spent), 2) AS Avg_Spend,
    ROUND(SUM(Total_Spent), 2) AS Segment_Revenue
FROM (
    SELECT
        CustomerID,
        ROUND(SUM(Revenue), 2) AS Total_Spent,
        CASE
            WHEN SUM(Revenue) > 5000 THEN 'High Value (>$5000)'
            WHEN SUM(Revenue) > 1000 THEN 'Mid Value ($1000-$5000)'
            WHEN SUM(Revenue) > 100  THEN 'Low Value ($100-$1000)'
            ELSE 'Micro (<$100)'
        END AS Segment
    FROM sales
    GROUP BY CustomerID
)
GROUP BY Segment
ORDER BY Segment_Revenue DESC;
""", "Customer Segmentation", "06_customer_segments.csv")

run_query("""
SELECT
    CustomerID,
    COUNT(DISTINCT InvoiceNo) AS Total_Orders,
    ROUND(SUM(Revenue), 2) AS Total_Spent,
    ROUND(SUM(Revenue) / COUNT(DISTINCT InvoiceNo), 2) AS Avg_Order_Value
FROM sales
GROUP BY CustomerID
ORDER BY Total_Spent DESC
LIMIT 10;
""", "Top 10 Customers", "07_top_customers.csv")

conn.close()
print("\nDone. Results saved to outputs/")