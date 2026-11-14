"""
SQL risk & exception queries (SQLite) - companion to risk_exception_analysis.py
Runs directly on the RAW file so every figure can be reconciled to the source.
Run:  python sql_risk_queries.py     (writes outputs/risk/sql/*.csv)
"""
import os
import sqlite3

import pandas as pd

os.makedirs("outputs/risk/sql", exist_ok=True)

raw = pd.read_csv("data/online_retail.csv", encoding="latin1")
raw["Revenue"] = raw["Quantity"] * raw["UnitPrice"]
raw["Month"] = raw["InvoiceDate"].str[:7]
conn = sqlite3.connect(":memory:")
raw.to_sql("raw_sales", conn, index=False)
# De-duplicated view of the same data (exact duplicate lines removed) used by queries 3-7
raw.drop_duplicates(subset=[c for c in raw.columns if c not in ("Revenue", "Month")]).to_sql("sales", conn, index=False)


def run(title, filename, query):
    df = pd.read_sql_query(query, conn)
    print(f"\n{title}\n{df.head(12).to_string(index=False)}")
    df.to_csv(f"outputs/risk/sql/{filename}", index=False)


# 1. Row-count reconciliation by exception type -------------------------------------------------
run("1. Data-quality scorecard", "q1_data_quality_scorecard.csv", """
SELECT 'Total rows' AS Check_Name, COUNT(*) AS Lines, ROUND(SUM(Revenue),2) AS Net_Value FROM raw_sales
UNION ALL SELECT 'Missing CustomerID', COUNT(*), ROUND(SUM(Revenue),2) FROM raw_sales WHERE CustomerID IS NULL
UNION ALL SELECT 'Missing Description', COUNT(*), ROUND(SUM(Revenue),2) FROM raw_sales WHERE Description IS NULL
UNION ALL SELECT 'Cancelled invoices (C-prefix)', COUNT(*), ROUND(SUM(Revenue),2) FROM raw_sales WHERE InvoiceNo LIKE 'C%'
UNION ALL SELECT 'Zero unit price', COUNT(*), ROUND(SUM(Revenue),2) FROM raw_sales WHERE UnitPrice = 0
UNION ALL SELECT 'Negative unit price', COUNT(*), ROUND(SUM(Revenue),2) FROM raw_sales WHERE UnitPrice < 0
UNION ALL SELECT 'Negative qty, not a cancellation', COUNT(*), ROUND(SUM(Revenue),2)
       FROM raw_sales WHERE Quantity < 0 AND InvoiceNo NOT LIKE 'C%'
UNION ALL SELECT 'Quantity >= 5,000 units on one line', COUNT(*), ROUND(SUM(Revenue),2)
       FROM raw_sales WHERE Quantity >= 5000;
""")

# 2. Exact duplicate lines ------------------------------------------------------------------------
run("2. Exact duplicate lines by month", "q2_duplicates_by_month.csv", """
SELECT Month, SUM(cnt - 1) AS Duplicate_Lines
FROM (SELECT Month, COUNT(*) AS cnt FROM raw_sales
      GROUP BY InvoiceNo, StockCode, Description, Quantity, InvoiceDate, UnitPrice, CustomerID, Country)
WHERE cnt > 1
GROUP BY Month ORDER BY Month;
""")

# 3. Cancellation rate by month (product SKUs only) -----------------------------------------------
run("3. Cancellation rate by month", "q3_cancellation_rate_by_month.csv", """
SELECT Month,
       ROUND(SUM(CASE WHEN InvoiceNo NOT LIKE 'C%' AND Quantity > 0 THEN Revenue END), 2) AS Gross_Sales,
       ROUND(-SUM(CASE WHEN InvoiceNo LIKE 'C%' THEN Revenue END), 2) AS Cancelled_Value,
       ROUND(-SUM(CASE WHEN InvoiceNo LIKE 'C%' THEN Revenue END) * 100.0 /
              SUM(CASE WHEN InvoiceNo NOT LIKE 'C%' AND Quantity > 0 THEN Revenue END), 2) AS Cancel_Rate_Pct
FROM sales
WHERE StockCode NOT IN ('POST','D','M','DOT','BANK CHARGES','AMAZONFEE','CRUK','S','B','C2','PADS','m')
  AND InvoiceNo NOT LIKE 'A%'
GROUP BY Month ORDER BY Month;
""")

# 4. Price consistency per product (same item sold at very different prices) ----------------------
run("4. Products with the widest price spread", "q4_price_spread_by_product.csv", """
SELECT StockCode, MIN(Description) AS Description, COUNT(*) AS Lines,
       ROUND(MIN(UnitPrice), 2) AS Min_Price, ROUND(MAX(UnitPrice), 2) AS Max_Price,
       ROUND(MAX(UnitPrice) / MIN(UnitPrice), 1) AS Max_to_Min_Ratio
FROM sales
WHERE UnitPrice > 0 AND Quantity > 0 AND InvoiceNo NOT LIKE 'C%' AND InvoiceNo NOT LIKE 'A%'
  AND StockCode NOT IN ('POST','D','M','DOT','BANK CHARGES','AMAZONFEE','CRUK','S','B','C2','PADS','m')
GROUP BY StockCode
HAVING COUNT(*) >= 20 AND MAX(UnitPrice) / MIN(UnitPrice) >= 10
ORDER BY Max_to_Min_Ratio DESC
LIMIT 25;
""")

# 5. Customers whose cancellations are large relative to purchases ---------------------------------
run("5. Customer cancellation exposure", "q5_customer_cancellation_exposure.csv", """
WITH s AS (
  SELECT CustomerID,
         SUM(CASE WHEN InvoiceNo NOT LIKE 'C%' AND Quantity > 0 THEN Revenue ELSE 0 END) AS Gross_Sales,
         -SUM(CASE WHEN InvoiceNo LIKE 'C%' THEN Revenue ELSE 0 END) AS Cancelled_Value,
         COUNT(DISTINCT CASE WHEN InvoiceNo LIKE 'C%' THEN InvoiceNo END) AS Cancel_Invoices
  FROM sales
  WHERE CustomerID IS NOT NULL
    AND StockCode NOT IN ('POST','D','M','DOT','BANK CHARGES','AMAZONFEE','CRUK','S','B','C2','PADS','m')
    AND InvoiceNo NOT LIKE 'A%'
  GROUP BY CustomerID)
SELECT CAST(CustomerID AS INTEGER) AS CustomerID, ROUND(Gross_Sales, 2) AS Gross_Sales,
       ROUND(Cancelled_Value, 2) AS Cancelled_Value, Cancel_Invoices,
       ROUND(MIN(Cancelled_Value * 1.0 / NULLIF(Gross_Sales, 0), 1.0) * 100, 1) AS Cancel_Rate_Pct,
       CASE WHEN Cancelled_Value >= 1000 AND Cancelled_Value >= 0.30 * Gross_Sales THEN 'High'
            WHEN Cancelled_Value >= 500  AND Cancelled_Value >= 0.15 * Gross_Sales THEN 'Medium'
            ELSE 'Low' END AS Risk_Tier
FROM s
WHERE Cancelled_Value >= 500
ORDER BY Cancelled_Value DESC
LIMIT 25;
""")

# 6. Revenue concentration (window functions) -----------------------------------------------------
run("6. Country concentration with running share", "q6_country_concentration.csv", """
WITH c AS (
  SELECT Country, SUM(Revenue) AS Revenue
  FROM sales
  WHERE CustomerID IS NOT NULL AND InvoiceNo NOT LIKE 'C%'
    AND StockCode NOT IN ('POST','D','M','DOT','BANK CHARGES','AMAZONFEE','CRUK','S','B','C2','PADS','m')
  GROUP BY Country)
SELECT Country, ROUND(Revenue, 2) AS Revenue,
       ROUND(Revenue * 100.0 / SUM(Revenue) OVER (), 2) AS Share_Pct,
       ROUND(SUM(Revenue) OVER (ORDER BY Revenue DESC) * 100.0 / SUM(Revenue) OVER (), 2) AS Cumulative_Share_Pct
FROM c ORDER BY Revenue DESC;
""")

# 7. Stock write-off style adjustments by description ----------------------------------------------
run("7. Unexplained negative-quantity lines (write-offs / damages)", "q7_negative_qty_notes.csv", """
SELECT Description, COUNT(*) AS Lines, SUM(Quantity) AS Units_Adjusted
FROM sales
WHERE Quantity < 0 AND InvoiceNo NOT LIKE 'C%' AND InvoiceNo NOT LIKE 'A%' AND Description IS NOT NULL
GROUP BY Description ORDER BY Lines DESC LIMIT 15;
""")
conn.close()
print("\nDone. Results saved to outputs/risk/sql/")
