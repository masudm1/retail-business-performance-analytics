# Retail Business Performance Reporting & Analytics

## Executive Summary
This project analyzes 541,909 online retail transaction records to uncover sales trends, customer purchasing behavior, and product performance using Python, SQL, and Tableau.

The workflow combines data cleaning, exploratory analysis, SQL-based business querying, KPI generation, and dashboard visualization to transform raw transactional data into actionable business insights for reporting and decision-making.

---

## Business Problem
Retail businesses generate large volumes of transactional data, but deriving meaningful insights from this data can be challenging without structured analysis.

This project aims to:
- Identify sales and revenue trends
- Analyze customer purchasing behavior
- Evaluate top-performing products and regions
- Generate business KPIs for reporting
- Support data-driven decision-making through interactive dashboards

---

## Dataset Information
The dataset contains online retail transaction records, including:
- Invoice details
- Product descriptions
- Quantity sold
- Unit price
- Customer ID
- Country
- Transaction date

### Dataset Scale
- 541,909 transaction records
- 4,339 unique customers
- 18,536 completed orders

---

## Tools & Technologies
- Python
- Pandas & NumPy
- SQL (SQLite)
- Matplotlib
- Tableau Public
- Jupyter Notebook

---

## Data Cleaning & Preprocessing
The following preprocessing steps were performed using Python and Pandas:

- Removed duplicate records
- Handled missing customer identifiers
- Converted transaction dates into datetime format
- Removed cancelled transactions
- Created revenue-related calculated fields
- Standardized and validated transactional data

---

## Exploratory Data Analysis (EDA)
Key analytical tasks performed:

- Monthly revenue trend analysis
- Top-selling products analysis
- Revenue contribution by country
- Customer purchasing behavior analysis
- Average order value calculation
- Customer segmentation by lifetime spend
- Revenue concentration analysis

---

## SQL Business Analysis
After preprocessing in Pandas, the cleaned dataset was loaded into a SQLite database for structured business querying.

SQL queries were used to generate:
- KPI summaries
- Monthly sales and order trends
- Product-level revenue analysis
- Country-wise revenue contribution
- Customer lifetime value insights
- Spend-based customer segmentation
- Top customer rankings

### Customer Segmentation Logic
Customers were segmented based on lifetime spending:
- High Value
- Mid Value
- Low Value
- Micro Customers

---

## Key Business Insights
- Total Revenue: **$8.89M**
- Total Orders: **18,536**
- Total Customers: **4,339**
- Average Order Value: **$479**
- UK contributed approximately **82%** of total revenue
- Top 10 products generated approximately **14.4%** of total revenue
- A relatively small group of high-value customers contributed disproportionately to overall sales

---

## Business Recommendations
- Increase inventory planning before peak seasonal demand periods
- Develop retention strategies for high-value customers
- Expand marketing efforts beyond the UK market to reduce regional revenue concentration risk
- Focus promotional strategies on top-performing product categories
- Monitor low-performing products for pricing or inventory optimization opportunities

---

## Dashboard Features
The Tableau dashboard includes:

- KPI Cards
  - Total Revenue
  - Total Orders
  - Total Customers
  - Average Order Value

- Monthly Revenue Trend Visualization
- Top Products by Revenue
- Country-wise Revenue Breakdown
- Interactive Filters for business reporting

---

## Project Structure

```text
retail-sales-customer-analytics/
│
├── data/
├── notebooks/
├── outputs/
│   └── sql_query_results/
├── images/
├── sql_analysis.py
├── requirements.txt
└── README.md
```

---

## Dashboard Preview

![Dashboard Preview](images/dashboard_preview.png)

---

## Author
**Masud Mondol**




