"""
Retail Risk & Exception Monitoring
==================================
Extension of the Retail Business Performance project.

Starts from the RAW transaction file and answers the questions a risk /
MIS team asks of operational data:

  1. Reconciliation  - does every row and dollar between raw and reported figures tie out?
  2. Exceptions      - which transactions break business rules (duplicates, zero or negative
                       prices, unexplained negative quantities, price / quantity outliers,
                       non-product codes such as fees and manual adjustments)?
  3. Cancellations   - where are returns / cancellations concentrated (month, country,
                       product, customer)?
  4. Concentration   - how dependent is revenue on one country, a few products or a few customers?
  5. Anomalies       - which trading days look abnormal against recent history?
  6. Watchlists      - ranked customers and products for follow-up.

Run:   python risk_exception_analysis.py
Needs: data/online_retail.csv  (already in the repo)
Writes: outputs/risk/*.csv, images/risk_*.png, excel/Retail_Risk_Exception_Report.xlsx
"""

import os
import re

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

RAW_PATH = "data/online_retail.csv"
OUT_DIR = "outputs/risk"
IMG_DIR = "images"
XLSX_PATH = "excel/Retail_Risk_Exception_Report.xlsx"
for d in (OUT_DIR, IMG_DIR, "excel"):
    os.makedirs(d, exist_ok=True)

# ----------------------------------------------------------------------------
# Business-rule parameters (kept in one place so a reviewer can challenge them)
# ----------------------------------------------------------------------------
NON_PRODUCT_CODES = {
    "POST", "D", "M", "DOT", "BANK CHARGES", "AMAZONFEE", "CRUK",
    "S", "B", "C2", "PADS", "m",
}
PRICE_OUTLIER_MULT = 5          # unit price >= 5x or <= 1/5 of the product's median price
QTY_OUTLIER_MULT = 20           # quantity >= 20x the product's median quantity ...
QTY_OUTLIER_FLOOR = 500         # ... and at least this many units
QTY_ABSOLUTE_LIMIT = 5000       # any single line this large is flagged even without a benchmark
REVERSED_ORDER_MIN_VALUE = 5000 # report cancellations >= this value that exactly mirror an earlier sale
MIN_LINES_FOR_BENCHMARK = 20    # product needs this many sales lines to be benchmarked
HIGH_VALUE_THRESHOLD = 5000     # USD, escalates a Medium exception to High
CANCEL_RATE_HIGH, CANCEL_VALUE_HIGH = 0.30, 1000     # customer watchlist: High tier
CANCEL_RATE_MED, CANCEL_VALUE_MED = 0.15, 500        # customer watchlist: Medium tier
PRODUCT_MIN_SALES_VALUE = 1000  # product watchlist ignores tiny products
PRODUCT_CANCEL_RATE_FLAG = 0.15

NAVY, TEAL, AMBER, RED, GREY = "#1F3A5F", "#2A9D8F", "#E9A23B", "#C0392B", "#8A8F98"


def money(x):
    return f"${x:,.0f}"


# ----------------------------------------------------------------------------
# 1. Load raw data
# ----------------------------------------------------------------------------
raw = pd.read_csv(RAW_PATH, encoding="latin1", parse_dates=["InvoiceDate"])
raw["InvoiceNo"] = raw["InvoiceNo"].astype(str)
raw["StockCode"] = raw["StockCode"].astype(str)
raw["Revenue"] = raw["Quantity"] * raw["UnitPrice"]
raw["Month"] = raw["InvoiceDate"].dt.strftime("%Y-%m")
raw["IsCancel"] = raw["InvoiceNo"].str.startswith("C")
raw["IsDuplicate"] = raw.duplicated(keep="first")
raw["RowID"] = np.arange(2, len(raw) + 2)  # row number in the CSV file (header = row 1)

dedup = raw[~raw["IsDuplicate"]].copy()
is_service = dedup["StockCode"].isin(NON_PRODUCT_CODES) | dedup["InvoiceNo"].str.startswith("A")
products = dedup[~is_service].copy()   # analysis population for cancellations / concentration

# ----------------------------------------------------------------------------
# 2. Reconciliation bridge: raw file -> the cleaned dataset used in the README
# ----------------------------------------------------------------------------
n_raw, v_raw = len(raw), raw["Revenue"].sum()
step1 = raw[~raw["IsDuplicate"]]
step2 = step1[step1["CustomerID"].notna()]
step3 = step2[~step2["IsCancel"]]

bridge = pd.DataFrame(
    [
        ["Raw file (all rows)", n_raw, v_raw],
        ["Less: exact duplicate lines", -(n_raw - len(step1)), -(v_raw - step1["Revenue"].sum())],
        ["Less: lines with no CustomerID", -(len(step1) - len(step2)), -(step1["Revenue"].sum() - step2["Revenue"].sum())],
        ["Less: cancelled invoices (C-prefix)", -(len(step2) - len(step3)), -(step2["Revenue"].sum() - step3["Revenue"].sum())],
        ["Cleaned dataset used for reported KPIs", len(step3), step3["Revenue"].sum()],
    ],
    columns=["Step", "Rows", "Net_Value_USD"],
)
assert bridge["Rows"].iloc[:-1].sum() == bridge["Rows"].iloc[-1], "row bridge does not tie out"
assert abs(bridge["Net_Value_USD"].iloc[:-1].sum() - bridge["Net_Value_USD"].iloc[-1]) < 0.01, "value bridge does not tie out"

cancelled_identified_value = -step2[step2["IsCancel"]]["Revenue"].sum()
gross_identified_value = step3["Revenue"].sum()
no_cust_value = step1[step1["CustomerID"].isna()]["Revenue"].sum()

# ----------------------------------------------------------------------------
# 3. Exception rules (row level)
# ----------------------------------------------------------------------------
# Benchmarks per product, from genuine sales lines only
sales_lines = products[(~products["IsCancel"]) & (products["Quantity"] > 0) & (products["UnitPrice"] > 0)]
bench = sales_lines.groupby("StockCode").agg(
    Lines=("Quantity", "size"),
    Median_Price=("UnitPrice", "median"),
    Median_Qty=("Quantity", "median"),
)
bench = bench[bench["Lines"] >= MIN_LINES_FOR_BENCHMARK]

d = dedup.merge(bench, left_on="StockCode", right_index=True, how="left")
d["Est_Value_At_Risk"] = np.where(
    d["Revenue"] != 0, d["Revenue"].abs(), d["Quantity"].abs() * d["Median_Price"].fillna(0)
)

is_service_d = d["StockCode"].isin(NON_PRODUCT_CODES) | d["InvoiceNo"].str.startswith("A")
normal_sale = (~d["IsCancel"]) & (d["Quantity"] > 0) & (d["UnitPrice"] > 0) & (~is_service_d)

rules = {
    "BAD_DEBT_ADJUSTMENT":     (d["InvoiceNo"].str.startswith("A"),
                                "Manual 'Adjust bad debt' invoice", "High"),
    "NEGATIVE_UNIT_PRICE":     (d["UnitPrice"] < 0,
                                "Unit price below zero", "High"),
    "UNEXPLAINED_NEG_QTY":     ((d["Quantity"] < 0) & (~d["IsCancel"]) & (~d["InvoiceNo"].str.startswith("A")),
                                "Negative quantity on a non-cancellation invoice (write-off / damage / stock fix)", "Medium"),
    "ZERO_PRICE_SALE":         ((d["UnitPrice"] == 0) & (d["Quantity"] > 0) & (~d["IsCancel"]),
                                "Goods issued at zero price", "Low"),
    "NON_PRODUCT_CODE":        (d["StockCode"].isin(NON_PRODUCT_CODES) & (d["UnitPrice"] != 0),
                                "Fees, postage, manual or sample codes mixed into product data", "Low"),
    "PRICE_OUTLIER":           (normal_sale & ((d["UnitPrice"] >= PRICE_OUTLIER_MULT * d["Median_Price"]) |
                                               (d["UnitPrice"] <= d["Median_Price"] / PRICE_OUTLIER_MULT)),
                                f"Unit price >= {PRICE_OUTLIER_MULT}x or <= 1/{PRICE_OUTLIER_MULT} of product median", "Medium"),
    "QUANTITY_OUTLIER":        (normal_sale & (((d["Quantity"] >= QTY_OUTLIER_MULT * d["Median_Qty"]) &
                                                 (d["Quantity"] >= QTY_OUTLIER_FLOOR)) |
                                                (d["Quantity"] >= QTY_ABSOLUTE_LIMIT)),
                                f"Quantity >= {QTY_OUTLIER_MULT}x product median (and >= {QTY_OUTLIER_FLOOR} units), or any single line >= {QTY_ABSOLUTE_LIMIT:,} units", "Medium"),
}

frames = []
for name, (mask, desc, sev) in rules.items():
    f = d[mask].copy()
    f["Rule"], f["Rule_Description"], f["Severity"] = name, desc, sev
    frames.append(f)
log = pd.concat(frames, ignore_index=True)
log = log.drop_duplicates(subset="RowID", keep="first")   # one line per transaction; rules are listed in priority order

# Escalate Medium -> High for material amounts
log.loc[(log["Severity"] == "Medium") & (log["Est_Value_At_Risk"] >= HIGH_VALUE_THRESHOLD), "Severity"] = "High"
log["Severity"] = pd.Categorical(log["Severity"], ["High", "Medium", "Low"], ordered=True)
log = log.sort_values(["Severity", "Est_Value_At_Risk"], ascending=[True, False]).reset_index(drop=True)
log["CustomerID"] = log["CustomerID"].astype("Int64")

exception_log = log[
    ["Severity", "Rule", "InvoiceNo", "InvoiceDate", "StockCode", "Description", "Quantity", "UnitPrice",
     "Revenue", "Est_Value_At_Risk", "CustomerID", "Country", "Median_Price", "Median_Qty", "RowID"]
].copy()
exception_log["Month"] = exception_log["InvoiceDate"].dt.strftime("%Y-%m")
exception_log["InvoiceDate"] = exception_log["InvoiceDate"].dt.strftime("%Y-%m-%d %H:%M")
exception_log["Rule"] = exception_log["Rule"].astype(str)
exception_log["Severity"] = exception_log["Severity"].astype(str)

# Rule summary - include the two data-quality rules that are summarised, not logged line by line
dup_rows = raw[raw["IsDuplicate"]]
nocust_rows = dedup[dedup["CustomerID"].isna()]
rule_summary = (
    exception_log.groupby(["Rule", "Severity"])
    .agg(Lines=("Rule", "size"), Est_Value_At_Risk=("Est_Value_At_Risk", "sum"))
    .reset_index()
)
rule_summary = pd.concat(
    [
        rule_summary,
        pd.DataFrame(
            [
                ["EXACT_DUPLICATE_LINE", "Medium", len(dup_rows), dup_rows["Revenue"].abs().sum()],
                ["MISSING_CUSTOMER_ID", "Medium", len(nocust_rows), nocust_rows["Revenue"].abs().sum()],
            ],
            columns=rule_summary.columns,
        ),
    ],
    ignore_index=True,
).sort_values("Est_Value_At_Risk", ascending=False)
rule_summary["Description"] = rule_summary["Rule"].map(
    {k: v[1] for k, v in rules.items()}
    | {"EXACT_DUPLICATE_LINE": "Identical row repeated in the raw file",
       "MISSING_CUSTOMER_ID": "Transaction cannot be attributed to a customer"}
)
rule_summary["Logged_In_Exception_Log"] = np.where(
    rule_summary["Rule"].isin(["EXACT_DUPLICATE_LINE", "MISSING_CUSTOMER_ID"]), "Summary only", "Yes"
)

# Duplicates are summarised above; also keep a small sample-free count by month for trend
dup_by_month = dup_rows.groupby("Month").size().rename("Duplicate_Lines")

# ----------------------------------------------------------------------------
# 4. Cancellations / returns
# ----------------------------------------------------------------------------
gross = products[(~products["IsCancel"]) & (products["Quantity"] > 0)]
canc = products[products["IsCancel"]]


def cancel_table(by):
    g = gross.groupby(by).agg(Gross_Sales=("Revenue", "sum"), Sales_Lines=("Revenue", "size"),
                              Sales_Invoices=("InvoiceNo", "nunique"))
    c = canc.groupby(by).agg(Cancelled_Value=("Revenue", lambda s: -s.sum()), Cancel_Lines=("Revenue", "size"),
                             Cancel_Invoices=("InvoiceNo", "nunique"))
    t = g.join(c, how="outer").fillna(0)
    t["Cancel_Rate_Value"] = np.where(t["Gross_Sales"] > 0, t["Cancelled_Value"] / t["Gross_Sales"], np.nan)
    return t.reset_index()


canc_month = cancel_table("Month")
canc_month["Note"] = np.where(canc_month["Month"] == "2011-12", "Partial month (data ends 9 Dec 2011)", "")
canc_country = cancel_table("Country")
canc_country = canc_country[canc_country["Gross_Sales"] >= 5000].sort_values("Cancelled_Value", ascending=False)
canc_product = cancel_table(["StockCode", "Description"])

# ----------------------------------------------------------------------------
# 5. Concentration risk (cleaned population = reported KPIs)
# ----------------------------------------------------------------------------
clean = step3.copy()
clean = clean[~clean["StockCode"].isin(NON_PRODUCT_CODES)]


def concentration(series_rev, label):
    s = series_rev.sort_values(ascending=False)
    share = s / s.sum()
    pareto_n = int((share.cumsum() < 0.80).sum() + 1)
    return {
        "Dimension": label,
        "Entities": len(s),
        "Top1_Share_%": round(share.iloc[0] * 100, 1),
        "Top5_Share_%": round(share.iloc[:5].sum() * 100, 1),
        "Top10_Share_%": round(share.iloc[:10].sum() * 100, 1),
        "HHI_(0-10000)": round((share ** 2).sum() * 10000),
        "Entities_For_80%_Revenue": pareto_n,
        "Entities_For_80%_as_%_of_All": round(pareto_n / len(s) * 100, 1),
    }


country_rev = clean.groupby("Country")["Revenue"].sum()
product_rev = clean.groupby("StockCode")["Revenue"].sum()
cust_rev = clean.groupby("CustomerID")["Revenue"].sum()
conc = pd.DataFrame([
    concentration(country_rev, "Country"),
    concentration(product_rev, "Product (StockCode)"),
    concentration(cust_rev, "Customer"),
])
uk_ex = country_rev.drop("United Kingdom")
conc_note = {
    "UK_share_%": round(country_rev["United Kingdom"] / country_rev.sum() * 100, 1),
    "Largest_non_UK_country": uk_ex.idxmax(),
    "Largest_non_UK_share_%": round(uk_ex.max() / country_rev.sum() * 100, 1),
}

pareto_cust = cust_rev.sort_values(ascending=False).reset_index()
pareto_cust["Cum_Revenue_%"] = pareto_cust["Revenue"].cumsum() / pareto_cust["Revenue"].sum() * 100
pareto_cust["Cum_Customers_%"] = (np.arange(len(pareto_cust)) + 1) / len(pareto_cust) * 100

# ----------------------------------------------------------------------------
# 6. Customer and product watchlists
# ----------------------------------------------------------------------------
cust_g = products[products["CustomerID"].notna() & (~products["IsCancel"]) & (products["Quantity"] > 0)]
cust_c = products[products["CustomerID"].notna() & products["IsCancel"]]
cw = cust_g.groupby("CustomerID").agg(Country=("Country", "first"), Gross_Sales=("Revenue", "sum"),
                                      Sales_Invoices=("InvoiceNo", "nunique"))
cw = cw.join(cust_c.groupby("CustomerID").agg(Cancelled_Value=("Revenue", lambda s: -s.sum()),
                                              Cancel_Invoices=("InvoiceNo", "nunique")), how="outer").fillna(
    {"Gross_Sales": 0, "Sales_Invoices": 0, "Cancelled_Value": 0, "Cancel_Invoices": 0})
cw["Country"] = cw["Country"].fillna(cust_c.groupby("CustomerID")["Country"].first())
cw["Cancel_Rate_Value"] = np.where(cw["Gross_Sales"] > 0, (cw["Cancelled_Value"] / cw["Gross_Sales"]).clip(upper=1.0), 1.0)
cust_exc = log[log["CustomerID"].notna()].groupby("CustomerID").agg(
    Exception_Lines=("Rule", "size"), High_Exceptions=("Severity", lambda s: (s.astype(str) == "High").sum()))
cw = cw.join(cust_exc).fillna({"Exception_Lines": 0, "High_Exceptions": 0})
cw["Risk_Tier"] = np.select(
    [
        (cw["Cancel_Rate_Value"] >= CANCEL_RATE_HIGH) & (cw["Cancelled_Value"] >= CANCEL_VALUE_HIGH),
        (cw["Cancel_Rate_Value"] >= CANCEL_RATE_MED) & (cw["Cancelled_Value"] >= CANCEL_VALUE_MED),
    ],
    ["High", "Medium"], default="Low",
)
cw["Reason"] = np.where(
    cw["Risk_Tier"] == "High", "Cancelled >=30% of purchases and >=$1,000",
    np.where(cw["Risk_Tier"] == "Medium", "Cancelled >=15% of purchases and >=$500", ""),
)
cw["Revenue_Share_%"] = cw["Gross_Sales"] / cw["Gross_Sales"].sum() * 100
customer_watchlist = (
    cw[cw["Risk_Tier"] != "Low"].reset_index()
    .sort_values(["Risk_Tier", "Cancelled_Value"], ascending=[True, False])
)
customer_watchlist["CustomerID"] = customer_watchlist["CustomerID"].astype(int)

# Product price consistency (pricing discrepancies) + cancellation hot spots
pw = canc_product.copy()
price_stats = sales_lines.groupby("StockCode")["UnitPrice"].agg(
    Min_Price="min", Median_Price="median", Max_Price="max", Distinct_Prices="nunique", Lines="size")
price_stats["Max_to_Median"] = price_stats["Max_Price"] / price_stats["Median_Price"]
pw = pw.merge(price_stats[["Median_Price", "Max_Price", "Max_to_Median"]], left_on="StockCode", right_index=True, how="left")
pw["Flag_Cancel"] = (pw["Cancel_Rate_Value"] >= PRODUCT_CANCEL_RATE_FLAG) & (pw["Gross_Sales"] >= PRODUCT_MIN_SALES_VALUE)
pw["Flag_Price"] = (pw["Max_to_Median"] >= PRICE_OUTLIER_MULT) & (pw["Gross_Sales"] >= PRODUCT_MIN_SALES_VALUE)
product_watchlist = pw[pw["Flag_Cancel"] | pw["Flag_Price"]].copy()
product_watchlist["Flag"] = np.select(
    [product_watchlist["Flag_Cancel"] & product_watchlist["Flag_Price"], product_watchlist["Flag_Cancel"]],
    ["High cancellation + price spread", "High cancellation rate"], default="Price spread vs median")
product_watchlist = product_watchlist.drop(columns=["Flag_Cancel", "Flag_Price"]).sort_values(
    "Cancelled_Value", ascending=False).head(50)

# ----------------------------------------------------------------------------
# 6b. Reversed orders: large cancellations that exactly mirror an earlier sale
# ----------------------------------------------------------------------------
ident = products[products["CustomerID"].notna()]
c_big = ident[ident["IsCancel"] & (ident["Revenue"].abs() >= REVERSED_ORDER_MIN_VALUE)]
s_all = ident[(~ident["IsCancel"]) & (ident["Quantity"] > 0)]
rev = c_big.assign(MatchQty=-c_big["Quantity"]).merge(
    s_all, left_on=["CustomerID", "StockCode", "MatchQty"], right_on=["CustomerID", "StockCode", "Quantity"],
    suffixes=("_cancel", "_sale"))
rev = rev[rev["InvoiceDate_sale"] <= rev["InvoiceDate_cancel"]]
rev = rev.sort_values("InvoiceDate_sale").groupby("InvoiceNo_cancel", as_index=False).last()
reversed_orders = pd.DataFrame({
    "CustomerID": rev["CustomerID"].astype(int),
    "StockCode": rev["StockCode"],
    "Description": rev["Description_sale"],
    "Sale_Invoice": rev["InvoiceNo_sale"],
    "Sale_Date": rev["InvoiceDate_sale"].dt.strftime("%Y-%m-%d"),
    "Cancel_Invoice": rev["InvoiceNo_cancel"],
    "Cancel_Date": rev["InvoiceDate_cancel"].dt.strftime("%Y-%m-%d"),
    "Days_To_Reversal": (rev["InvoiceDate_cancel"] - rev["InvoiceDate_sale"]).dt.days,
    "Quantity": rev["Quantity_sale"],
    "Sale_Value": rev["Revenue_sale"],
    "Cancelled_Value": -rev["Revenue_cancel"],
}).sort_values("Sale_Value", ascending=False).reset_index(drop=True)
reversed_value = reversed_orders["Sale_Value"].sum()

# ----------------------------------------------------------------------------
# 7. Daily revenue anomalies (trailing 28-day robust z-score)
# ----------------------------------------------------------------------------
daily = clean.assign(Date=clean["InvoiceDate"].dt.normalize()).groupby("Date")["Revenue"].sum().reset_index()
roll = daily["Revenue"].rolling(28, min_periods=14)
daily["Rolling_Median"] = roll.median().shift(1)
daily["Rolling_MAD"] = daily["Revenue"].rolling(28, min_periods=14).apply(
    lambda x: np.median(np.abs(x - np.median(x))), raw=True).shift(1)
daily["Robust_Z"] = 0.6745 * (daily["Revenue"] - daily["Rolling_Median"]) / daily["Rolling_MAD"].replace(0, np.nan)
daily["Anomaly"] = daily["Robust_Z"].abs() > 3.5
daily["Direction"] = np.where(daily["Robust_Z"] > 3.5, "Spike", np.where(daily["Robust_Z"] < -3.5, "Drop", ""))
daily["Date"] = daily["Date"].dt.strftime("%Y-%m-%d")
anomalies = daily[daily["Anomaly"]].copy()

# ----------------------------------------------------------------------------
# 8. Save CSVs (Power BI / Tableau ready)
# ----------------------------------------------------------------------------
outputs = {
    "r01_reconciliation_bridge.csv": bridge,
    "r02_exception_rule_summary.csv": rule_summary,
    "r03_exception_log.csv": exception_log,
    "r04_cancellations_by_month.csv": canc_month,
    "r05_cancellations_by_country.csv": canc_country,
    "r06_product_watchlist.csv": product_watchlist,
    "r07_customer_watchlist.csv": customer_watchlist,
    "r08_concentration_summary.csv": conc,
    "r09_daily_revenue_anomalies.csv": daily,
    "r10_reversed_orders.csv": reversed_orders,
}
for fn, df in outputs.items():
    df.to_csv(os.path.join(OUT_DIR, fn), index=False)

# ----------------------------------------------------------------------------
# 9. Charts
# ----------------------------------------------------------------------------
plt.rcParams.update({"font.size": 10, "axes.spines.top": False, "axes.spines.right": False})

# 9a. reconciliation waterfall (rows)
fig, ax = plt.subplots(figsize=(9, 4.5))
labels = ["Raw rows", "Duplicates", "No CustomerID", "Cancelled", "Cleaned rows"]
vals = bridge["Rows"].tolist()
running = vals[0]
ax.bar(0, vals[0], color=NAVY)
for i, v in enumerate(vals[1:-1], start=1):
    ax.bar(i, v, bottom=running, color=RED)
    ax.text(i, running + v / 2, f"{v:,}", ha="center", va="center", color="white", fontsize=9)
    running += v
ax.bar(4, vals[-1], color=TEAL)
ax.text(0, vals[0] / 2, f"{vals[0]:,}", ha="center", color="white")
ax.text(4, vals[-1] / 2, f"{vals[-1]:,}", ha="center", color="white")
ax.set_xticks(range(5)); ax.set_xticklabels(labels)
ax.set_title("Row reconciliation: raw file to reported dataset (ties out exactly)")
ax.yaxis.set_major_formatter(lambda x, _: f"{x/1000:,.0f}k")
fig.tight_layout(); fig.savefig(f"{IMG_DIR}/risk_reconciliation_waterfall.png", dpi=150); plt.close(fig)

# 9b. cancellation rate by month
cm = canc_month[canc_month["Month"] != "2011-12"]
fig, ax = plt.subplots(figsize=(9, 4.5))
ax.bar(cm["Month"], cm["Cancel_Rate_Value"] * 100, color=AMBER)
ax.axhline(cm["Cancelled_Value"].sum() / cm["Gross_Sales"].sum() * 100, color=NAVY, ls="--", lw=1.2,
           label="Full-period average")
ax.set_ylabel("Cancelled value as % of gross sales"); ax.legend(frameon=False)
ax.set_title("Cancellation rate by month (Dec-2011 excluded: partial month)")
plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
fig.tight_layout(); fig.savefig(f"{IMG_DIR}/risk_cancellation_rate_by_month.png", dpi=150); plt.close(fig)

# 9c. exceptions by rule
base_sev = {k: v[2] for k, v in rules.items()} | {"EXACT_DUPLICATE_LINE": "Medium", "MISSING_CUSTOMER_ID": "Medium"}
rs = rule_summary.groupby("Rule", as_index=False).agg(Lines=("Lines", "sum"), Est_Value_At_Risk=("Est_Value_At_Risk", "sum"))
rs = rs[rs["Rule"] != "MISSING_CUSTOMER_ID"].sort_values("Est_Value_At_Risk")
fig, ax = plt.subplots(figsize=(9, 4.8))
colors = rs["Rule"].map(base_sev).map({"High": RED, "Medium": AMBER, "Low": GREY})
ax.barh(rs["Rule"], rs["Est_Value_At_Risk"], color=colors)
for y, (v, n) in enumerate(zip(rs["Est_Value_At_Risk"], rs["Lines"])):
    ax.text(v, y, f"  {money(v)}  ({n:,} lines)", va="center", fontsize=8)
ax.set_xlim(0, rs["Est_Value_At_Risk"].max() * 1.35)
ax.xaxis.set_major_formatter(lambda x, _: f"${x/1000:,.0f}k")
ax.set_title("Exceptions by rule: absolute value at risk (colour = rule's base severity)")
fig.tight_layout(); fig.savefig(f"{IMG_DIR}/risk_exceptions_by_rule.png", dpi=150); plt.close(fig)

# 9d. customer Pareto
fig, ax = plt.subplots(figsize=(7.5, 4.5))
ax.plot(pareto_cust["Cum_Customers_%"], pareto_cust["Cum_Revenue_%"], color=NAVY, lw=2)
ax.plot([0, 100], [0, 100], color=GREY, ls=":", label="Perfect equality")
n80 = conc.loc[conc["Dimension"] == "Customer", "Entities_For_80%_as_%_of_All"].iloc[0]
ax.axhline(80, color=AMBER, ls="--", lw=1); ax.axvline(n80, color=AMBER, ls="--", lw=1)
ax.annotate(f"{n80:.0f}% of customers\ngenerate 80% of revenue", (n80, 80), (n80 + 8, 55), arrowprops=dict(arrowstyle="->"))
ax.set_xlabel("% of customers (largest first)"); ax.set_ylabel("Cumulative % of revenue"); ax.legend(frameon=False)
ax.set_title("Customer revenue concentration")
fig.tight_layout(); fig.savefig(f"{IMG_DIR}/risk_customer_pareto.png", dpi=150); plt.close(fig)

# 9e. daily revenue with anomalies
dd = daily.copy(); dd["Date_dt"] = pd.to_datetime(dd["Date"])
fig, ax = plt.subplots(figsize=(10, 4.5))
ax.plot(dd["Date_dt"], dd["Revenue"], color=NAVY, lw=1)
a = dd[dd["Anomaly"]]
ax.scatter(a["Date_dt"], a["Revenue"], color=RED, zorder=3, s=22, label=f"Anomalous days ({len(a)})")
ax.set_title("Daily revenue with robust-z anomaly flags (|z| > 3.5 vs trailing 28 trading days)")
ax.yaxis.set_major_formatter(lambda x, _: f"${x/1000:,.0f}k"); ax.legend(frameon=False)
fig.tight_layout(); fig.savefig(f"{IMG_DIR}/risk_daily_revenue_anomalies.png", dpi=150); plt.close(fig)

# ----------------------------------------------------------------------------
# 10. Excel report (formulas live; Exception_Log is an Excel Table, pivot-ready)
# ----------------------------------------------------------------------------
HDR_FILL = PatternFill("solid", fgColor="1F3A5F")
HDR_FONT = Font(bold=True, color="FFFFFF")
thin = Side(style="thin", color="D0D4DA")


def write_df(ws, df, start_row=1, table_name=None, money_cols=(), pct_cols=(), int_cols=()):
    for j, c in enumerate(df.columns, 1):
        cell = ws.cell(row=start_row, column=j, value=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for i, row in enumerate(df.itertuples(index=False), start_row + 1):
        for j, v in enumerate(row, 1):
            if isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = None if np.isnan(v) else float(v)
            elif isinstance(v, (np.bool_, bool)):
                v = bool(v)
            elif v is pd.NA or (isinstance(v, float) and np.isnan(v)):
                v = None
            ws.cell(row=i, column=j, value=v)
    for j, c in enumerate(df.columns, 1):
        letter = get_column_letter(j)
        width = max(len(str(c)), *(len(str(x)) for x in df[c].head(200))) + 2
        ws.column_dimensions[letter].width = min(max(width, 10), 55)
        fmt = "$#,##0.00" if c in money_cols else "0.0%" if c in pct_cols else "#,##0" if c in int_cols else None
        if fmt:
            for i in range(start_row + 1, start_row + len(df) + 1):
                ws.cell(row=i, column=j).number_format = fmt
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    if table_name:
        ref = f"A{start_row}:{get_column_letter(len(df.columns))}{start_row + len(df)}"
        t = Table(displayName=table_name, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)


wb = Workbook()

# --- Exception_Log (build first so other sheets can reference it)
ws_log = wb.active
ws_log.title = "Exception_Log"
log_x = exception_log.copy()
write_df(ws_log, log_x, table_name="tblExceptions",
         money_cols=("Revenue", "Est_Value_At_Risk", "UnitPrice", "Median_Price"), int_cols=("Quantity", "Median_Qty"))
n_log = len(log_x)
col = {c: get_column_letter(i + 1) for i, c in enumerate(log_x.columns)}
rng = lambda c: f"Exception_Log!${col[c]}$2:${col[c]}${n_log + 1}"

# --- Reconciliation
ws_rec = wb.create_sheet("Reconciliation")
ws_rec.append(["Step", "Rows", "Net value (USD)"])
for c in ws_rec[1]:
    c.fill, c.font = HDR_FILL, HDR_FONT
for r in bridge.itertuples(index=False):
    ws_rec.append([r.Step, int(r.Rows), float(r.Net_Value_USD)])
ws_rec.append(["Check: sum of steps less cleaned (must be 0)", "=SUM(B2:B5)-B6", "=ROUND(SUM(C2:C5)-C6,2)"])
ws_rec["A8"] = "Memo: cancelled value among identified customers (not netted in reported revenue)"
ws_rec["C8"] = float(cancelled_identified_value)
ws_rec["A9"] = "Memo: cancelled value as % of reported (gross) revenue"
ws_rec["C9"] = "=C8/C6"
ws_rec["A10"] = "Memo: net value of lines with no CustomerID (excluded from reported revenue)"
ws_rec["C10"] = float(no_cust_value)
for r in range(2, 11):
    ws_rec.cell(row=r, column=2).number_format = "#,##0"
    ws_rec.cell(row=r, column=3).number_format = "$#,##0.00"
ws_rec["C9"].number_format = "0.0%"
ws_rec.column_dimensions["A"].width = 78
ws_rec.column_dimensions["B"].width = 14
ws_rec.column_dimensions["C"].width = 20

# --- Exception_Matrix: pivot-style summary built from live COUNTIFS / SUMIFS formulas
ws_mx = wb.create_sheet("Exception_Matrix")
rules_logged = [r for r in dict.fromkeys(rule_summary["Rule"]) if r not in ("EXACT_DUPLICATE_LINE", "MISSING_CUSTOMER_ID")]
ws_mx.append(["Rule (from Exception_Log)", "High lines", "Medium lines", "Low lines", "Total lines",
              "Est. value at risk (USD)", "% of total value"])
for c in ws_mx[1]:
    c.fill, c.font = HDR_FILL, HDR_FONT
for i, rname in enumerate(rules_logged, start=2):
    ws_mx.append([
        rname,
        f'=COUNTIFS({rng("Rule")},$A{i},{rng("Severity")},"High")',
        f'=COUNTIFS({rng("Rule")},$A{i},{rng("Severity")},"Medium")',
        f'=COUNTIFS({rng("Rule")},$A{i},{rng("Severity")},"Low")',
        f"=SUM(B{i}:D{i})",
        f'=SUMIFS({rng("Est_Value_At_Risk")},{rng("Rule")},$A{i})',
        f"=F{i}/$F${len(rules_logged) + 2}",
    ])
tot = len(rules_logged) + 2
ws_mx.append(["TOTAL"] + [f"=SUM({get_column_letter(c)}2:{get_column_letter(c)}{tot - 1})" for c in range(2, 7)] + [f"=F{tot}/$F${tot}"])
for c in ws_mx[tot]:
    c.font = Font(bold=True)
for r in range(2, tot + 1):
    for c in range(2, 6):
        ws_mx.cell(row=r, column=c).number_format = "#,##0"
    ws_mx.cell(row=r, column=6).number_format = "$#,##0"
    ws_mx.cell(row=r, column=7).number_format = "0.0%"
ws_mx.cell(row=tot + 2, column=1, value="Month-by-rule view (COUNTIFS on Exception_Log[Month])").font = Font(bold=True)
months = sorted(exception_log["Month"].unique())
hdr_r = tot + 3
ws_mx.cell(row=hdr_r, column=1, value="Month")
for j, rname in enumerate(rules_logged, start=2):
    ws_mx.cell(row=hdr_r, column=j, value=rname)
ws_mx.cell(row=hdr_r, column=len(rules_logged) + 2, value="Total")
for c in ws_mx[hdr_r]:
    if c.value:
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(wrap_text=True)
for k, m in enumerate(months, start=hdr_r + 1):
    ws_mx.cell(row=k, column=1, value=m)
    for j, rname in enumerate(rules_logged, start=2):
        L = get_column_letter(j)
        ws_mx.cell(row=k, column=j, value=f'=COUNTIFS({rng("Month")},$A{k},{rng("Rule")},{L}${hdr_r})')
    last = get_column_letter(len(rules_logged) + 1)
    ws_mx.cell(row=k, column=len(rules_logged) + 2, value=f"=SUM(B{k}:{last}{k})")
ws_mx.column_dimensions["A"].width = 34
for j in range(2, 12):
    ws_mx.column_dimensions[get_column_letter(j)].width = 18
ws_mx.freeze_panes = "B2"

# --- Data sheets
sheets = [
    ("Rule_Summary", rule_summary, dict(money_cols=("Est_Value_At_Risk",), int_cols=("Lines",))),
    ("Cancel_By_Month", canc_month, dict(money_cols=("Gross_Sales", "Cancelled_Value"), pct_cols=("Cancel_Rate_Value",))),
    ("Cancel_By_Country", canc_country, dict(money_cols=("Gross_Sales", "Cancelled_Value"), pct_cols=("Cancel_Rate_Value",))),
    ("Product_Watchlist", product_watchlist, dict(money_cols=("Gross_Sales", "Cancelled_Value", "Median_Price", "Max_Price"), pct_cols=("Cancel_Rate_Value",))),
    ("Customer_Watchlist", customer_watchlist, dict(money_cols=("Gross_Sales", "Cancelled_Value"), pct_cols=("Cancel_Rate_Value",))),
    ("Reversed_Orders", reversed_orders, dict(money_cols=("Sale_Value", "Cancelled_Value"), int_cols=("Quantity",))),
    ("Concentration", conc, {}),
    ("Daily_Anomalies", anomalies, dict(money_cols=("Revenue", "Rolling_Median", "Rolling_MAD"))),
]
for name, df, kw in sheets:
    ws = wb.create_sheet(name)
    write_df(ws, df, **kw)

TIER_COL = get_column_letter(list(customer_watchlist.columns).index("Risk_Tier") + 1)
# --- Summary sheet (first tab)
ws = wb.create_sheet("Summary", 0)
ws["A1"] = "Retail Risk & Exception Monitoring: Management Summary"
ws["A1"].font = Font(size=16, bold=True, color="1F3A5F")
ws["A2"] = "Source: UCI Online Retail transactions, 1 Dec 2010 to 9 Dec 2011. All figures in USD as recorded in the file."
rows = [
    ("Raw rows", "=Reconciliation!B2", "#,##0"),
    ("Cleaned rows used for reported KPIs", "=Reconciliation!B6", "#,##0"),
    ("Reconciliation check (must be 0)", "=Reconciliation!B7", "#,##0"),
    ("Reported (gross) revenue", "=Reconciliation!C6", "$#,##0"),
    ("Cancelled value not netted in reported revenue", "=Reconciliation!C8", "$#,##0"),
    ("  as % of reported revenue", "=Reconciliation!C9", "0.0%"),
    ("Logged exception lines (excl. duplicates and missing IDs)", "=Exception_Matrix!E%d" % tot, "#,##0"),
    ("Estimated value at risk in logged exceptions", "=Exception_Matrix!F%d" % tot, "$#,##0"),
    ("High-severity exception lines", "=Exception_Matrix!B%d" % tot, "#,##0"),
    ("High-risk customers on watchlist", '=COUNTIF(Customer_Watchlist!%s2:%s%d,"High")' % (TIER_COL, TIER_COL, len(customer_watchlist) + 1), "#,##0"),
    ("Sales later fully reversed (>= $5k, matched cancellations)", "=SUM(Reversed_Orders!J2:J%d)" % (len(reversed_orders) + 1), "$#,##0"),
    ("Anomalous trading days", "=COUNTA(Daily_Anomalies!A2:A%d)" % (len(anomalies) + 1), "#,##0"),
]
for i, (lab, f, fmt) in enumerate(rows, start=4):
    ws.cell(row=i, column=1, value=lab)
    c = ws.cell(row=i, column=2, value=f)
    c.number_format = fmt
    c.font = Font(bold=True)
ws["A17"] = "How to read this workbook"
ws["A17"].font = Font(bold=True, size=12)
guide = [
    "Reconciliation: row and dollar bridge from raw file to reported KPIs. The check cell must read 0.",
    "Exception_Log: one line per flagged transaction (Excel Table 'tblExceptions'); select any cell > Insert > PivotTable to slice by rule, severity, month, country or customer.",
    "Exception_Matrix: formula-driven summary (COUNTIFS / SUMIFS) that updates if you edit or filter the log.",
    "Cancel_By_*: cancellation rate = cancelled value / gross sales value, product SKUs only.",
    "Customer_Watchlist / Product_Watchlist: ranked follow-up lists with the rule that triggered each entry.",
    "Daily_Anomalies: days whose revenue deviates >3.5 robust-z from the trailing 28 trading days (seasonal peaks can trigger this).",
    "Thresholds are set at the top of risk_exception_analysis.py and can be changed and re-run.",
]
for i, g in enumerate(guide, start=18):
    ws.cell(row=i, column=1, value="- " + g)
ws.column_dimensions["A"].width = 62
ws.column_dimensions["B"].width = 20

wb.save(XLSX_PATH)

# ----------------------------------------------------------------------------
# 11. Console summary (paste-ready findings)
# ----------------------------------------------------------------------------
print("\n=== RECONCILIATION ===")
print(bridge.to_string(index=False))
print(f"Cancelled value among identified customers: {money(cancelled_identified_value)} "
      f"({cancelled_identified_value / gross_identified_value:.1%} of reported revenue)")
print(f"Net value of lines with no CustomerID: {money(no_cust_value)}")
print("\n=== EXCEPTION RULES ===")
print(rule_summary[["Rule", "Severity", "Lines", "Est_Value_At_Risk"]].to_string(index=False))
print(f"Logged exception lines: {len(exception_log):,}   High: {(exception_log['Severity'] == 'High').sum():,}")
print("\n=== CANCELLATIONS ===")
tot_rate = canc_month["Cancelled_Value"].sum() / canc_month["Gross_Sales"].sum()
print(f"Overall cancel rate (product SKUs): {tot_rate:.1%}")
print(canc_month[["Month", "Gross_Sales", "Cancelled_Value", "Cancel_Rate_Value"]].round(3).to_string(index=False))
print(canc_country.head(8)[["Country", "Gross_Sales", "Cancelled_Value", "Cancel_Rate_Value"]].round(3).to_string(index=False))
print("\n=== REVERSED ORDERS ===")
print(reversed_orders.to_string(index=False))
print("Total sale value later fully reversed:", money(reversed_value), f"= {reversed_value/gross_identified_value:.1%} of reported revenue")
print("\n=== CONCENTRATION ===")
print(conc.to_string(index=False)); print(conc_note)
print("\n=== WATCHLISTS ===")
print(customer_watchlist["Risk_Tier"].value_counts().to_dict())
print(customer_watchlist.head(8)[["CustomerID", "Country", "Gross_Sales", "Cancelled_Value", "Cancel_Rate_Value", "Risk_Tier"]].round(2).to_string(index=False))
print(product_watchlist.head(8)[["StockCode", "Description", "Gross_Sales", "Cancelled_Value", "Cancel_Rate_Value", "Max_to_Median", "Flag"]].round(2).to_string(index=False))
print("\n=== DAILY ANOMALIES ===")
print(len(anomalies), "anomalous days;", anomalies["Direction"].value_counts().to_dict())
print(anomalies.head(12)[["Date", "Revenue", "Rolling_Median", "Robust_Z", "Direction"]].round(1).to_string(index=False))
print("\nDone. CSVs in outputs/risk/, charts in images/risk_*.png, workbook at", XLSX_PATH)
