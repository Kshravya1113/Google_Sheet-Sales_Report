# Regional Sales Analysis: Excel Formula-Driven Analytics

An Excel-based analytics project that turns transactional sales data into a state-level performance report. It uses lookup, conditional aggregation, ranking, date, and classification formulas to produce business-ready metrics for sales managers and analysts.

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?logo=microsoft-excel&logoColor=white)
![Domain](https://img.shields.io/badge/Domain-Sales%20Analytics-blue)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen)

---

## Table of Contents
- [Overview](#overview)
- [Key Results](#key-results)
- [Dataset](#dataset)
- [Derived Columns and Formulas](#derived-columns-and-formulas)
- [Worked Example: Texas](#worked-example-texas)
- [Key Insights](#key-insights)
- [Implementation Notes](#implementation-notes)
- [Repository Structure](#repository-structure)
- [Skills Demonstrated](#skills-demonstrated)
- [Author](#author)

---

## Overview

The project takes transaction-level data (`Sheet1`) and builds a summary table of **20 states grouped by 4 regions** (West, East, Central, South). Each state is measured on:

- Order counts and total sales
- Sales rank and percentage contribution
- Customer segment splits (Corporate, Consumer, Home Office)
- First and last order dates, with recency calculated from today
- Average sale, total profit, and profit after tax
- Profit slab classification, with low-profit states highlighted

**Audience:** Business analysts and sales managers.

---

## Key Results

| Metric | Value |
|---|---|
| Total Orders | **3,467** |
| Total Sales | **438,097** |
| States Analyzed | **20** |
| Top State | **California** (96,863, 22.11% of sales) |
| Top Region | **West** (189,145) |
| Dominant Segment | **Consumer** |

---

## Dataset

Source data lives in `Sheet1` (transaction rows). The summary sheet references these columns:

| Sheet1 Column | Field |
|---|---|
| `C` | Order Date |
| `G` | Customer Segment |
| `I` | State |
| `J` | Region |
| `P` | Sales |
| `Q` | Quantity |
| `S` | Profit |

---

## Derived Columns and Formulas

| Col | Field | Formula | Purpose |
|---|---|---|---|
| B | Region | `=VLOOKUP(A2,Sheet1!I2:J2539,2,FALSE)` | Maps each state to its region |
| C | Label | `=CONCATENATE(A2," : ",B2)` | Builds a "State : Region" label |
| D | Total Orders | `=COUNTIF(Sheet1!I:I,Sheet1!I2)` | Counts orders per state |
| E | Total Sales | `=SUMIF(Sheet1!I:I,A2,Sheet1!P:P)` | Sums sales per state |
| F | Rank | `=RANK(E2,E$2:E$21,0)` | Ranks states by sales (descending) |
| G | Contribution % | `=E2/E$22` | State share of total sales |
| H–J | Segment Qty | `=SUMIFS(Sheet1!Q:Q,Sheet1!I:I,A2,Sheet1!G:G,$H$1)` | Quantity by segment (Corporate, Consumer, Home Office) |
| K | Max Segment | `=XLOOKUP(MAX(H2:J2),H2:J2,$H$1:$J$1)` | Top segment by quantity |
| L | First Order | `=MINIFS(Sheet1!C:C,Sheet1!I:I,A2)` | Earliest order date |
| M | Last Order | `=MAXIFS(Sheet1!C:C,Sheet1!I:I,A2)` | Latest order date |
| N | Recency | `=DATEDIF(M2,TODAY(),"Y") & " Years " & DATEDIF(M2,TODAY(),"YM") & " Months " & DATEDIF(M2,TODAY(),"MD") & " Day "` | Time since last order |
| O | Last Order Month | `=TEXT(M2,"mmm-yy")` | Short month label (e.g., Dec-17) |
| Q | Avg Sale / Customer | `=AVERAGEIF(Sheet1!I:I,A2,Sheet1!P:P)` | Average sale per state |
| R | Total Profit | `=SUMIF(Sheet1!I:I,A2,Sheet1!S:S)` | Total profit per state |
| S | Profit After Tax (18%) | `=SUMIF(Sheet1!I:I,A2,Sheet1!S:S)` (tax applied to positive profit) | Post-tax profit |
| T | Profit Slab | `=IF(R2<10000,"<10K",IF(AND(R2>=10000,R2<=20000),"10-20K",">20K"))` | Buckets states by profit |
| U | Highlight | Conditional formatting rule `=R2<10000` | Flags low-profit states in red |

---

## Worked Example: Texas

| Step | Result |
|---|---|
| Region lookup | Central |
| Label | Texas : Central |
| Total Orders | 232 |
| Total Sales | 32,876.26 |
| Sales Rank | 3 |
| Contribution | 7.50% |
| Segment Qty | Corporate 259, Consumer 406, Home Office 160 → **Consumer** |
| Order Range | 1/3/2014 → 26/12/2017 |
| Avg Sale / Customer | 142 |
| Total Profit | -3,323.43 → slab `<10K`, highlighted red |

---

## Key Insights

1. **Sales are concentrated.** California, New York, and Texas drive about **46.9%** of total sales, which makes them priority markets for campaign investment.
2. **Consumer leads.** The Consumer segment is the top contributor in nearly every state, so Consumer-focused promotions are likely to give the best yield.
3. **Profitability is a concern.** Most states fall in the `<10K` profit slab, and several show negative profit. Pricing, returns, and cost allocation in these markets deserve investigation.

---

## Implementation Notes

1. Confirm that the `Sheet1` ranges (I, P, Q, S) contain the expected fields (State, Sales, Quantity, Profit).
2. Use absolute references (`E$2:E$21`, `$H$1:$J$1`) to avoid copy-down errors.
3. Ensure `H2:J2` are numeric and the `$H$1:$J$1` headers exactly match the segment names in the source data.
4. Apply conditional formatting with the rule `=R2<10000` to color low-profit cells red.
5. Test edge cases: missing dates (`DATEDIF`), empty segments, and negative profits.

> **Compatibility:** `XLOOKUP`, `MINIFS`, and `MAXIFS` require Excel 2019 / Microsoft 365 or later.

---

## Repository Structure

```
├── Sales_Analysis.xlsx          # Workbook with source data and summary sheet
├── Sales_Report-by_Kshravya.pdf # Presentation documenting formulas and findings
└── README.md
```

---

## Skills Demonstrated

- **Excel:** `VLOOKUP`, `XLOOKUP`, `SUMIF(S)`, `COUNTIF`, `AVERAGEIF`, `MINIFS/MAXIFS`, `RANK`, `DATEDIF`, `TEXT`, nested `IF/AND`
- **Data analysis:** aggregation, ranking, segmentation, contribution analysis
- **Data visualization:** dashboards, bar and donut charts, conditional formatting
- **Business reporting:** turning raw data into recommendations for decision-makers

---

## Author

**Kshravya**
