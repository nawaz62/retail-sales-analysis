# Indian E-Commerce Retail Sales Analysis

A self-directed data analytics project built in Excel, analyzing 18 months (Jan 2024 – Jun 2025)
of retail order data across Indian regions, categories, and customer segments.

## Objective
Practice end-to-end data analysis workflow — from raw transactional data to a decision-ready
KPI dashboard — using the same tools (Excel pivot logic, dynamic formulas, charts) commonly
expected in Data Analyst internship roles.

## Dataset
`retail_sales_raw.csv` — **1,800 synthetic but realistically distributed orders** (Jan 2024–Jun 2025),
generated to mimic a real Indian e-commerce retailer: seasonal demand spikes around the festive
season (Oct–Nov), category-specific pricing and margins, discount-driven profit compression, and
regional/city-level spread across North, South, East, and West India.
*(This is a practice dataset built for portfolio purposes, not scraped or proprietary company data.)*

**Columns:** Order ID, Order Date, Region, City, Customer Segment, Category, Sub-Category,
Quantity, Unit Price, Discount, Sales, Profit, Ship Mode.

## Tools & Techniques
- **Excel** — `SUMIF`/`SUMIFS`, `COUNTIF`, dynamic formulas (no hardcoded values — every summary
  cell recalculates from the raw data table)
- **Excel Tables** for the raw dataset (structured, filterable, sortable)
- **Native Excel charts** — bar, line, and pie — built directly from formula-driven summary tables

## What's in the workbook (`Retail_Sales_Analysis.xlsx`)
| Sheet | Contents |
|---|---|
| **Dashboard** | Top-line KPIs (Total Sales, Total Profit, Profit Margin %, Orders, AOV) + narrative insights |
| **Category Summary** | Sales/profit/margin by product category, with bar + pie charts |
| **Region Summary** | Sales/profit/margin by region, with bar chart |
| **Monthly Trend** | Month-by-month sales and profit line chart, Jan 2024–Jun 2025 |
| **Top Sub-Categories** | Top 10 sub-categories ranked by sales |
| **Raw Data** | Full 1,800-row transactional dataset as a structured Excel Table |

## Key Insights
- Electronics and Clothing drive the largest share of total sales, but Clothing carries a
  meaningfully higher profit margin per rupee sold than Electronics.
- Sales peak in October–November (festive season) and dip during summer months — a seasonal
  pattern that should inform inventory and staffing decisions.
- Discounts of 20% or more visibly compress profit margin, suggesting deep discounts above that
  threshold warrant a manual review step before approval.
- West and South regions generate the most orders; North and East show room for targeted
  marketing investment.

## How to Explore
1. Open `Retail_Sales_Analysis.xlsx` in Excel.
2. Start on the **Dashboard** tab for the headline numbers.
3. Drill into **Category Summary**, **Region Summary**, **Monthly Trend**, and **Top Sub-Categories**
   for the supporting breakdowns.
4. All summary values are live formulas pulling from **Raw Data** — filter or edit the raw table
   and the dashboard updates automatically.

---
*Built by Md Nawazish Husain as a practice project to demonstrate Excel-based data analysis skills.*
