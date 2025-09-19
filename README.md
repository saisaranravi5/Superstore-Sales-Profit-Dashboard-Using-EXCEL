📊 Superstore Sales & Profit Dashboard


🔹 Project Overview

This project uses the Sample Superstore Sales dataset (9,994 rows, 21 columns) to design an interactive Excel dashboard that provides insights into sales, profit, discounts, customers, products, and regions.

The dashboard simulates how a retail merchandising team (like Dollar General, Walmart, or Target) might analyze performance and promotions.

🔹 Objectives

Clean and prepare raw data for analysis.

Build PivotTables for sales, profit, discount, and customer insights.

Design interactive PivotCharts (column, line, scatter, bar, map alternatives).

Add Slicers and a Timeline for dynamic filtering.

Create KPI cards for quick executive summaries.

Automate tasks using recorded Macros for refresh & snapshots.

🔹 Step 1: Data Cleaning & Preparation

Removed unnecessary column (Sa).

Standardized dates (Order Date, Ship Date → mm/dd/yyyy).

Replaced missing values:

Sales / Profit → 0

Category / Customer → "Unknown"

Cleaned text fields with TRIM/PROPER.

Added calculated fields:

Profit Margin = Profit ÷ Sales

Year-Month = TEXT(Order Date,"YYYY-MM")

Enriched dataset with XLOOKUP mappings:

State → Region

Sub-Category → Product ID

Category → Max Discount allowed

🔹 Step 2: PivotTables

Created 6 PivotTables for analysis:

Category & Sub-Category Performance → Sales, Profit, Profit Margin

Regional Performance → Sales & Profit by Region & State

Top 10 Products → by Sales & Profit

Customer Segment Analysis → Sales, Profit, Avg. Discount

Monthly Sales Trend → Sales & Profit over time

Discount vs Profit → Grouped discount brackets & profit

🔹 Step 3: PivotCharts

Column Chart → Category/Sub-Category vs Sales & Profit

Column Chart → Sales & Profit by State (grouped by Region)

Bar Chart → Top 10 Products

Combo Chart → Sales, Profit, and Avg. Discount by Segment

Line Chart → Monthly Sales & Profit Trend (2014–2017)

Scatter/Line → Impact of Discounts on Profit

🔹 Step 4: Interactivity

Added Slicers: Region, Category, Segment

Added Timeline: Order Date

Linked slicers/timeline to all PivotTables using Report Connections.

🔹 Step 5: KPI Cards

At the top of the dashboard, added KPI cards with shapes + linked formulas:

Total Sales ($2.29M)

Total Profit ($286K)

Avg. Discount (15.6%)

Profit Margin (12.5%)

🔹 Step 6: Dashboard Layout

Clean title: “Superstore Sales & Profit Dashboard”

KPIs at top row.

6 charts arranged in 3×2 grid below.

Slicers & Timeline stacked on the right side.

Consistent theme: Sales = Blue, Profit = Orange, Discount = Gray.

Removed gridlines & unnecessary borders.

🔹 Step 7: Macros & Automation

Added recorded macros (no VBA coding required) for automation:

Refresh & Reset Dashboard

Refresh all PivotTables.

Reset all slicers & timeline to “All”.

Assigned to button: Refresh Dashboard.

Snapshot Current View

Copies the dashboard as a picture into a new sheet.

Preserves filtered state for record-keeping.

Assigned to button: Save Snapshot.

(Optional) Set Default View

Applies favorite slicer settings (e.g., Region = West, Segment = Consumer).

Assigned to button: Set Default View.

🔹 Skills Demonstrated

Excel Data Cleaning (TRIM, PROPER, XLOOKUP, IF, TEXT).

PivotTables & PivotCharts.

Interactive Dashboard Design (slicers, timeline, KPIs).

Conditional Formatting.

Macros & Automation (recorded, button-driven).

Visualization & Storytelling.

🔹 Final Deliverable

A professional, interactive Excel dashboard that allows business users to:

1.)Filter by Region, Category, Segment, Time.
2.)See real-time updates in KPIs & charts.
3.)Reset filters or save snapshots with one click.

<img width="757" height="514" alt="Screenshot 2025-09-18 at 10 45 09 PM" src="https://github.com/user-attachments/assets/97f0eecd-62ea-4723-8a6f-148028a8e419" />


🔹 Why This Project?

Retail and e-commerce businesses generate massive amounts of sales and customer data. Managers and analysts often struggle to quickly answer critical questions like:

Which products and categories drive the most profit?
Which customer segments respond best to discounts?
Are promotions actually improving profitability, or just cutting margins?
How are sales and profits trending across time and regions?
This project was done to simulate a real-world merchandising analytics use case: taking raw transactional data and transforming it into a decision-ready dashboard.

🔹 Who Is It Helpful For?

Merchandising Managers → identify profitable product lines and categories.
Regional Sales Teams → track performance by state/region and target weak areas.
Marketing & Promotions Teams → understand the impact of discounts on profitability.
Executives → view high-level KPIs like Total Sales, Profit, Discount %, and Margin in one place.
Analysts → use slicers & timelines to drill down dynamically without writing queries.

🔹 Outcomes of the Project

Built a professional Excel dashboard that consolidates KPIs, visuals, and filters into a single interactive view.
Allowed dynamic exploration of data across products, regions, and customer segments.
Provided insights such as:
Technology and Furniture categories have higher margins compared to Office Supplies.
Some high-discount promotions erode profit, visible in the Discount vs Profit chart.
Consumer segment generates the largest share of sales, but Corporate contributes steady profit.
Seasonal sales spikes are clear in the Monthly Sales Trend (e.g., year-end holidays).
Enhanced with macros to reset filters and save snapshots, making it easy for non-technical users to use and share.

