# Superstore Sales Analytics Dashboard (Excel)

## Project Overview

This project presents an advanced **Excel analytics dashboard** built from the *Superstore Orders* dataset.

The goal of this project is to demonstrate how Excel can be used as a **Business Intelligence tool** by combining:

* Power Query (data cleaning and transformation)
* Power Pivot (data modeling and measures)
* Pivot Tables
* Interactive dashboard elements
* VBA automation

The dashboard provides insights into **sales performance, profitability, discount impact, customer behavior, and shipping efficiency**.

This project was designed as a **portfolio project for Data Analyst / BI roles**, showcasing advanced Excel data modeling and dashboarding skills.

---

# Dataset

Dataset: Superstore Orders

The dataset contains transactional sales data including:

* Order information
* Customer segments
* Product categories and subcategories
* Sales and profit values
* Discounts
* Shipping information
* Regions and markets

Main fields used:

* Order ID
* Order Date
* Product Category / Subcategory
* Customer Segment
* Region
* Sales
* Profit
* Discount
* Shipping Cost

---

# Data Preparation

Data preparation was performed using **Power Query**.

Key transformations included:

* Data type corrections (numeric / decimal values)
* Cleaning of inconsistent numeric formats
* Creation of calculated fields:

  * `shipping_days`
  * `shipping_time_cat`
* Data structuring for analytical modeling

The cleaned dataset was then loaded into **Power Pivot**.

---

# Data Model

The analytical model was built using **Power Pivot**.

Structure:

Fact table:

* `fact_sales`

Supporting dimensions created from the dataset:

* Products
* Customers
* Regions
* Dates

Measures were created using **DAX** to compute key performance indicators.

Main measures include:

* Total Sales
* Total Profit
* Margin %
* Average Discount
* Total Quantity

---

# Dashboard Features

The dashboard provides interactive business analysis with **slicers and KPIs**.

### Interactive Filters

Users can filter the dashboard by:

* Region
* Product Category
* Customer Segment
* Year

A **VBA button** was implemented to reset all filters instantly.

---

# Key Visualizations

### Sales & Profit by Region

Shows regional performance by comparing total sales and profit.

### Category Profitability

Analyzes the profitability of each product category.

### Margin vs Discount

Scatter plot analyzing the relationship between discounts and profit margin.

### Top 10 Products

Highlights the best performing products by sales.

### Sales Trend

Displays sales evolution over time.

### Shipping Performance

Provides insights into shipping delays and delivery performance.

### Shipping Time Distribution

Shows how long orders typically take to ship.

### Worst 5 Products

Identifies the least profitable products.

---

# Tools Used

* Microsoft Excel
* Power Query
* Power Pivot
* Pivot Tables
* DAX Measures
* VBA (for dashboard interaction)

---

# Project Goal

This project demonstrates how Excel can be used to build a **complete business intelligence solution**, including:

* Data transformation
* Data modeling
* KPI creation
* Interactive dashboards

The objective was to replicate the type of **analytical dashboards used in real business environments** using Excel.

