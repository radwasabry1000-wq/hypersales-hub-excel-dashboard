# 🛒 HyperSales Hub – End-to-End Excel Sales Analysis

> A complete analytics workflow built entirely in **Microsoft Excel**: from raw data exploration to an interactive sales dashboard, using **Power Query**, **Power Pivot (Data Model + DAX)** and **PivotTables**.

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power%20Query-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Power Pivot](https://img.shields.io/badge/Power%20Pivot-0078D4?style=for-the-badge)
![DAX](https://img.shields.io/badge/DAX-Measures-8A2BE2?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-In%20Progress-orange?style=for-the-badge)

![Dashboard Preview](images/dashboard.png)

---

## 📑 Table of Contents

- [🎯 Project Overview](#-project-overview)
- [🗂️ Repository Structure](#️-repository-structure)
- [🔄 Workflow at a Glance](#-workflow-at-a-glance)
- [🔍 Stage 1: Data Exploration](#-stage-1-data-exploration-in-excel)
- [🧹 Stage 2: Data Cleaning & Transformation](#-stage-2-data-cleaning--transformation-with-power-query)
- [🧩 Stage 3: Data Modeling](#-stage-3-data-modeling-with-power-pivot)
- [📈 Stage 4: Analysis with PivotTables](#-stage-4-analysis-with-pivottables)
- [📊 Stage 5: Dashboard](#-stage-5-excel-dashboard)
- [💡 Key Insights](#-key-insights)
- [🛠️ Tools & Skills](#️-tools--skills)
- [🚀 How to Use](#-how-to-use)
- [👤 Author](#-author)

---

## 🎯 Project Overview

| 📌 Item | 📝 Details |
|---|---|
| **Project Name** | HyperSales Hub |
| **Business Goal** | Understand sales performance by region, product, and customer to support commercial decisions |
| **Period Covered** | 2024 |
| **Scale** | 💰 $101.66M sales · 📦 108,591 units · 🧾 19,706 orders · 👥 300 customers |
| **Geography** | 5 regions: Mansoura, Aswan, Cairo, Giza, Alexandria |
| **Catalog** | 25 products (Laptop, Tablet, Jacket, Doll, Puzzle, …) |
| **Tools** | Excel, Power Query, Power Pivot, DAX, PivotTables, Slicers |
| **Deliverable** | Interactive Excel dashboard |

---

## 🗂️ Repository Structure

```text
📦 hypersales-hub
 ┣ 📂 data
 ┃ ┣ 📄 raw_data.xlsx            # Original dataset
 ┃ ┗ 📄 cleaned_data.xlsx        # Output after Power Query
 ┣ 📂 power_query
 ┃ ┗ 📄 cleaning_steps.m         # M code of all transformations
 ┣ 📂 images
 ┃ ┣ 🖼️ data_model.png           # Power Pivot schema
 ┃ ┣ 🖼️ pivots.png               # PivotTable analysis
 ┃ ┗ 🖼️ dashboard.png            # Final dashboard
 ┣ 📊 HyperSales_Hub.xlsx         # Final workbook
 ┗ 📄 README.md
```

---

## 🔄 Workflow at a Glance

| Stage | Icon | Phase | Tool | Output |
|:---:|:---:|---|---|---|
| 1 | 🔍 | Data Exploration | Excel | Data understanding & issues list |
| 2 | 🧹 | Cleaning & Transformation | Power Query | Clean, analysis-ready tables |
| 3 | 🧩 | Data Modeling | Power Pivot + DAX | Star schema & measures |
| 4 | 📈 | Analysis | PivotTables | Aggregated insights |
| 5 | 📊 | Visualization | Charts, KPI cards | Interactive dashboard |

```text
Raw Data ➜ Explore ➜ Clean (Power Query) ➜ Model (Power Pivot) ➜ Analyze (Pivots) ➜ Dashboard
```

---

## 🔍 Stage 1: Data Exploration in Excel

**Goal:** Understand the structure, quality, and content of the raw data before changing anything.

### 📋 Source Tables

| Table | Role | Key Columns | Description |
|---|:---:|---|---|
| 👥 **Customers** | Dimension | `CustomerID`, `Firstname`, `Lastname`, `Region` | 300 customers across 5 regions |
| 🏷️ **Products** | Dimension | `ProductID`, `ProductName`, `Category`, `Price` | 25 products with category and unit price |
| 🧾 **Orders** | Fact | `Order ID`, `CustomerID`, `ProductID`, `Quantity`, `Date`, `Sales` | One row per order line (transactions) |

### 🔎 What Was Checked

- ✅ Number of rows and columns in each table
- ✅ Data types of each column (especially dates and numbers)
- ✅ Missing / blank values
- ✅ Duplicate records and duplicate keys
- ✅ Outliers and inconsistent values
- ✅ Text formatting issues (spaces, casing)

### ⚠️ Issues Found

> 🚧 *Add the real issues you found while exploring (the table below is an example).*

| # | Issue | Column | Severity |
|:---:|---|---|:---:|
| 1 | `[e.g., Date stored as text]` | `Date` | 🔴 High |
| 2 | `[e.g., Extra spaces / inconsistent casing]` | `Region`, `ProductName` | 🟡 Low |
| 3 | `[e.g., Missing values]` | `[Column]` | 🟠 Medium |

---

## 🧹 Stage 2: Data Cleaning & Transformation with Power Query

**Goal:** Fix every issue found in Stage 1 using repeatable, documented **Applied Steps**.

> 🚧 *The M code file was empty when uploaded. Once it is sent, this section will be filled with the exact steps.*

### 🪜 Cleaning Steps Summary

| Step | Action | Purpose |
|:---:|---|---|
| 1 | `[Source]` | Load the raw data |
| 2 | `[Promote headers / Change type]` | Correct column names and data types |
| 3 | `[Remove duplicates / nulls]` | Improve data quality |
| 4 | `[Create "New Date" column]` | Standardized date used by the Calendar relationship |

### 💻 Full M Code

```powerquery
// Paste the final M code here
```

---

## 🧩 Stage 3: Data Modeling with Power Pivot

**Goal:** Build a relational model that makes analysis fast, flexible, and time-intelligent.

<img width="878" height="664" alt="Screenshot 2026-10-08 183000" src="https://github.com/user-attachments/assets/7065464b-54cc-449c-88c8-85865e74df52" />


### 🏗️ Model Design: Star Schema ⭐

The model follows a **star schema**: one central **fact table** (`Orders`) surrounded by three **dimension tables** (`Customers`, `Products`, `Calendar`). Filters flow from the dimensions down to the fact table, which keeps calculations simple and fast.

| Table | Type | Columns | Purpose |
|---|:---:|---|---|
| 🧾 **Orders** | 🟦 Fact | `Order ID`, `CustomerID`, `ProductID`, `Quantity`, `New Date`, `Sales` | Transactions and all DAX measures |
| 👥 **Customers** | 🟩 Dimension | `CustomerID`, `Firstname`, `Lastname`, `Region` | Slice by customer / region |
| 🏷️ **Products** | 🟩 Dimension | `ProductID`, `ProductName`, `Category`, `Price` | Slice by product / category |
| 📅 **Calendar** | 🟩 Dimension | `Date`, `Year`, `Month Number`, `Month`, `MMM-YYYY`, `Day Of Week Number`, `Day Of Week` | Time analysis, with a `Date Hierarchy` (Year ➜ Month ➜ Date) |

### 🔗 Relationships

| From (One side) | To (Many side) | Key | Cardinality |
|---|---|---|:---:|
| `Customers[CustomerID]` | `Orders[CustomerID]` | CustomerID | `1 : *` |
| `Products[ProductID]` | `Orders[ProductID]` | ProductID | `1 : *` |
| `Calendar[Date]` | `Orders[New Date]` | Date | `1 : *` |

### 📐 DAX Measures (in `Orders`)

| Measure | Purpose |
|---|---|
| 💰 `Total Sales` | Total revenue |
| 📦 `Total Quantity` | Total units sold |
| 🧮 `Avg Sales per Order` | Average order value |
| 👥 `Customer Count` | Number of distinct customers |
| 🧾 `Total Orders` | Number of orders |

> 💡 *Paste the DAX formulas here if you want them documented.*

---

## 📈 Stage 4: Analysis with PivotTables

**Goal:** Summarize the model to answer the key business questions.

<img width="917" height="502" alt="Screenshot 2026-10-08 201150" src="https://github.com/user-attachments/assets/0571c115-a4db-4f9c-a312-07876f32d363" />


### 🧮 Pivot: Product Performance (2024)

| Pivot Area | Fields |
|---|---|
| **Rows** | `ProductName` |
| **Columns** | `Year` (2024) |
| **Values** | `Total Sales`, `Total Quantity`, `Total Orders` |

This pivot feeds the product charts on the dashboard. Highlights:

| 🏆 Metric | Product | Value |
|---|---|---:|
| 💰 Highest sales | Doll | $8,155,620 |
| 💰 2nd highest sales | Shirt | $7,261,592 |
| 🧾 Most orders | Tablet | 848 |
| 📦 Highest quantity | Tablet | 4,677 |
| 🔻 Lowest sales | Sofa | $265,104 |
| 🔻 Lowest sales (2nd) | Headphones | $901,800 |

| ✅ Grand Total | Value |
|---|---:|
| Total Sales | $101,662,127 |
| Total Quantity | 108,591 |
| Total Orders | 19,706 |

### 🌍 Regional Analysis (feeds the pie charts)

| 🏙️ Region | 💰 Total Sales | 🧾 Total Orders |
|---|---:|---:|
| 🥇 Mansoura | $21,343,646 | 4,092 |
| Aswan | $20,714,471 | 4,050 |
| Cairo | $20,462,507 | 3,954 |
| Giza | $20,419,418 | 3,946 |
| 🔻 Alexandria | $18,722,085 | 3,664 |

---

## 📊 Stage 5: Excel Dashboard

**Goal:** Present the insights in one interactive, easy-to-read view.

<img width="1600" height="1149" alt="WhatsApp Image 2026-10-08 at 4 00 00 PM" src="https://github.com/user-attachments/assets/47d4a57e-7569-432a-b6c9-b67342a70318" />


| Component | Type | Insight |
|---|:---:|---|
| 🔢 **KPI Cards** | Cards | Total Sales, Total Quantity, Total Orders, Customer Count |
| 🥧 **Highest Sales vs. Lowest** | Pie chart | Sales by region: Mansoura highest 🟢, Alexandria lowest 🔴 |
| 🥧 **Highest Orders vs. Lowest** | Pie chart | Orders by region with the same highlighting |
| 📊 **Top 10 Order Leaders** | Bar chart | Products with the most orders |
| 💎 **Top 10 Best-Selling Products** | Bar chart | Products ranked by sales value |

### 🔢 KPI Summary

| 💰 Total Sales | 📦 Total Quantity | 🧾 Total Orders | 👥 Customer Count |
|:---:|:---:|:---:|:---:|
| **$101,662,127** | **108,591** | **19,706** | **300** |

### 🏅 Top 10 Rankings

| # | 🧾 By Orders | Orders | 💎 By Sales | Sales |
|:---:|---|---:|---|---:|
| 1 | Tablet | 848 | Doll | $8,155,620 |
| 2 | Table | 840 | Shirt | $7,261,592 |
| 3 | Headphones | 839 | Pants | $6,569,580 |
| 4 | Jacket | 832 | Shoes | $6,409,192 |
| 5 | Camera | 819 | Tablet | $6,407,490 |
| 6 | Pants | 811 | Car Toy | $6,307,520 |
| 7 | Ball | 809 | Puzzle | $6,027,520 |
| 8 | Shoes | 809 | Comics | $5,672,538 |
| 9 | Novel | 806 | Jacket | $5,647,701 |
| 10 | Chair | 801 | Laptop | $5,566,320 |

---

## 💡 Key Insights

1. 🥇 **Mansoura leads** in both sales ($21.34M) and orders (4,092), while **Alexandria trails** in both ($18.72M, 3,664 orders).
2. 📉 **Regions are close:** the top four regions sit within about 5% of each other in sales, so the gap is mainly Alexandria's underperformance.
3. 🧸 **Doll is the #1 product by revenue** ($8.16M) even though it is not in the top 10 by orders. It earns far more per unit than the order leaders.
4. 🎧 **Headphones is a top-3 product by orders (839) but has very low revenue ($901,800).** It is high-volume and low-value.
5. 💵 **Average order value is about $5,159**, and each order contains about 5.5 units on average.
6. 🛋️ **Sofa is the weakest product** with only $265,104 in sales, which makes it a candidate for review.

---

## 🛠️ Tools & Skills

| Tool | Usage |
|---|---|
| 📗 Microsoft Excel | Exploration, PivotTables, Dashboard |
| ⚡ Power Query | ETL: cleaning and transformation |
| 🧩 Power Pivot | Data modeling (star schema) |
| 🧮 DAX | KPI measures |
| 🎛️ PivotCharts & Slicers | Interactive visuals |

---

## 🚀 How to Use

1. 📥 Clone the repository
   ```bash
   git clone https://github.com/[username]/hypersales-hub.git
   ```
2. 📂 Open `HyperSales_Hub.xlsx` in Excel (2016 or later, with Power Pivot enabled)
3. 🔄 Go to **Data ➜ Refresh All** to reload the data
4. 🎛️ Explore the dashboard and PivotTables

---

## 👤 Author

**[Your Name]**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-0A66C2?style=flat&logo=linkedin&logoColor=white)](https://linkedin.com/in/your-profile)
[![GitHub](https://img.shields.io/badge/GitHub-181717?style=flat&logo=github&logoColor=white)](https://github.com/your-username)

⭐ *If you found this project useful, please give it a star!*
