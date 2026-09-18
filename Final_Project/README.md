<div align="center">

# -- ! Sales DATA Analytics Dashboard ! --
### *Interactive Excel Analytics, High-Value Customer Profiling & Data-Driven Scenario Modeling*

[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Pivot Tables](https://img.shields.io/badge/Pivot_Tables-Data_Aggregation-107C41?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Dashboard](https://img.shields.io/badge/Dashboard-Executive_Reporting-0078D4?style=for-the-badge&logo=powerbi&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Statistics](https://img.shields.io/badge/Analytics-Regression_%26_ANOVA-FFB900?style=for-the-badge&logo=python&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)

<br/>

> *"Data isn't just numbers on a grid — structure it well, and it drives strategy."*

</div>

---

## 📋 Table of Contents

- [-- ! Sales DATA Analytics Dashboard ! --](#--sales-data-analytics-dashboard--)
  - [📋 Table of Contents](#-table-of-contents)
  - [📌 Overview](#-overview)
  - [🎯 Problem Statement](#-problem-statement)
  - [✨ Key Features](#-key-features)
  - [🏗️ Project Structure](#️-project-structure)
  - [🔺 Part A — Executive Dashboard \& Visual Analytics](#-part-a--executive-dashboard--visual-analytics)
    - [📝 1. Dashboard Architecture](#-1-dashboard-architecture)
    - [🗺️ 2. Core Dashboard KPI Summary](#️-2-core-dashboard-kpi-summary)
  - [🔢 Part B — Master Dataset, Pivots \& High-Value Analytics](#-part-b--master-dataset-pivots--high-value-analytics)
    - [🔍 3. Transaction Master Log](#-3-transaction-master-log)
    - [📊 4. Pivot Tables \& Categorical Breakdown](#-4-pivot-tables--categorical-breakdown)
    - [👑 5. High-Value Customer Profile \& KPI Summaries](#-5-high-value-customer-profile--kpi-summaries)
    - [🧮 6. Statistical Regression Analysis](#-6-statistical-regression-analysis)
  - [🛠️ Tech Stack](#️-tech-stack)
  - [📈 Results \& Insights](#-results--insights)
  - [🏆 Key Business Advantages](#-key-business-advantages)
  - [📄 License](#-license)
  - [👤 Author](#-author)
  - [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Sales DATA Analytics Dashboard** is a comprehensive Microsoft Excel analytical workbook and executive reporting project designed to deliver deep visibility into multi-year transactional records, product category performance, payment gateway preferences, customer tenure dynamics, and revenue regression dynamics.

This project is structured to:
- Process granular, order-level transactional data ($250$ orders across FY-2024 and FY-2025).
- Aggregate monthly sales performance across categories (Appliances, Electronics, Furniture).
- Evaluate customer tenure trends and isolate top revenue-generating accounts.
- Deploy linear regression analysis to evaluate statistical relationships across transaction variables.

---

## 🎯 Problem Statement

> **Objective:** Build an interactive executive sales dashboard and statistical analytics model in Microsoft Excel to monitor multi-year sales, evaluate category performance, profile high-value customers, analyze payment channel distributions, and perform regression modeling.

Enterprise retail platforms handle high transaction volumes across multiple product lines, regions, and payment options. Raw data alone obscures key trend lines, critical sales drop-offs, product performance disparities, and payment preferences.

| 📂 Module | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| **Executive Dashboard** | Visual Interface | Multi-chart executive hub with core KPI cards and trend visuals |
| **Transaction Master Log** | Data Engine | Clean dataset containing 250 records with temporal and tenure fields |
| **Pivot Table Aggregations** | Summarization | Multi-perspective tables breaking down counts, totals, and monthly trends |
| **Customer Profiling & KPIs** | Strategic Analytics | Isolates top spenders, customer tenure, and category order volumes |
| **Statistical Regression** | Quantitative Analysis | Linear regression summary evaluating statistical significance and ANOVA |

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📊 **Dynamic Executive Dashboard** | Summary view featuring top KPI banners, category bar charts, product pie distribution, and payment trend lines |
| 🏷️ **Executive KPI Banners** | Instant evaluation of Sales in FY-2024 ($1,60,444.73), FY-2025 ($68,747.74), Average Sales ($916.77), Most Ordered Product (Bookshelf), and Top Payment Method (Cash) |
| 🍕 **3D Product Sales Pie Distribution** | Visual breakdown showing product contributions (Bookshelf 13%, Smartphone 13%, Blender 12%, Monitor 12%, Keyboard 12%, etc.) |
| 📈 **Monthly Payment Gateway Line Trends** | Multi-line temporal visualization tracking Cash, Credit Card, Debit Card, and PayPal usage over 13 months |
| 👥 **High-Value Customer Profiling** | Identifies top revenue contributor (Mark Carter with **$15,659.65**) and tracks total individual customer spend |
| 🧪 **Regression Analysis & ANOVA** | Performs linear regression modeling with R Square ($0.2098$), F-significance ($2.25 \times 10^{-14}$), and detailed p-value outputs |

---

## 🏗️ Project Structure

```text
Sales-DATA-Analytics-Dashboard/
│
├── README.md                           # Main project documentation
├── Final_Project.xlsx                  # Core Excel analytical workbook
└── screenshots/
    ├── Screenshot 2026-09-18 155332.png # Sales DATA Analytics Dashboard UI
    ├── Screenshot 2026-09-18 155350.png # Master Transaction Log Dataset
    ├── Screenshot 2026-09-18 155404.png # Multi-Perspective Pivot Tables
    ├── Screenshot 2026-09-18 155414.png # Summary Output Regression Analysis
    └── Screenshot 2026-09-18 155428.png # High-Value Customer & KPI Summary
```

---

## 🔺 Part A — Executive Dashboard & Visual Analytics

### 📝 1. Dashboard Architecture

The main executive dashboard aggregates transactional metrics into polished KPI summary cards along with three distinct visual perspective charts: Category Monthly Breakdown, Product Sales Percentage Contribution, and Payment Method Usage Lines.

![Sales DATA Analytics Dashboard](Final_Project/Dashboard.png)

---

### 🗺️ 2. Core Dashboard KPI Summary

| KPI Metric | Value | Insights |
|------------|-------|----------|
| 📅 **Sales in FY-2024** | **$ 1,60,444.73** | Principal baseline revenue across 178 orders |
| 📉 **Sales in FY-2025** | **$ 68,747.74** | Secondary period sales across 72 orders |
| 💵 **Average Sales** | **$ 916.77** | Mean gross transaction value per order |
| 📦 **Most Ordered Product** | **Bookshelf** | Top product volume leader ($102$ total units ordered) |
| 💳 **Most Used Payment Method** | **Cash** | Dominant payment method ($76$ total orders) |
| 💰 **Total Combined Revenue** | **$ 2,29,192.47** | Cumulative sales revenue across all 250 orders |

---

## 🔢 Part B — Master Dataset, Pivots & High-Value Analytics

### 🔍 3. Transaction Master Log

> Raw granular order records incorporating Transaction IDs, Customer Details, Product IDs, Categories, Unit Prices, Payment Channels, Customer Tenure, and Computed Totals.

![Master Transaction Log](Final_Project/Transaction.png)

**Dataset Characteristics:**
- **Total Records:** 250 transactions across 50 unique customer accounts (`CUST001` - `CUST050`).
- **Product Portfolio:** 10 products (`P001` - `P010`) covering Electronics, Appliances, and Furniture.
- **Regions Covered:** North, South, East, West, Central.

---

### 📊 4. Pivot Tables & Categorical Breakdown

> Multi-perspective pivot tables summarizing monthly category counts, product quantity metrics, payment method distributions, and revenue totals.

![Pivot Tables Summary](Final_Project/Pivot_Tables.png)

**Pivot Insights Breakdown:**

| Category | Order Count | Total Revenue ($) | Top Contributing Product |
|----------|-------------|-------------------|--------------------------|
| **Electronics** | **131** | **$1,71,756.05** | Laptop ($67,499.25) & Smartphone ($67,199.04) |
| **Furniture** | **76** | **$49,038.68** | Desk ($23,399.22) & Bookshelf ($15,298.98) |
| **Appliances** | **43** | **$8,338.74** | Blender ($5,219.13) |
| **Grand Total** | **250** | **$2,29,192.47** | **10 Core Products** |

---

### 👑 5. High-Value Customer Profile & KPI Summaries

> Highlights top spending customer accounts and summarizes critical macro KPIs across categories, payment channels, and monthly peaks.

![High-Value Customer & KPI Summary](Final_Project/High-Value_Customer.png)

**Key Analytics Summary:**
- **👑 High-Value Customer:** **Mark Carter** (`CUST025`) with **$15,659.65** in total spend.
- **Top Spenders List:** Edward Mitchell ($11,919.77), Barbara Young ($10,649.80), Patricia Moore ($9,799.73), Paul Baker ($8,309.74), and Dorothy Nelson ($8,269.79).
- **Highest Sales Month:** **December 2024** ($23,039.30).
- **Payment Method Split:** Cash (Highest), Credit Card, Debit Card, PayPal (Least Used).

---

### 🧮 6. Statistical Regression Analysis

> Linear regression output modeling statistical relationships across sales transaction observations ($N = 250$).

![Regression Analysis Output](Final_Project/Regression.png)

**Regression Model Metrics:**
- **Multiple R:** $0.45808$
- **R Square:** $0.20984$
- **Adjusted R Square:** $0.20665$
- **Standard Error:** $1.19068$
- **Observations:** $250$
- **ANOVA Significance F:** $2.25865 \times 10^{-14}$ (Statistically highly significant)
- **X Variable 1 Coefficient:** $0.000593$ ($t\text{-stat} = 8.115$, $p\text{-value} = 2.25 \times 10^{-14}$)

---

## 🛠️ Tech Stack

| Tool / Layer | Technology | Application |
|--------------|------------|-------------|
| 📊 **Core Engine** | Microsoft Excel | Data storage, dynamic formulas, VLOOKUP / INDEX-MATCH |
| 🔄 **Data Aggregation** | Excel Pivot Tables | Category summarization, monthly time-series groupings |
| 🎛️ **Visualization** | Excel Charts | Bar charts, 3D pie distribution, multi-line trend charts |
| 🧪 **Statistical Analytics** | Analysis ToolPak | Linear regression summary, ANOVA table calculation |
| 🎨 **UI / UX Design** | Custom Formatting | Card-style KPI headers, visual contrast panels |

---

## 📈 Results & Insights

- 🟢 **Primary Revenue Driver:** **Electronics** is the dominant category, generating **$1,71,756.05** (74.9% of total company revenue). Laptops and Smartphones account for the vast majority of this segment.
- 📦 **High Volume, Lower Margin:** **Bookshelf** is the most ordered item ($102$ units), but ranks behind high-ticket electronics in gross sales dollars.
- 🗓️ **Seasonal Peak:** **December 2024** was the single highest sales month ($23,039.30), reflecting strong end-of-year seasonal purchasing.
- 💳 **Payment Preferences:** **Cash** and **Credit Cards** represent the bulk of transaction processing, while PayPal remains the least utilized payment channel.

---

## 🏆 Key Business Advantages

| Feature | Advantage |
|---------|-----------|
| 💼 **Executive Readability** | High-level KPI cards allow executives to grasp performance in seconds |
| 📊 **Rigorous Statistics** | In-depth regression analysis ensures decision-making is backed by empirical data |
| 🎯 **Targeted Marketing** | Clear high-value customer identification allows for focused loyalty programs |
| ⚡ **Fully Native** | Built entirely within Microsoft Excel without requiring third-party plugins |

---

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).

---

## 👤 Author

<div align="center">

### Data Analytics Team

> *"Data transformed into structure becomes insight; insight transformed into strategy becomes growth."*

**🛠️ Specialization:** Excel Dashboards · Data Modeling · Business Intelligence · Quantitative Analytics

</div>

---

## 🙏 Acknowledgements

- 📚 [Microsoft Excel Documentation](https://support.microsoft.com/en-us/excel) — Charting & Analysis ToolPak guidance
- 📊 [Exceljet](https://exceljet.net/) — Formula references & visual layout inspiration
- 🎓 Analytical Best Practices — Standard statistical linear modeling frameworks

---

<div align="center">

---

*Sales DATA Analytics Dashboard — Excel Analytics Project*

</div>
