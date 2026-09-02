<div align="center">

# -- ! Excel Data Analysis & Automated Analytics System ! --
### *Interactive Multi-Sheet Workbook for Student Performance, Employee Records & Sales Analysis*

[![Excel](https://img.shields.io/badge/Excel-Workbook-107C41?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/excel)
[![Formulas](https://img.shields.io/badge/Formulas-VLOOKUP%20%7C%20INDIRECT%20%7C%20XMATCH-FF6F00?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/excel)
[![Analytics](https://img.shields.io/badge/Analytics-Data%20Processing-4CAF50?style=for-the-badge&logo=pagekit&logoColor=white)](https://www.microsoft.com/excel)
[![Functions](https://img.shields.io/badge/Functions-Logical%20%26%20Statistical-9C27B0?style=for-the-badge&logo=python&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"Data is the new raw material of business — turn numbers into insights, and worksheets into intelligence."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🎓 Part A — Student Data Analysis](#-part-a--student-data-analysis)
- [💼 Part B — Employee Management & Dynamic Analysis](#-part-b--employee-management--dynamic-analysis)
- [📊 Part C — Sales Transaction & Discount Modeling](#-part-c--sales-transaction--discount-modeling)
- [🖼️ Screenshots](#️-screenshots)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Excel Data Analysis & Automated Analytics System** is a comprehensive, multi-tab Excel project designed to demonstrate advanced spreadsheet modeling, data lookup, logical condition handling, dynamic range calculations, and financial/academic performance analysis.

This project is structured around three dedicated modules:
- **Student Data Analysis**: Evaluates subject scores, determines pass/fail status using threshold logic, generates performance grades, and calculates conditional averages.
- **Employee Management**: Tracks workforce records across departments, enables dynamic column referencing using `INDIRECT` and `AVERAGE`, and implements exact lookup capabilities with `VLOOKUP` / `XLOOKUP`.
- **Sales Data Processing**: Models transactional data, calculates multi-tier discount percentages based on transaction thresholds, standardizes salesperson names via `UPPER`, and performs lookup operations with `XMATCH`.

---

## 🎯 Problem Statement

> **Objective:** Design an interactive, formula-driven Excel workbook to process student academic performance, employee salary records, and transactional sales metrics.

Modern organizational management requires streamlined tools to analyze heterogeneous data sources accurately without manual intervention. This Excel workbook automates evaluation routines across academic, human resources, and sales domain scenarios.

| 📂 Module | 📄 Sheet Name | 🔍 Description |
|-----------|--------------|----------------|
| Academic Performance | `Student_Data` | Evaluates subject marks, pass/fail state, grades, and conditional student averages |
| Human Resources | `Employees` | Manages employee details, dynamic column salary averages, and lookup profiles |
| Sales & Revenue | `Sales Data` | Automates multi-tiered discount brackets, name text formatting, and position matching |

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📊 **Multi-Tab Architecture** | Clean separation of business domains (`Student_Data`, `Employees`, `Sales Data`) |
| 🧮 **Automated Grading System** | Dynamic pass/fail classification and letter grades (A, B, C, D, F) based on subject score thresholds |
| 🔍 **Dynamic Lookup Queries** | Fast record retrieval using `VLOOKUP`, `XLOOKUP`, and `XMATCH` functions |
| 🔀 **Dynamic Range Referencing** | Advanced cell formula evaluation using `INDIRECT` and string concatenations |
| 🏷️ **Multi-Tier Discount Engine** | Nested `IF` statements categorizing order discounts from 0% up to 25% |
| 🧹 **Text Formatting & Cleaning** | Standardized string transformation using functions like `UPPER` |
| 🛡️ **Error Handling & Exclusions** | Prevents non-applicable calculations (e.g., suppressing averages for failing students) |

---

## 🏗️ Project Structure

```
📦 excel-analytics-project/
│
├── 📄 Project1.xlsx         ← Main Excel Workbook (Multi-sheet data model)
│
└── 📄 README.md             ← Project documentation
```

---

## 🔄 Project Workflow

```
Excel Workbook Opened (Project1.xlsx)
                 │
      ┌──────────┼──────────┐
      ▼          ▼          ▼
┌───────────┐┌───────────┐┌───────────┐
│Student_Data││ Employees ││Sales Data │
└─────┬─────┘└─────┬─────┘└─────┬─────┘
      │            │            │
      ▼            ▼            ▼
 ┌─────────┐  ┌─────────┐  ┌─────────┐
 │Subject  │  │Salary   │  │Discount │
 │Scores & │  │Records &│  │Tiers &  │
 │Grades   │  │INDIRECT │  │XMATCH   │
 └────┬────┘  └────┬────┘  └────┬────┘
      │            │            │
      └──────────┼──────────┘
                 ▼
      ┌────────────────────┐
      │ Interactive Lookup │
      │  & Summary Output  │
      └────────────────────┘
```

---

## 🎓 Part A — Student Data Analysis

### 📝 1. Academic Performance Modeling

The `Student_Data` worksheet records individual subject scores (Maths, English, Science, Statistics) and applies formula-based evaluation logic to calculate totals, results, and letter grades.

### 🗺️ 2. Core Logic & Formulas

| Metric | Formula / Logic Used |
|--------|----------------------|
| **Total Marks** | `=SUBTOTAL(9, Table1[[#This Row],[Column4]:[Column7]])` |
| **Pass/Fail Result** | `=IF(OR(D2<33, E2<33, F2<33, G2<33), "Fail", "Pass")` |
| **Grade Assignment** | `=IF(I2="Pass", IF((H2/4)>80,"A", IF((H2/4)>65,"B", IF((H2/4)>45,"C","D"))), "F")` |
| **Science & Maths Check**| `=IF(AND(F2>=80, D2>=80), "True", "False")` |
| **Maths > 50 Count** | `=COUNTIF(D2:D15, ">50")` |
| **Filtered Student Lookup**| `=VLOOKUP(10, A2:C15, 2, FALSE)` |

---

## 💼 Part B — Employee Management & Dynamic Analysis

### 🔍 3. HR Record Processing

The `Employees` worksheet maintains staff metadata including Employee ID, Full Name, Department, Joining Date, and Annual Salary.

### 🗺️ 4. Advanced Dynamic Referencing

- **Dynamic Range Salary Average**: Uses `=AVERAGE(INDIRECT(G3 & "2:" & G3 & "16"))` where cell `G3` holds the column identifier (e.g., `E`), enabling dynamic column range targeting without formula modification.
- **Employee Search**: Quickly retrieves full employee details using exact key matching.

---

## 📊 Part C — Sales Transaction & Discount Modeling

### 🏷️ 5. Sales Analytics Engine

The `Sales Data` worksheet tracks transactions across regions, salespersons, product items, order amounts, and discount rules.

### 🗺️ 6. Discount & Lookup Logic

| Operation | Formula / Logic | Description |
|-----------|-----------------|-------------|
| **Discount Calculation** | `=IF(F2<50,"Not Applicable", IF(F2<100,"5%", IF(F2<200,"10%", IF(F2<500,"15%", IF(F2<1000,"20%","25%")))))` | Tiered discount scaling based on order total |
| **Name Normalization** | `=UPPER(E2)` | Standardizes salesperson names to uppercase |
| **Position Matching** | `=_xlfn.XMATCH(L1, C2:C16, 0, 2)` | Locates product index/position within dataset |

---

## 🖼️ Screenshots

### 1. Student Data Analysis Sheet
![Student Data Worksheet](input_file_0.png)

### 2. Employee Records & Dynamic Lookup Sheet
![Employee Records Worksheet](input_file_1.png)

### 3. Sales Transactions & Discount Analysis Sheet
![Sales Data Worksheet](input_file_2.png)

---

## 🛠️ Tech Stack

| Tool / Function | Purpose |
|-----------------|---------|
| 📊 **Microsoft Excel** | Primary data modeling platform (.xlsx) |
| 🔍 **VLOOKUP / XLOOKUP** | Exact key matching and record retrieval |
| 🔀 **INDIRECT** | Dynamic range reference construction |
| 🎯 **XMATCH** | Dynamic array and position searching |
| 🧮 **SUBTOTAL / AVERAGE / COUNTIF** | Aggregation, filtering, and statistical analysis |
| 🔀 **IF / AND / OR** | Multi-branch logical evaluations |

---

## 📈 Results & Insights

- 🎓 **Academic Overview**: Automated pass/fail classification preventing grade inflation for students failing individual passing thresholds (<33).
- 💼 **Payroll & HR Insights**: Flexible salary analytics utilizing indirect column targeting for dynamic data dashboards.
- 📊 **Revenue & Sales Performance**: Clear bracketed discount metrics incentivizing higher transaction volumes alongside upper-case standardized salesperson tracking.

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| ⚡ **Automated Workflow** | Zero manual re-calculation required upon updating input values |
| 🛡️ **Error Resistant** | Handles conditional edge-cases like failing marks or dynamic range bounds |
| 📈 **Scalable Architecture** | Modular design makes adding new sheets or parameters straightforward |
| 📊 **Professional Formatting** | Clear tabular structure with visual separation of headers and calculated panels |

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Neev Shankar

> *"Data is the new raw material of business — turn numbers into insights, and worksheets into intelligence."*

**🎓 Role:** Junior Python Developer | Data Analyst Enthusiast **📍 Location:** India**🛠️ Skills:** Excel Analytics · Data Analysis · Python · Data Modeling · Logic Building

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Documentation](https://support.microsoft.com/en-us/excel) — Official Excel functions reference
- 📊 [Chandoo.org — Excel Formulas](https://chandoo.org/) — Advanced Excel and dashboard tutorials
- 🧮 [Exceljet](https://exceljet.net/) — Quick and precise formula guides
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: 27 May, 2026*

</div>
