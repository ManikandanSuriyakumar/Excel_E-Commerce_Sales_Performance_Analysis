# 🛒 E-Commerce Sales Performance Analysis

An end-to-end Excel analytics project that cleans, transforms, and analyzes e-commerce sales data to uncover sales trends, customer behavior, and performance gaps, and turns them into actionable business recommendations.

\---

## 📑 Table of Contents

* [Project Overview](#-project-overview)
* [Objectives](#-objectives)
* [Repository Structure](#-repository-structure)
* [Data Sources](#-data-sources)
* [Data Cleaning](#-data-cleaning)
* [Data Transformation](#-data-transformation)
* [Exploratory Data Analysis](#-exploratory-data-analysis)
* [Dashboard Insights](#-dashboard-insights)
* [Types of Analytics](#-types-of-analytics)
* [Recommendations](#-recommendations)
* [Conclusion](#-conclusion)
* [Tools \& Technologies](#-tools--technologies)
* [Key Learnings](#-key-learnings)
* [Author](#-author)
* [Acknowledgement](#-acknowledgement)

\---

## 📌 Project Overview

|Detail|Description|
|-|-|
|**Project Title**|E-Commerce Sales Performance Analysis|
|**Primary Tool**|Microsoft Excel|
|**Source Dataset**|`Ecommerce\\\_Sales\\\_Raw\\\_Dataset.xlsx`|
|**Analysis Period**|14 October 2023 – 13 October 2025|
|**Data Volume**|2,000 sales records · 500 customers · 100 products · 20 stores|
|**Scope**|Data preparation, descriptive analysis, Pivot Tables, dashboard interpretation, and recommendations|

This project combines sales transactions with customer, product, and store information to evaluate business performance. Excel is used for data cleaning and transformation, KPI calculation, exploratory data analysis (EDA), descriptive statistics, and an interactive dashboard built with Pivot Tables, Pivot Charts, and Slicers.

\---

## 🎯 Objectives

* Measure overall sales, order volume, units sold, unique customers, and average discount.
* Identify year-wise and month-wise sales patterns within the available date range.
* Compare performance across categories, regions, products, payment types, and loyalty levels.
* Prepare consistent, analysis-ready data by handling missing values and creating calculated fields.
* Build an interactive Excel dashboard and translate findings into actionable recommendations.

\---

## 📁 Repository Structure

```text
📦 Excel\\\_E-Commerce\\\_Sales\\\_Performance\\\_Analysis
 ┣ 📂 Dashboard
 ┃ ┗ 📄 E-Commerce\\\_Sales\\\_Performance\\\_Dashboard.xlsx                  # Main dashboard workbook (all stages)
 ┣ 📂 Cleaned\\\_Dataset
 ┃ ┗ 📄 Ecommerce\\\_Sales\\\_Data\\\_Cleaning\\\_and\\\_Transformation.xlsx        # Cleaned dataset
 ┣ 📂 Raw\\\_Dataset
 ┃ ┗ 📄 Ecommerce\\\_Sales\\\_Raw\\\_Dataset.xlsx                             # Raw dataset
 ┣ 📂 Analysis\\\_Report
 ┃ ┗ 📄 Ecommerce\\\_Sales\\\_Performance\\\_Analysis\\\_Report.pdf              # Project documentation
 ┣ 📂 Screenshots
 ┃ ┗ 📂 Ecommerce\\\_Pivot\\\_Table\\\_Charts
 ┃   ┗ 🖼️ \\\*.png                                                      # Pivot table and chart screenshots
 ┣ 🖼️ E-commerce\\\_Sales\\\_Dashboard.png                                 # Dashboard screenshot
 ┗ 📄 README.md                                                      # Markdown documentation
```

\---

## 🗂️ Data Sources

**Source file:** `Ecommerce\\\_Sales\\\_Raw\\\_Dataset.xlsx`
**Timeline:** 14 October 2023 – 13 October 2025

> ℹ️ 2023 covers October–December only, and 2025 covers January–October only.

|Dataset|Records|Fields|Purpose|
|-|-:|-:|-|
|`Sales\\\_DB`|2,000|10|Transaction-level sales and calculated amount|
|`Customer\\\_DB`|500|8|Customer profile and loyalty level|
|`Product\\\_DB`|100|7|Product, category, sub-category, cost, and stock|
|`Store\\\_DB`|20|5|Store and region attributes|

### 

### Key Fields

|Attribute|Type|Description|
|-|-|-|
|`Sales ID`|Text|Unique transaction identifier|
|`Order\\\_Date`|Date|Date of transaction|
|`Customer ID`|Text|Key linking sales to the customer table|
|`Product ID`|Text|Key linking sales to the product table|
|`Store ID`|Text|Key linking sales to the store table|
|`Quantity`|Numeric|Transaction quantity|
|`Unit\\\_Price`|Numeric|Transaction unit price|
|`Discount`|Numeric|Transaction discount rate|
|`Total\\\_Amount`|Numeric|Transaction amount|
|`Category` / `Sub\\\_Category`|Categorical|Product classification|
|`Region`|Categorical|Store geography|
|`Store\\\_Type`|Categorical|Store type|
|`Loyalty\\\_Level`|Categorical|Customer loyalty segment|

\---

## 🧹 Data Cleaning

**Cleaned workbook:** `Ecommerce\\\_Sales\\\_Data\\\_Cleaning\\\_and\\\_Transformation.xlsx`

Cleaned helper columns are kept alongside the original fields so imputed values can be compared with the source.

|Field|Missing|Treatment|
|-|-:|-|
|`Product\\\_DB` – Cost|6|Sub-category average (`AVERAGEIF`), rounded to 2 decimals|
|`Product\\\_DB` – Stock|3|Category and sub-category average (`AVERAGEIFS`), rounded to a whole number|
|`Sales\\\_DB` – Quantity|5|Derived from cleaned amount, unit price, and discount|
|`Sales\\\_DB` – Unit\_Price|3|Derived from amount, quantity, and discount|
|`Sales\\\_DB` – Discount|4|Derived from amount, unit price, and quantity|
|`Sales\\\_DB` – Total\_Amount|4|Calculated as `Quantity × Unit\\\_Price × (1 − Discount)`|

### 

### Imputation Formulas

```excel
' Product\\\_DB
Cleaned\\\_Cost  = IF(ISBLANK(\\\[@Cost]), ROUND(AVERAGEIF(\\\[Sub\\\_Category],\\\[@\\\[Sub\\\_Category]],\\\[Cost]),2), \\\[@Cost])
Cleaned\\\_Stock = IF(ISBLANK(\\\[@Stock]), ROUND(AVERAGEIFS(\\\[Stock],\\\[Category],\\\[@Category],\\\[Sub\\\_Category],\\\[@\\\[Sub\\\_Category]]),0), \\\[@Stock])

' Sales\\\_DB
Cleaned\\\_Quantity     = IF(ISBLANK(\\\[@Quantity]), ROUND(\\\[@\\\[Cleaned\\\_Total\\\_Amount]]/(\\\[@\\\[Cleaned\\\_Unit\\\_Price]]\\\*(1-\\\[@Discount])),0), \\\[@Quantity])
Cleaned\\\_Unit\\\_Price   = IF(ISBLANK(\\\[@\\\[Unit\\\_Price]]), \\\[@\\\[Cleaned\\\_Total\\\_Amount]]/(\\\[@Quantity]\\\*(1-\\\[@Discount])), \\\[@\\\[Unit\\\_Price]])
Cleaned\\\_Discount     = IF(ISBLANK(\\\[@Discount]), 1-(\\\[@\\\[Cleaned\\\_Total\\\_Amount]]/(\\\[@\\\[Cleaned\\\_Unit\\\_Price]]\\\*\\\[@\\\[Cleaned\\\_Quantity]])), \\\[@Discount])
Cleaned\\\_Total\\\_Amount = IF(ISBLANK(\\\[@\\\[Total\\\_Amount]]), \\\[@Quantity]\\\*\\\[@\\\[Unit\\\_Price]]\\\*(1-\\\[@Discount]), \\\[@\\\[Total\\\_Amount]])
```

**Additional checks**

* `Customer\\\_DB` and `Store\\\_DB` have no blank cells in the inspected fields.
* No fully duplicated rows were found in any of the four source tables.
* The cleaned helper columns contain no blanks for the fields listed above.

\---

## 🔄 Data Transformation

* Created `Order\\\_Year` using `YEAR(Order\\\_Date)`.
* Created `Order\\\_Month` using `TEXT(Order\\\_Date, "mmm-yyyy")`.
* Added `Cleaned\\\_Cost` and `Cleaned\\\_Stock` in `Product\\\_DB`.
* Added `Cleaned\\\_Quantity`, `Cleaned\\\_Unit\\\_Price`, `Cleaned\\\_Discount`, and `Cleaned\\\_Total\\\_Amount` in `Sales\\\_DB`.
* Applied `Total\\\_Amount = Quantity × Unit\\\_Price × (1 − Discount)` as the transaction amount relationship.
* Prepared KPI and Pivot Table summaries by year, month, category, region, product, payment type, and loyalty level.
* Dashboard KPI totals reconcile to the cleaned transaction amount: **2,174,394.70 across 2,000 orders**.

\---

## 📊 Exploratory Data Analysis

**Dashboard workbook:** `E-Commerce\\\_Sales\\\_Performance\\\_Dashboard.xlsx`

The dashboard contains **5 KPI measures** and **7 chart views** covering monthly sales, category, top products, region, year, payment type, and loyalty level.

### Key Performance Indicators

|KPI|Result|
|-|-:|
|Total Sales|2,174,394.70|
|Total Orders|2,000|
|Units Sold|4,964|
|Unique Customers|492|
|Average Discount|14.80%|

### 

### Descriptive Statistics (`Total\\\_Amount`)

|Statistic|Value|
|-|-:|
|Count|2,000|
|Sum|2,174,394.70|
|Mean|1,087.20|
|Median|868.72|
|Mode|No mode (all amounts are unique)|
|Standard Deviation (sample)|819.62|
|Minimum|18.46|
|Maximum|3,908.24|
|Range|3,889.78|
|Skewness|0.879|
|Kurtosis|0.046|

**Interpretation:** The mean is higher than the median and skewness is positive (0.879), indicating a **right-skewed distribution** in which a few high-value transactions pull the average upward.

\---

## 💡 Dashboard Insights

### 📈 Monthly Sales Trend

* Sales peaked in **March 2025 (107,588.86)** and were lowest in **October 2025 (38,467.82)**.
* Sales generally improved in 2025 versus 2024 in several months, with noticeable declines in February, June, and September.

### 🏷️ Category-wise Sales

* **Sports** led with **542,010.17 (24.93%)**, followed by **Electronics** at **504,311.35 (23.19%)**.
* **Beauty** recorded the lowest sales (**321,561.33**).
* Sports and Electronics together contribute **48.12%** of total sales.

### 🌍 Region-wise Sales

* **North** led with **646,100.88**, followed closely by **South (635,213.10)** and **East (622,404.32)**.
* **West** was lowest at **270,676.41 (12.45%)** and needs improvement.

### 📅 Year-wise Sales

* **2024** contributed the highest share (**\~50%**), followed by **2025 (39%)**.
* **2023** contributed only **10%**.
* Sales rose significantly from 2023 to 2024, then declined in 2025.

### 💳 Payment Type Analysis

* **PayPal** led with **756,754.57 (34.80%)**, followed by **Credit Card (715,180.71)**, with **COD** lowest at **702,459.42**.
* Sales are fairly evenly distributed across the three methods.

### 🏆 Top 10 Products

* **Portable Power Bank** is the top performer (**41,852.50**); **Soccer Training Cone Set** is the lowest of the top ten (**28,324.05**).
* The top ten together contributed **336,087.33 (\~15.46%)** of total sales, with no single product dominating.

### 🎖️ Loyalty Analysis

* **Platinum** is the largest segment (**141 customers; 28.66%**), followed by **Bronze (123)**; **Silver** is the smallest (**103**).
* Only **4** customers have an *Unknown* loyalty level.

\---

## 🧠 Types of Analytics

### 1\. Descriptive — *What happened?*

**Techniques:** KPIs, descriptive statistics, Pivot Tables, Pivot Charts

* **Overall:** 2.17M in sales across 2,000 orders, 4,964 units, and 492 unique customers; average order value of 1,087.20.
* **Year:** 2024 contributed the highest share (50.31%).
* **Category:** Sports led with 24.93%, followed by Electronics at 23.19%.
* **Region:** North was highest; West was lowest (12.45%).
* **Product:** The top ten products contributed 15.46% of sales.
* **Payment:** PayPal generated the highest sales value (34.80%).
* **Loyalty:** Platinum was the largest customer segment (28.66%).

### 2\. Diagnostic — *Why did it happen?*

**Techniques:** Comparative analysis across business dimensions

The dashboard identifies performance gaps but **does not establish their causes**. West's order volume, customer activity, store performance, and product availability require further investigation.

### 3\. Predictive — *What may happen?*

**Status:** No forecasting model is developed in this project. Monthly sales forecasting is a possible future enhancement, with accuracy validated against actual results.

### 4\. Prescriptive — *What should be done?*

**Approach:** Business recommendations based on observed findings (see below).

\---

## ✅ Recommendations

|Area|Recommendation|
|-|-|
|**Regional Performance**|Investigate the West region at store and product level by comparing order count, average order value, stock availability, and customer count before identifying causes.|
|**Category Performance**|Maintain availability and promotional support for Sports and Electronics; explore product mix, pricing, and conversion in Beauty.|
|**Sales Monitoring**|Use complete-month comparisons and clearly flag partial months and years to avoid misleading conclusions.|
|**Payment Analysis**|Compare both sales value and order count across payment methods.|
|**Profitability**|Incorporate cost, selling price, gross profit, and gross margin to evaluate profitability beyond revenue.|
|**Data Quality**|Review unknown loyalty records; maintain a data dictionary and imputation log.|

### 

### 🎯 SMART Framework Example — West Region

|Criterion|Action|
|-|-|
|**S**pecific|Improve sales performance in the West region.|
|**M**easurable|Use the current 12.45% contribution as a baseline and target an agreed improvement in sales or order count.|
|**A**chievable|Review store-level sales, availability, promotion coverage, and customer activity before choosing interventions.|
|**R**elevant|West is the lowest-contributing region on the dashboard.|
|**T**ime-bound|Run a 90-day pilot and compare against the preceding comparable 90-day period.|

\---

## 🏁 Conclusion

The analysis evaluates **2,000 e-commerce transactions** across customer, product, and store dimensions. Sports and Electronics, the North, South, and East regions, and PayPal transactions show strong contributions, while the **West region** records the lowest sales share. Transaction amounts follow a right-skewed distribution, and year-wise comparisons require caution because 2023 and 2025 are partial years.

Recommended next steps: deeper regional analysis, profitability measurement, consistent period comparisons, and validated forecasting.

\---

## 🛠️ Tools \& Technologies

* Microsoft Excel
* Power Query
* Advanced Excel formulas and structured table references
* Pivot Tables, Pivot Charts, and Slicers

\---

## 📚 Key Learnings

* **Data cleaning matters:** Imputing missing values with `AVERAGEIF`, `AVERAGEIFS`, and arithmetic relationships between fields produces analysis-ready data, but every assumption should be documented and validated.
* **Structured tables and formulas:** Structured table references and Power Query make cleaning and transformation repeatable and easier to audit.
* **Descriptive statistics reveal shape:** A mean above the median and positive skewness showed that a few high-value orders pull the average upward.
* **Interactive dashboards:** Pivot Tables, Pivot Charts, and Slicers turn raw transactions into KPIs and visuals that are easy to explore.
* **Partial periods need care:** 2023 and 2025 cover only part of each year, so comparisons should use complete periods to avoid misleading conclusions.
* **Descriptive is not diagnostic:** The dashboard shows *where* gaps exist (such as the West region) but not *why*, so further investigation is required before acting.
* **Beyond revenue:** Profit and gross-margin analysis, plus validated forecasting, are natural next steps.

\---

## 👤 Author

**Manikandan S**
Aspiring Data Analyst | Learning Excel and developing practical data analysis projects.

* 🔗 LinkedIn: https://www.linkedin.com/in/manikandan042/
* 💻 GitHub: https://github.com/ManikandanSuriyakumar
* 📧 Email: manisuriyakumar042@gmail.com

\---

## 📚 Acknowledgement

This Project is developed for educational and learning purposes to demonstrates end-to-end data analysis in Excel: data cleaning and transformation, descriptive analysis, formula-based imputation of missing values, and an interactive PivotTable-driven dashboard with KPIs and slicers.

\---

⭐ If you found this project helpful, consider giving the repository a star!

