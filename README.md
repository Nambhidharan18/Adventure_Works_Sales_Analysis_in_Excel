# Adventure Works Sales Analysis in Excel
*Sales Analysis Dashboard using Advance Excel*

## 📘 Introduction

This is a sales analysis project built in Excel using the Adventure Works dataset.
The project is a recreation for learning purposes, inspired by a YouTube tutorial series [link below](https://github.com/Nambhidharan18/Adventure_Works_Sales_Analysis_in_Excel/blob/main/README.md#-resources) .
It demonstrates end-to-end sales analysis using Excel, Power Query, and dashboards.

---
## 🎯 Objective

The objective of this **Excel dashboard** is to analyze **Four years of transactional data** to evaluate profit performance. The dashboard will use visualizations to track and compare **profit across products**, **customer locations** (countries), and **time periods**, enabling the identification of **trends**, **patterns**, and **key drivers of profitability**.

---
## 🗂️ Dataset

The dataset consists of **six related tables** containing transactional and dimensional data used to analyze business performance over time. The data follows a **star-schema structure**, with one fact table and multiple dimension tables, enabling detailed analysis of profit across products, customers, locations, and time.

### Tables Overview

*(table_name — Rows × Columns)*

- **FactInternetSales** *(60,389 × 13)* <br>
Contains detailed transactional sales records, including order quantities, costs, product costs, and order dates. This table serves as the primary source for profit calculations and time-based analysis.

- **DimSalesTerritory** *(11 × 5)* <br>
Provides sales territory information, including country-level data used to analyze performance across different customer locations.

- **DimProduct** *(606 × 3)* <br>
Stores product details such as product name and color, enabling product-level and category-based profit analysis.

- **DimGeography** *(655 × 4)* <br>
Includes geographic attributes linked to customers, supporting location-based analysis and regional insights.

- **DimCustomer** *(2,191 × 8)* <br>
Contains customer demographic information such as age and gender, allowing for customer segmentation and behavioral analysis.

- **DimDate** *(18,484 × 8)* <br>
Holds calendar attributes related to order dates, facilitating time-series analysis, trend identification, and year-over-year comparisons.

### Important Columns

The following key columns are used extensively in the analysis and dashboard visualizations:

- **OrderQuantity**           – Number of units sold per transaction
- **Cost / Product Cost**     – Cost incurred per product, used for profit calculation
- **OrderDate**               – Transaction date, linked to the DimDate table
- **Product Name**            – Identifies individual products
- **Color**                   – Product attribute used for grouping and comparison
- **SalesTerritoryCountry**   – Customer country for geographic analysis
- **Customer Age**            – Used for demographic segmentation
- **Gender**                  – Enables gender-based performance analysis

---


## 📌 Methodology

### 📥 Data Loading & Integration

Data was imported into Excel using **Power Query (Get Data)**.
All tables were loaded using the **Create Connection Only** option and added to the **Excel Data Model** to enable efficient data modeling and analysis.

**Tables Used:**

* `FactInternetSales` (Fact Table)
* `DimProduct`
* `DimSalesTerritory`
* `DimDate`
* `DimCustomer`
* `DimGeography`

Power Query was used as the primary ETL tool for data transformation, cleaning, and preparation.

---

### 🧹 Data Cleaning & Preparation

Data cleaning and preparation were performed **table by table** to ensure consistency, accuracy, and analytical readiness.



#### 🔹 FactInternetSales

* Retained only relevant columns required for analysis, including product, customer, date, territory keys, quantity, pricing, and cost fields
* Renamed `ProductStandardCost` to **Cost**
* Corrected data types across all columns
* Created calculated columns:

  * **Total Revenue** = `OrderQuantity × UnitPrice`
  * **COGS (Cost of Goods Sold)** = `OrderQuantity × Cost`
  * **Total Profit** = `Total Revenue − COGS`
* Formatted all financial fields as **Currency**
* Created **Product Price Type**:

  * *Less Expensive* (≤ 150)
  * *Expensive* (> 150)



#### 🔹 DimSalesTerritory

* Removed non-analytical column (`SalesTerritoryImage`)
* Filtered out records containing null values



#### 🔹 DimProduct

* Selected relevant columns such as Product Key, Product Name, and Color
* Renamed `EnglishProductName` to **Product Name**
* Replaced `"NA"` values in the Color column with **“Unspecified”**



#### 🔹 DimGeography

* Selected Geography, City, Country, and Sales Territory attributes
* Renamed `EnglishCountryRegionName` to **Country**



#### 🔹 DimCustomer

* Created **Full Name** by merging First Name and Last Name
* Selected key customer attributes including demographic and geographic information
* Calculated **Customer Age** using Birth Date and the current date
* Converted age values to integer format
* Created **Age Group** categories:

  * 25–30
  * 31–35
  * 36–40
  * 41–45
  * 46–50
  * 50+



#### 🔹 DimDate

* Retained the primary date field and renamed it to **Date**
* Created derived date attributes:

  * **Year** (filtered to 2009–2010)
  * **Month Number**
  * **Month Name** (abbreviated)
  * **Day Name** (abbreviated)
  * **Day of Week Number**
* Classified dates into **Weekday** and **Weekend**

---

### 🗄️ Data Modeling

A **star schema** was implemented with `FactInternetSales` as the central fact table connected to all dimension tables.

**Model Characteristics:**

* Many-to-one (*:1) relationships
* All relationships set as **active**
* Single-directional filtering from dimension tables to the fact table
* Optimized for accurate aggregations and dashboard performance

**Key Relationships:**

* `DimCustomer → DimGeography`
* `FactInternetSales → DimCustomer`
* `FactInternetSales → DimDate`
* `FactInternetSales → DimProduct`
* `FactInternetSales → DimSalesTerritory`


---

### 📈 Data Analysis


This project presents a structured data analysis performed using **Microsoft Excel**, with **Pivot Tables** as the primary analytical tool.
The objective of the analysis is to evaluate overall business performance, identify time-based trends, and conduct detailed product and customer-level deep-dive analyses.

The analysis is organized into **four worksheets**, each addressing a specific analytical objective:

* **Analysis-1 & Analysis-2:** Time-series and performance trend insights
* **Product Analysis & Customer Analysis:** In-depth dashboards for product and customer behavior


#### Sheet Overview

**number_of_sheets:** 4
**titles:**

* Analysis-1
* Analysis-2
* Product Analysis
* Customer Analysis

Each worksheet focuses on a specific analytical objective and collectively supports descriptive, trend-based, comparative, segmentation, and contribution analysis using Excel Pivot Tables.

---

#### 🔹 Analysis-1

##### Analytical Techniques Applied

* Descriptive Analysis
* Trend Analysis
* Comparative Analysis
* Contribution Analysis

##### Analysis Description

This sheet provides a **high-level business performance overview**. Pivot tables were used to aggregate core metrics such as transactions, revenue, cost, profit, product coverage, and margins.
Time-based comparisons were then performed to evaluate year-over-year performance across key financial and operational indicators.

##### Key Findings

* The business recorded **60.4K transactions**, generating **$307.09M in revenue** and **$126.29M in profit**, resulting in a strong **profit margin of 41.1%**, indicating healthy operational efficiency.
* Out of **606 available products**, only **158 were sold**, leaving **448 unsold products**, highlighting a significant opportunity for product portfolio optimization.
* Revenue and profitability showed **consistent year-over-year growth from 2005 to 2007**, with revenue increasing from **$33M to over $102M**.
* Transaction volume surged sharply in 2007 and 2008, suggesting improved market penetration or expanded sales channels.
* Profit margins remained stable (around 40–42%) across all years, reflecting controlled costs despite rapid growth.
* While 2008 maintained high revenue and profit levels, growth plateaued, indicating potential market saturation or the need for strategic expansion.

---

#### 🔹 Analysis-2

##### Analytical Techniques Applied

* Trend Analysis
* Contribution Analysis
* Time-Based Pattern Analysis

##### Analysis Description

This sheet focuses on **temporal profitability patterns**. Profit was analyzed across months, days of the week, weekdays vs weekends, and quarters to identify seasonality and behavioral trends.
All analyses were performed using aggregated pivot tables across the full dataset.

##### Key Findings

* Monthly profit trends indicate **strong performance in Q2 and Q4**, with **May and December** emerging as peak profit months, suggesting seasonal demand patterns.
* Profit distribution across weekdays is relatively balanced; however, **Thursdays and Fridays** show marginally higher profitability.
* **Weekdays contribute 72% of total profit**, compared to **28% from weekends**, indicating that core business activity is concentrated during working days.
* Quarter-wise analysis shows:

  * **Q2 as the strongest quarter**, contributing **31% of total profit**
  * **Q3 as the weakest**, contributing only **19%**, highlighting potential opportunities for targeted campaigns or promotions during this period
* Overall, the analysis confirms **clear seasonality and time-based concentration of profits**, useful for planning marketing and inventory strategies.

---

#### 🔹 Product Analysis

##### Analytical Techniques Applied

* Comparative Analysis
* Segmentation Analysis
* Contribution Analysis

##### Analysis Description

This sheet delivers a **deep dive into product-level performance**. Products were compared based on profitability, segmented by attributes such as color and price category, and analyzed for their contribution to total profit.

##### Key Findings

* The **top 5 products alone contribute 24.8% of total profit**, indicating high revenue concentration among a small subset of products.
* These top-performing products are predominantly from the **Mountain-200 series**, highlighting strong customer preference for this product line.
* Product color analysis reveals:

  * **Black** as the most profitable color
  * Followed by **Red** and **Silver**
* Pricing segmentation shows that **Expensive products contribute 95.4% of total profit**, confirming that profitability is driven primarily by premium offerings rather than volume-based low-cost products.
* The results suggest opportunities to:

  * Rationalize underperforming products
  * Strengthen premium product positioning
  * Optimize inventory around high-margin segments

---

#### 🔹 Customer Analysis

##### Analytical Techniques Applied

* Comparative Analysis
* Segmentation Analysis
* Contribution Analysis

##### Analysis Description

This sheet analyzes **customer-level profitability and demographics**. Customers were segmented by age group, gender, and geography to identify high-value segments and assess revenue concentration risks.

##### Key Findings

* The **top 5 customers contribute only 0.3% of total profit**, indicating a **well-diversified customer base** with low dependency on individual customers.
* Gender-based analysis shows an almost equal contribution:

  * **Female customers:** 50.4%
  * **Male customers:** 49.6%
* Age-based segmentation reveals that the **50+ age group contributes 38.9% of total profit**, making it the most valuable demographic segment.
* Country-level analysis shows that **the United States and Australia together account for 62.7% of total profit**, highlighting geographic revenue concentration.
* These insights support targeted marketing strategies focused on:

  * Older customer segments
  * High-performing regions
  * Balanced customer acquisition rather than reliance on a few large customers

---

### ✅ Summary

The Data Analysis phase applied structured, multi-dimensional analysis using Excel Pivot Tables to uncover trends, performance drivers, and concentration risks across **time, products, and customers**.
The findings provide a strong foundation for data-driven decision-making and dashboard storytelling.



---

## 📊 Data Visualization

This project includes two interactive Excel dashboards designed to analyze profit performance across time, products, customers, and demographics.

### 1. Time Series Dashboard

Focuses on performance trends and time-based analysis.

* **KPI Comparison (YoY):** Comparison of COGS, Revenue, Quantity, Profit, Profit Margin, and Transactions against the previous year.
* **Above-Average Year Performance:** Total Revenue, Profit, and Transactions for years exceeding average performance.
* **Monthly Profit Trends:** Analysis of profit trends on a monthly basis.
* **Profit by Week Type:** Profit comparison across different week types.
* **Quarterly Profit Analysis:** Evaluation of profit performance by quarter.
* **Profit by Weekday:** Analysis of profit trends across weekdays.


### 2. Detail Dashboard

Provides detailed insights into profitability across products, customers, demographics, pricing, and geography.

* **Top 5 Profitable Products:** Percentage contribution of the top five products versus others.
* **Top 5 Profitable Customers:** Percentage contribution of top customers compared to others.
* **Profit by Gender:** Profit distribution by gender.
* **Profit by Product Color:** Identification of best-selling and most profitable colors.
* **Profit by Pricing Types:** Profit analysis across different pricing strategies.
* **Country-wise Profit:** Geographic visualization of profit by country.
* **Profit by Age Groups:** Profit contribution segmented by age groups.


These dashboards provide clear, visual insights into revenue, customers, products, and trends.

---

## 🖼️ Screenshots

**Time Series Dashboard**

<img width="1191" height="618" alt="Time series analysis" src="https://github.com/user-attachments/assets/d8417e08-e39b-4e7e-ab95-535d3bf7223b" />


**Detail Dashboard**

<img width="1169" height="566" alt="Detail Dashboard" src="https://github.com/user-attachments/assets/697ab131-bb5e-4be9-b8fb-c87bbb678f3c" />

**Data Modelling( Star Shema)**

<img width="971" height="593" alt="Data model" src="https://github.com/user-attachments/assets/adf7ca7b-da28-4cb2-8524-a3472f3c039a" />



---

## 🛠️ Skills Used

### Excel

* Pivot Tables for aggregation, comparison, and trend analysis
* Pivot Charts for visualizing time-based and categorical patterns
* Slicers for interactive filtering and exploratory analysis
* Conditional Formatting to highlight key metrics and performance variances

### Power Query (ETL)

* Data cleaning and preprocessing across multiple tables
* Data transformation and column derivation
* Data type standardization and validation
* Query-based data preparation for analytical readiness

### Data Modeling

* Building and managing relationships between fact and dimension tables
* Implementing a star schema for efficient analysis
* Using the Excel Data Model for scalable reporting

### Dashboard Design & Visualization

* KPI cards to summarize key business metrics
* Interactive filters for dynamic insights
* Visual layout optimization for clarity and usability

### Sales & Business Analytics

* Revenue, cost, and profit analysis
* Profit margin and contribution analysis
* Product performance and segmentation analysis
* Customer demographic and geographic analysis
* Time series and seasonality analysis

---
## Folder Structure


```
[Adventure_Works_Sales_Analysis_in_Excel]/
├── Dashboard
|   └── Adventure Works Sales Dashboard
├── Database
|   └──  AdventureWorks
├── Images/
│   └── analytics
|   └── profit
|   └── barcode (1)
|   └── ready-stock (1)
|   └── barcode
|   └── ready-stock
|   └── calendar
|   └── sale
|   └── D-profit
|   └── earning
|   └── sales
|   └── growth
|   └── shopping-cart
|   └── in-stock
└── README.md
```

___

## ✅ Conclusion

The final dashboards deliver **clear, interactive, and insight-driven views** of Adventure Works’ sales performance across time, products, and customers. By integrating structured data modeling with Pivot Tables and Power Query, the analysis enables effective exploration of key business metrics such as revenue, profitability, product contribution, and customer behavior.

This project strengthened practical skills in **Excel-based data analysis**, **ETL using Power Query**, and **dashboard design focused on usability and decision-making**. The resulting dashboards are designed to support business users in identifying trends, performance drivers, and areas for strategic improvement.

For those interested in recreating or extending this project, reference materials and learning resources are provided below.

---

## 🎥 Resources 

YouTube tutorial series by the original author:

Part 1: https://youtu.be/VxOOt2dP8Jw?si=okSncDr4spyx2NxO

Part 2: https://youtu.be/sJlVb32jyoQ?si=8ZCmzqgsUT7sDuHk

Part 3: https://youtu.be/LKwyKSw6PhU?si=gW-HOMf8zcBvHi2w

Part 4: https://youtu.be/a1OF_wgRK_U?si=lU2eQ-0mCcuvPGLi
