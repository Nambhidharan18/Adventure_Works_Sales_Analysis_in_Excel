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

---

#### Sheet Structure Summary

* **number_of_sheets:** 4
* **sheet_titles:**

  * Analysis-1
  * Analysis-2
  * Product Analysis
  * Customer Analysis

---

#### 1.Analysis-1

##### Analytical Techniques Applied

* Descriptive Analysis
* Trend Analysis
* Comparative Analysis

###### Analysis Description

The *Analysis-1* worksheet provides a consolidated view of overall business performance. Pivot tables were used to summarize key performance indicators, including transaction volume, revenue, cost of goods sold (COGS), profit, order quantity, and profit margins.

Year-wise aggregations enabled trend identification and comparative evaluation of performance across different periods. Additional product-level summaries, such as sold and unsold products, were included to assess sales effectiveness and inventory movement.

###### Key Findings

* Business performance exhibited noticeable year-over-year variations in revenue, profit, and transaction volume.
* Profit margins fluctuated across years, indicating changes in cost structure, pricing strategies, or product mix.
* A gap between total products and sold products highlighted opportunities for improved inventory utilization.
* Higher transaction volumes were strongly associated with increased revenue and profit levels.

---

#### 2.Analysis-2

##### Analytical Techniques Applied

* Time-Series Analysis
* Trend Analysis
* Contribution Analysis

###### Analysis Description

The *Analysis-2* worksheet focuses on profit behavior across multiple time dimensions. Pivot tables were used to analyze profit trends at the monthly, daily, and quarterly levels.

Additionally, profit contribution was evaluated by categorizing transactions into weekdays and weekends and by calculating each quarter’s percentage contribution to total profit. This approach enabled identification of seasonal patterns and operational performance across time periods.

###### Key Findings

* Monthly profit analysis revealed recurring seasonal trends, with certain months consistently generating higher profit.
* Daily profit patterns showed variability, suggesting differences in customer activity throughout the week.
* Weekdays contributed a larger share of total profit compared to weekends, indicating stronger business activity during standard operating days.
* Profit contribution was uneven across quarters, with specific quarters accounting for a significant portion of annual profitability.

---

#### 3.Product Analysis

##### Analytical Techniques Applied

* Comparative Analysis
* Contribution Analysis
* Segmentation Analysis

###### Analysis Description

The *Product Analysis* worksheet delivers a detailed assessment of product-level profitability. Pivot tables were used to compare profit contributions across individual products and identify high-performing items.

Products were further segmented by attributes such as color and price category (expensive vs. less expensive). This segmentation allowed evaluation of how product characteristics and pricing tiers influence overall profitability.

###### Key Findings

* A small subset of products contributed a disproportionately large share of total profit.
* The top-performing products significantly outperformed the remaining product portfolio in terms of profit contribution.
* Profitability varied across product colors, indicating differing customer preferences or demand patterns.
* Expensive products generated higher profit per unit, while less expensive products contributed through higher sales volume.

---

#### 4. Customer Analysis

##### Analytical Techniques Applied

* Descriptive Analysis
* Comparative Analysis
* Segmentation Analysis
* Contribution Analysis

###### Analysis Description

The *Customer Analysis* worksheet focuses on understanding customer profitability and demographic patterns. Pivot tables were used to analyze customer profit contribution alongside demographic attributes such as age, gender, and geographic location.

Customers were grouped into top-performing and other segments to evaluate revenue concentration. Additional segmentation by age group, gender, and country enabled deeper insight into customer behavior and market distribution.

###### Key Findings

* A limited number of top customers accounted for a significant portion of total profit.
* Profit contribution varied across gender segments, indicating differences in purchasing behavior.
* Certain age groups consistently generated higher profit, reflecting stronger engagement or purchasing power.
* Geographic analysis showed that a few countries dominated overall profit contribution, while other regions presented growth opportunities.

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

### 🛠️ Skills Used

**Excel:**


- Pivot Tables
- Pivot Charts
- Slicers
- Conditional Formatting

**Power Query:**


- Data Cleaning
- Data Transformation
- Data Type Formatting

**Data Modeling:**


- Relationships
- Star Schema Understanding

**Dashboard Design:**


- KPI Cards
- Interactive Filters
- Visual Layout and Formatting

**Sales Analytics Concepts:** 


- Profit Trends
- Product Performance
- Customer Insights
- Time Series Analysis

## 📌 Conclusion

The final dashboards provide clear, interactive summaries of Adventure Works' sales data.
This project helped solidify concepts in **Excel analytics, Power Query,** and **dashboard design**.
Anyone wishing to recreate the project can refer to the YouTube resources below.

## 🎥 Resources 

YouTube tutorial series by the original author:

Part 1: https://youtu.be/VxOOt2dP8Jw?si=okSncDr4spyx2NxO

Part 2: https://youtu.be/sJlVb32jyoQ?si=8ZCmzqgsUT7sDuHk

Part 3: https://youtu.be/LKwyKSw6PhU?si=gW-HOMf8zcBvHi2w

Part 4: https://youtu.be/a1OF_wgRK_U?si=lU2eQ-0mCcuvPGLi
