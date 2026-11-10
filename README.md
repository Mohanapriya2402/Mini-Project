# 🛍️ Retail Sales Performance Analysis | Excel & Power BI

## 📌 Project Overview

The **Retail Sales Performance Analysis** project focuses on analyzing retail sales transactions, customer information, product categories, payment methods, and sales channels using Microsoft Excel and Power BI.

The project involves data cleaning, data transformation, formula calculations, Pivot Table analysis, data modeling, DAX measures, and interactive dashboard development.

The objective is to transform raw retail sales data into meaningful business insights, identify sales trends, evaluate product category performance, understand customer purchasing patterns, and support data-driven business decisions.

This is an end-to-end Data Analytics project that demonstrates practical skills in Excel and Power BI.

## 🎯 Project Objectives

- Clean and preprocess raw retail sales data using Excel.
- Handle inconsistent text formats and improve data quality.
- Calculate revenue using Excel formulas.
- Perform sales analysis using Pivot Tables.
- Analyze sales performance across product categories.
- Compare sales across cities and sales channels.
- Develop DAX measures using Power BI.
- Create an interactive retail sales dashboard.
- Identify key business insights and provide recommendations.

## 📂 Dataset Description

**Dataset Name:** Retail Sales Dataset

**Domain:** Retail / Sales Analytics

**Dataset Size:** 500 records and 16 columns (as documented in the project report).

**Dataset Source:** [Zoho Analytics – Sample Data Generator](https://www.zoho.com/analytics/sample-data-generator.html)

### Dataset Attributes

| Column | Description |
|---|---|
| Order ID | Unique order identification number |
| Order Date | Date of the customer order |
| Customer Name | Name of the customer |
| Gender | Customer gender category |
| City | Customer's city |
| State | Customer's state |
| Product Name | Name of the purchased product |
| Product Category | Category of the product |
| Quantity | Number of units purchased |
| Unit Price | Selling price per unit |
| Discount Percent | Discount applied to the order |
| Sales Amount | Sales amount generated |
| Payment Method | Method used for payment |
| Order Status | Current status of the order |
| Sales Channel | Channel through which the order was placed |
| Revenue | Revenue generated from the transaction |

## 🛠️ Tools & Technologies

- **Microsoft Excel** – Data cleaning, transformation, formulas, and Pivot Tables.
- **Microsoft Power BI** – Data modeling, DAX calculations, and dashboard development.
- **DAX** – Business metrics and KPI calculations.
- **Data Visualization** – Interactive charts, cards, and slicers.

# 🧹 1. Data Cleaning & Preprocessing Using Excel

Performed data cleaning and transformation to improve the quality and consistency of the retail sales dataset.

### 🔹 Text Cleaning and Standardization

Used the following Excel functions:

```excel
=PROPER(TRIM(CLEAN(C2)))
```

**Functions and Purpose:**

- `CLEAN()` – Removes non-printing characters.
- `TRIM()` – Removes unnecessary spaces.
- `PROPER()` – Converts names into proper capitalization.

### 🔹 Revenue Calculation

Calculated revenue using Quantity, Unit Price, and Discount Percentage.

```excel
=ROUND(I2*J2*(1-K2/100),2)
```

**Purpose:**

- Calculates the discounted revenue.
- Applies the discount percentage.
- Uses `ROUND()` to maintain two decimal places.

### 🔹 Data Preprocessing Techniques

- Removed duplicate records where appropriate.
- Handled missing values.
- Standardized text formats.
- Created calculated fields.
- Applied filtering and sorting.
- Prepared the dataset for further analysis.

# 📊 2. Pivot Table Analysis Using Excel

Created Pivot Tables to summarize and analyze retail sales performance across different business dimensions.

### Analyses Performed

**1. Product Category-Wise Sales Analysis**
- Compared sales across product categories.
- Identified high-performing and low-performing categories.

**2. Payment Method-Wise Sales Analysis**
- Analyzed sales distribution across payment methods.
- Compared customer payment preferences.

**3. Order Status-Wise Sales Analysis**
- Evaluated sales based on order status.
- Compared different order processing stages.

**4. Sales Channel-Wise Analysis**
- Compared sales generated through different sales channels.
- Identified the highest and lowest-performing channels.

**5. Month-Wise Sales Analysis**
- Analyzed monthly sales performance.
- Identified variations in sales across different months.

# 🧮 3. Data Modeling & DAX Calculations Using Power BI

Imported the cleaned retail sales dataset into Power BI and created DAX measures to calculate important business metrics.

### DAX Measures

**1. Total Revenue**

```dax
Total Revenue =
SUM(SampleData[revenue])
```

Calculates the overall revenue generated from retail transactions.

**2. Total Orders**

```dax
Total Orders =
DISTINCTCOUNT(SampleData[order_id])
```

Calculates the total number of unique orders.

**3. Total Quantity**

```dax
Total Quantity =
SUM(SampleData[quantity])
```

Calculates the total number of products sold.

**4. Average Order Value**

```dax
Average Order Value =
DIVIDE([Total Revenue], [Total Orders], 0)
```

Calculates the average revenue generated per order.

# 📈 4. Power BI Dashboard & Data Visualization

Developed an interactive **Retail Sales Performance Analysis Dashboard** to provide a clear overview of business performance.

### Key Performance Indicators (KPIs)

| KPI | Description |
|---|---|
| Total Revenue | Overall revenue generated |
| Total Orders | Total number of unique orders |
| Total Quantity Sold | Total units sold |
| Average Order Value | Average revenue per order |

### Visualizations Created

**1. Sales by Product Category**
- Compares revenue across product categories.
- Identifies top-performing categories.

**2. Monthly Sales Trends**
- Displays monthly sales performance.
- Highlights sales fluctuations over time.

**3. Sales by Sales Channel**
- Compares revenue across different sales channels.
- Identifies the most effective sales channel.

**4. Sales by Order Status**
- Analyzes sales based on order status.
- Provides an overview of order performance.

**5. Sales by Payment Method**
- Displays sales distribution across payment methods.
- Helps understand payment preferences.

### Interactive Dashboard Features

Added interactive slicers to enable dynamic filtering:

- Product Category
- Order Date
- Sales Channel

These filters allow users to explore sales performance from different perspectives.

# 🔍 5. Key Business Insights & Findings

Based on the analysis performed in Excel and Power BI, the following insights were identified:

### 🔹 Product Category Performance

- Electronics recorded the highest sales, generating **39,277.50**.
- Travel Gear recorded the lowest sales, contributing **2,642.66**.

### 🔹 Sales Channel Performance

- Website was the highest-performing sales channel, generating **22,282.44**.
- Social Media recorded the lowest sales, contributing **241.40**.

### 🔹 Overall Sales Performance

- Total revenue generated: **64,706.43**
- Unique orders reported: **501**
- Total quantity sold: **934 products**

### 🔹 Monthly Sales Performance

Month-wise analysis helped identify variations in sales performance across different months.

# 📋 6. Business Recommendations

Based on the findings, the following recommendations were identified:

- Maintain the availability of high-performing Electronics products.
- Improve marketing strategies through the Website sales channel.
- Develop promotional strategies to increase Travel Gear sales.
- Explore opportunities to improve Social Media sales.
- Monitor monthly sales trends to support business planning.
- Analyze customer payment preferences to improve the purchasing experience.

# 📊 7. Types of Data Analytics

This project demonstrates different analytical approaches:

| Analytics Type | Application |
|---|---|
| Descriptive Analytics | Summarizing historical sales performance |
| Diagnostic Analytics | Comparing sales across categories, channels, and payment methods |
| Predictive Analytics | Exploring historical monthly sales patterns for potential future planning |
| Prescriptive Analytics | Recommending actions to improve sales performance |

*Note: Predictive analysis in this project is based on historical trends. A separate forecasting model was not developed.*

# 🧠 Skills Demonstrated

Through this project, I developed practical experience in:

- Data Cleaning and Preprocessing
- Excel Formulas and Functions
- Text Standardization
- Revenue Calculations
- Pivot Table Analysis
- Data Modeling
- DAX Measures
- KPI Development
- Data Visualization
- Interactive Dashboard Development
- Sales Performance Analysis
- Business Insights and Recommendations

# 📁 Project Structure

```text
Retail-Sales-Performance-Analysis/
│
├── README.md
│
├── Dataset/
│   └── Retail_Sales_Dataset.xlsx
│
├── Excel/
│   └── Cleaned_Retail_Sales_Dataset.xlsx
│
├── PowerBI/
│   └── RETAIL_SALES_ANALYSIS.pbix
│
└── Report/
    └── Retail_Sales_Analysis_Final.pdf
```

# 🎓 8. Learning Outcomes

This project provided hands-on experience in the complete data analytics workflow, from raw data preprocessing to interactive business reporting.

I gained practical knowledge of:

- Cleaning and preparing real-world-style retail sales data using Excel.
- Applying Excel formulas for data transformation and revenue calculations.
- Using Pivot Tables for business data summarization.
- Creating DAX measures for important business KPIs.
- Building interactive dashboards using Power BI.
- Identifying sales patterns and translating data into business insights.

## 👩‍💻 About Me

I am an **Aspiring Data Analyst** passionate about exploring data, identifying patterns, and transforming raw information into meaningful business insights.

I am continuously improving my technical and analytical skills through practical projects and hands-on learning.

**My Data Analytics Learning Journey:**

Microsoft Excel | SQL | Python | Power BI | DAX | Data Analytics

This project is part of my growing Data Analyst portfolio, demonstrating my practical learning and progress in the field of Data Analytics.

---

⭐ If you find this project interesting, feel free to explore my repository!
