# 🛒 E-Commerce Sales Analysis & Interactive Dashboard using Excel

## 📌 Project Overview

This project is part of my journey toward becoming a **Data Analyst**.

The objective of this project is to perform **data cleaning, transformation, analysis, summarization, statistical analysis, and visualization using Microsoft Excel**.

The project uses an **E-commerce Sales Dataset** containing customer, product, store, and sales information. Multiple tables were combined and analyzed to identify useful business insights related to sales, products, customers, categories, brands, quantities, and revenue.

The assignment focuses on practical Excel skills including data cleaning, formulas, lookup functions, PivotTables, descriptive statistics, charts, PivotCharts, and dashboard development.

The assignment requires data import and organization, cleaning, Excel formulas, XLOOKUP/VLOOKUP, PivotTables, descriptive statistics, at least three chart types, an interactive dashboard, and PivotCharts.

---

# 🎯 Project Objectives

The main objectives of this project are to:

* Import and organize E-commerce data in Excel.
* Create structured Excel tables.
* Identify and remove duplicate records.
* Handle missing and inconsistent data.
* Correct formatting and data-quality issues.
* Remove irrelevant data where applicable.
* Combine multiple tables using `XLOOKUP` / `VLOOKUP`.
* Apply Excel formulas for calculations and logical analysis.
* Perform descriptive statistical analysis.
* Create PivotTables for data summarization.
* Create multiple charts and PivotCharts.
* Build an interactive E-commerce sales dashboard.
* Identify and communicate important business insights.

---

# 📂 Dataset

**Dataset:** E-commerce Store Dataset

The dataset contains information related to:

### Customer Information

* Customer ID
* Customer Name
* Age
* Gender
* City
* State
* Country
* Loyalty Level

### Product Information

* Product ID
* Product Name
* Category
* Brand
* Cost
* Stock

### Store Information

* Store ID
* Store Name
* Region
* City
* Store Type

### Sales Information

* Sales ID
* Order Date
* Customer ID
* Product ID
* Store ID
* Quantity
* Unit Price
* Discount
* Payment Type
* Total Amount

---

# 🗂️ Workbook Structure

The Excel workbook contains multiple sheets used for different stages of the analysis.

| Sheet                        | Purpose                                                 |
| ---------------------------- | ------------------------------------------------------- |
| `Customer_Dim`               | Customer master data                                    |
| `Product_Dim`                | Product master data                                     |
| `Store_Dim`                  | Store master data                                       |
| `Sales_Fact`                 | Main sales transaction data                             |
| `forecast`                   | Sales forecasting-related data                          |
| `month vs sales`             | Monthly sales analysis                                  |
| `cate vs quantity vs unit`   | Category, quantity and unit-price analysis              |
| `cate vs total sales amount` | Category-level sales analysis                           |
| `bramd vs sales`             | Brand-level sales analysis                              |
| `combine_tables`             | Combined customer, product, store and sales information |
| `dashboarc`                  | Interactive E-commerce sales dashboard                  |

---

# 🧹 1. Data Entry & Organization

The E-commerce dataset was imported into Microsoft Excel and organized into structured tables.

The workbook follows a dimensional structure with:

```text
Customer Dimension
        ↓
Product Dimension
        ↓
Store Dimension
        ↓
Sales Fact
        ↓
Combined Analysis Dataset
```

This structure makes it easier to connect different datasets and perform analysis.

---

# 🧹 2. Data Cleaning & Transformation

Data cleaning was performed before analysis to improve data quality.

## Duplicate Records

Duplicate records were identified and removed where applicable.

Excel's **Remove Duplicates** feature was used to prevent duplicate transactions from affecting the analysis.

---

## Missing Values

The dataset was checked for missing or blank values.

Where appropriate, missing values were:

* Investigated.
* Corrected using available information.
* Standardized where possible.
* Retained when a reliable replacement could not be determined.

---

## Inconsistent Data

The dataset was checked for:

* Inconsistent text values.
* Extra spaces.
* Formatting differences.
* Incorrect or inconsistent category/brand values.
* Incorrect data types.

Excel functions such as `TRIM`, `CLEAN`, and `PROPER` were used where applicable to improve consistency.

Example:

```excel
=TRIM(A2)
```

```excel
=CLEAN(A2)
```

```excel
=PROPER(A2)
```

---

# 🔗 3. Combining Tables Using XLOOKUP / VLOOKUP

The E-commerce dataset contains separate tables for customers, products, stores, and sales.

Lookup functions were used to combine information from these tables.

## XLOOKUP

Example:

```excel
=XLOOKUP([@[Customer_ID]],Customer_Dim[Customer_ID],Customer_Dim[Name])
```

This retrieves the customer name based on the Customer ID.

Similarly, product and store information can be retrieved using their respective IDs.

---

## VLOOKUP

Where applicable, `VLOOKUP` can also be used to retrieve information from related tables.

Example:

```excel
=VLOOKUP(A2,Customer_Dim!A:H,2,FALSE)
```

Lookup functions helped create the **combined analysis dataset** used for further analysis.

---

# 🧮 4. Excel Formulas & Functions

Several Excel formulas were used for calculations, logical analysis, and data manipulation.

### SUM

Used to calculate total sales/revenue.

```excel
=SUM(Sales_Fact[Total_Amount])
```

### AVERAGE

Used to calculate average values such as unit price.

```excel
=AVERAGE(Sales_Fact[Unit_Price])
```

### COUNTIF

Used to count records based on conditions.

```excel
=COUNTIF(Sales_Fact[Payment_Type],"Credit Card")
```

### IF

Used to create logical categories.

Example:

```excel
=IF([@[Total_Amount]]>=500,"High Value","Regular")
```

### TEXT

Used to extract or format date information.

Example:

```excel
=TEXT([@[Order_Date]],"mmm")
```

This can be used to create a Month column for monthly analysis.

---

# 📊 5. Descriptive Statistics

Descriptive statistics were calculated to understand the distribution and characteristics of the sales data.

The analysis included:

* Mean
* Median
* Mode
* Standard Deviation
* Minimum
* Maximum
* Count

### Mean

```excel
=AVERAGE(range)
```

### Median

```excel
=MEDIAN(range)
```

### Mode

```excel
=MODE.SNGL(range)
```

### Standard Deviation

```excel
=STDEV.S(range)
```

### Minimum

```excel
=MIN(range)
```

### Maximum

```excel
=MAX(range)
```

These statistics provide a basic understanding of the sales and pricing distribution.

The assignment specifically evaluates mean, standard deviation, mode and Analysis ToolPak/similar statistical analysis.

---

# 📌 6. PivotTable Analysis

PivotTables were created to summarize the E-commerce dataset from multiple dimensions.

Examples of analysis include:

### Monthly Sales

```text
Month → Total Sales
```

### Category Analysis

```text
Category → Quantity → Unit Price
```

### Category Revenue

```text
Category → Total Sales Amount
```

### Brand Analysis

```text
Brand → Number of Sales → Total Amount
```

PivotTables make it easier to compare different business dimensions and identify patterns.

The assignment requires PivotTables to summarize multi-dimensional data.

---

# 📈 7. Data Visualization

Different chart types were used to present the analysis visually.

## Column Chart

Used to compare sales across categories or time periods.

## Bar Chart

Used to compare brands or categories based on sales.

## Pie Chart

Used to show the contribution of different categories or payment methods.

## Sparklines

Used to provide compact visual trends within tables.

## PivotCharts

PivotCharts were created from PivotTables to provide interactive visual analysis.

The assignment requires at least three chart types, clear labels, an organized dashboard, and basic PivotCharts.

---

# 📊 8. Interactive Dashboard

An interactive **E-commerce Sales Dashboard** was created to provide a high-level overview of business performance.

### Dashboard KPIs

The dashboard includes key metrics such as:

* **Average Price**
* **Total Units Sold**
* **Total Revenue**

The dashboard also presents visual analysis for sales performance across different business dimensions.

### Dashboard Components

```text
┌──────────────────────────────────────────────┐
│       E-COMMERCE STORE SALES ANALYSIS        │
├──────────────┬──────────────┬────────────────┤
│ Average      │ Total Units  │ Total Revenue  │
│ Price        │ Sold         │                │
├──────────────┴──────────────┴────────────────┤
│                                              │
│          Sales Visualization                 │
│                                              │
├──────────────────────┬───────────────────────┤
│ Category Analysis    │ Brand / Sales Trend   │
│                      │                       │
├──────────────────────┴───────────────────────┤
│              Interactive Filters             │
└──────────────────────────────────────────────┘
```

The dashboard was designed to make important information easier to understand and communicate.

---

# 🔍 9. Key Analysis Areas

The project focuses on several important business questions:

### Sales Performance

* What is the total revenue?
* How many units were sold?
* How does sales performance change by month?

### Category Performance

* Which category generates the highest sales?
* Which category has the highest quantity sold?
* How does unit price vary by category?

### Brand Performance

* Which brands generate the highest sales?
* Which brands have the highest number of transactions?

### Customer Analysis

* How are customers distributed across loyalty levels?
* What customer characteristics are associated with sales?

### Store Analysis

* Which regions generate higher sales?
* How do different store types perform?

---

# 💡 Key Insights

The analysis was used to identify meaningful patterns from the E-commerce sales data.

Examples of insights investigated include:

* Overall revenue and sales volume.
* Monthly sales trends.
* Top-performing product categories.
* Brand-level sales performance.
* Quantity sold by category.
* Unit price differences across categories.
* Customer and store-level patterns.

The final insights are presented through the PivotTables, charts, and dashboard.

The assignment requires a clear summary connecting findings to the visualizations.

---

# 📋 Skills Demonstrated

| Skill                  | Techniques Used                                |
| ---------------------- | ---------------------------------------------- |
| Spreadsheet Management | Excel tables, sheets and data organization     |
| Data Import            | Dataset import and organization                |
| Data Cleaning          | Duplicates, missing values, consistency checks |
| Data Transformation    | Combining and restructuring tables             |
| Excel Formulas         | SUM, IF, COUNTIF, AVERAGE, TEXT                |
| Lookup Functions       | XLOOKUP / VLOOKUP                              |
| Statistics             | Mean, Median, Mode, Standard Deviation         |
| Data Summarization     | PivotTables                                    |
| Visualization          | Bar, Column, Pie Charts, Sparklines            |
| Advanced Visualization | PivotCharts                                    |
| Dashboard              | Interactive E-commerce dashboard               |
| Communication          | Findings and business insights                 |

---

# 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Tables
* Excel Formulas & Functions
* XLOOKUP / VLOOKUP
* PivotTables
* PivotCharts
* Conditional Formatting
* Descriptive Statistics
* Excel Dashboard

---

# 📂 Project Structure

```text
excel-ecommerce-sales-analysis/
│
├── README.md
│
├── Ecommerce_Sales_Dataset_updated.xlsx
│
└── screenshots/
    ├── data-cleaning.png
    ├── pivot-table.png
    ├── charts.png
    └── dashboard.png
```

---

# 📈 Data Analysis Workflow

```text
Raw E-commerce Data
        ↓
Data Import
        ↓
Table Creation
        ↓
Data Cleaning
        ↓
Missing Value Handling
        ↓
Duplicate Removal
        ↓
Data Validation
        ↓
XLOOKUP / VLOOKUP
        ↓
Combined Dataset
        ↓
Excel Formulas
        ↓
Descriptive Statistics
        ↓
PivotTables
        ↓
Charts & PivotCharts
        ↓
Interactive Dashboard
        ↓
Business Insights
```

---

# 💡 Key Learnings

Through this project, I developed practical skills in:

* Importing and organizing datasets in Excel.
* Cleaning real-world-style data.
* Handling missing and inconsistent values.
* Removing duplicate records.
* Combining multiple tables using lookup functions.
* Applying Excel formulas for analysis.
* Calculating descriptive statistics.
* Creating PivotTables and PivotCharts.
* Building charts and interactive dashboards.
* Communicating data-driven findings.

---

# 🚀 Future Improvements

As I continue developing my Data Analyst skills, I plan to extend this project by:

* Recreating the analysis using **SQL**.
* Performing the same analysis using **Python and Pandas**.
* Building a more advanced dashboard using **Power BI**.
* Performing customer segmentation.
* Performing product-level profitability analysis.
* Creating advanced sales forecasting models.
* Developing automated reporting workflows.

---

# 👨‍💻 About Me

I am an aspiring **Data Analyst** building my portfolio through practical projects.

This project demonstrates my ability to use **Microsoft Excel for data cleaning, analysis, visualization, and reporting**.

I am continuously developing my skills in:

* 📊 Microsoft Excel
* 🗄️ SQL
* 🐍 Python
* 📈 Power BI
* 📉 Statistics
* 🧹 Data Cleaning
* 📊 Data Visualization

This project is one step in my journey toward becoming a professional Data Analyst.

---

⭐ **If you find this project useful, feel free to explore the repository and follow my Data Analyst learning journey.**
