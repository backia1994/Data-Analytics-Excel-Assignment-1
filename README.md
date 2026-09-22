📊 Excel Assignment 1 – Data Exploration
📌 Project Overview

This project is part of my **Data Analytics learning journey**, where I am building foundational skills in Microsoft Excel for data exploration and analysis.

The objective of this assignment is to analyze a **Product Dataset** using Excel formulas and functions. The analysis focuses on summarizing numerical data, applying logical conditions, performing conditional calculations, and extracting meaningful information from Product IDs.

This project demonstrates my ability to use Excel as a basic data analysis and data transformation tool.

---

🎯 Problem Statement

As a Data Analyst, the task is to perform basic exploratory analysis on a product dataset containing information such as:

* Product ID
* Product Name
* Brand Name
* Price
* Quantity
* Category

The analysis uses Excel functions to generate summary statistics, categorize products based on price, perform category-based calculations, and extract information from Product IDs.

---
📂 Dataset Description

**Dataset:** Product Dataset

| Column       | Description                                                      |
| ------------ | ---------------------------------------------------------------- |
| Product ID   | Unique identifier containing day, month, and country information |
| Product Name | Name of the product                                              |
| Brand Name   | Product brand                                                    |
| Price ($)    | Price of the product                                             |
| Quantity     | Quantity of products                                             |
| Category     | Product category                                                 |

Additional columns were created during the analysis:

* Price Range
* Day
* Country Code
* Month

---

🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Formulas & Functions
* Data Exploration
* Data Summarization
* Logical Analysis
* Conditional Aggregation
* Text Manipulation

---

🔎 Analysis Performed

1. Basic Data Exploration

I used Excel functions to calculate important summary statistics from the dataset.

*Total Price

Calculated the total price of all products using:

```excel
=SUM(D2:D35)
```

### Number of Products

Calculated the total number of products using:

```excel
=COUNT(D2:D35)
```

### Average Price

Calculated the average product price using:

```excel
=AVERAGE(D2:D35)
```

These calculations provide a basic understanding of the numerical characteristics of the dataset.

---

## 2. Minimum and Maximum Price

To understand the price range of the products, I calculated the minimum and maximum prices.

### Minimum Price

```excel
=MIN(D2:D35)
```

### Maximum Price

```excel
=MAX(D2:D35)
```

This helps identify the lowest-priced and highest-priced products in the dataset.

---

## 3. Logical Function – IF

A new column named **Price Range** was created to categorize products based on their price.

### Business Rule

* Price ≥ $500 → **HIGH PRICE**
* Price < $500 → **STANDARD PRICE**

### Formula Used

```excel
=IF(D2>=500,"HIGH PRICE","STANDARD PRICE")
```

The formula was applied to all product records.

This demonstrates the use of **logical conditions in Excel** to transform numerical data into meaningful categories.

---

## 4. Conditional Aggregation – SUMIF

The `SUMIF` function was used to calculate the total price of products belonging to the **Electronics** category.

### Formula

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

Where:

* `F2:F35` → Category range
* `"Electronics"` → Criteria
* `D2:D35` → Price range

This demonstrates how conditional aggregation can be used to analyze specific segments of a dataset.

---

## 5. Conditional Counting – COUNTIF

The `COUNTIF` function was used to determine the number of products with a price below $100.

### Formula

```excel
=COUNTIF(D2:D35,"<100")
```

This helps identify how many products fall below a particular price threshold.

---

## 6. Text Manipulation – LEFT, RIGHT & MID

The Product ID contains structured information in the following format:

```text
28-JAN-US
```

Excel text functions were used to extract specific parts of the Product ID.

### Day

The first two characters were extracted using the `LEFT` function.

```excel
=LEFT(A2,2)
```

Example:

```text
28-JAN-US → 28
```

---

### Country Code

The last two characters were extracted using the `RIGHT` function.

```excel
=RIGHT(A2,2)
```

Example:

```text
28-JAN-US → US
```

---

### Month

The month portion was extracted using the `MID` function.

```excel
=MID(A2,4,6)
```

Example:

```text
28-JAN-US → JAN
```

These transformations demonstrate how text functions can be used to extract structured information from identifiers.

---

# 📊 Excel Functions Used

| Function  | Purpose                                  |
| --------- | ---------------------------------------- |
| `SUM`     | Calculate total price                    |
| `COUNT`   | Count products                           |
| `AVERAGE` | Calculate average price                  |
| `MIN`     | Find minimum price                       |
| `MAX`     | Find maximum price                       |
| `IF`      | Categorize products based on price       |
| `SUMIF`   | Calculate price for Electronics category |
| `COUNTIF` | Count products priced below $100         |
| `LEFT`    | Extract Day from Product ID              |
| `RIGHT`   | Extract Country Code                     |
| `MID`     | Extract Month                            |

---

# 📈 Skills Demonstrated

Through this assignment, I practiced the following Data Analytics skills:

### Data Exploration

* Summarizing numerical data
* Understanding basic dataset characteristics
* Identifying minimum and maximum values

### Logical Analysis

* Applying conditional logic
* Creating business-based categories

### Conditional Aggregation

* Filtering calculations based on criteria
* Using `SUMIF` and `COUNTIF`

### Text Manipulation

* Extracting structured information from text
* Working with `LEFT`, `RIGHT`, and `MID`

### Spreadsheet Proficiency

* Creating calculated columns
* Applying formulas across datasets
* Organizing analytical results
* Using Excel as a data analysis tool

---

# 📁 Project Structure

```text
Excel-Assignment-1-Data-Exploration/
│
├── README.md
│
├── Excel/
│   └── Excel_Assignment_1_Data_Exploration.xlsx
│
├── PDF/
│   └── Excel_Assignment_1_Data_Exploration.pdf
│
└── Screenshots/
    ├── Basic_Calculations.png
    ├── Min_Max_Analysis.png
    ├── Price_Range_IF.png
    ├── SUMIF_COUNTIF.png
    └── Text_Functions.png
```

---

# 📸 Screenshots & Documentation

The project documentation includes screenshots demonstrating:

* SUM, COUNT and AVERAGE calculations
* MIN and MAX calculations
* Price Range classification using IF
* SUMIF and COUNTIF calculations
* LEFT, RIGHT and MID text extraction
* Excel formulas displayed in the formula bar

---

# 📝 Key Learning Outcomes

Through this assignment, I gained practical experience in:

* Performing basic data exploration using Excel
* Applying Excel formulas to real-world-style datasets
* Using logical functions for data categorization
* Performing conditional calculations
* Extracting information from structured text
* Creating calculated columns
* Presenting analytical work in a structured format

This assignment forms part of my **Data Analytics portfolio** and represents my initial hands-on practice in spreadsheet-based data analysis.

---

# 🚀 Future Learning

As I continue developing my Data Analytics skills, I plan to expand this project portfolio by working with:

* Advanced Excel
* SQL
* Power BI
* Python
* Pandas
* Data Cleaning
* Exploratory Data Analysis (EDA)
* Data Visualization
* Statistical Analysis
* Real-world Data Analytics Projects

---

## 👩‍💻 About Me

I am an **aspiring Data Analyst** currently developing my technical and analytical skills through hands-on projects.

My goal is to build a strong portfolio demonstrating practical knowledge of:

**Excel → SQL → Power BI → Python → Data Analysis**

I am continuously learning and applying these skills through projects and assignments.
