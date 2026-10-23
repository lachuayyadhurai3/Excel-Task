# Excel-Task
This repository contains my Microsoft Excel tasks and Practical projects focused on data analysis and visualization. It showcases my ability to clean, organize, analyze, and interpret data using Excel.


# 📊 Data Exploration Using Excel

## 📌 Project Overview

This project focuses on exploring and analyzing a product dataset using Microsoft Excel. It demonstrates how Excel functions can be used to summarize product prices, identify minimum and maximum values, classify products based on price, calculate conditional totals, and extract information from Product IDs.

The project uses a product dataset containing **34 records** with details such as Product ID, Product Name, Brand Name, Price, Quantity, and Category.

## 🎯 Project Objectives

- Calculate total, count, and average product prices.
- Identify minimum and maximum product prices.
- Classify products using the IF function.
- Calculate category-wise totals using SUMIF.
- Count products meeting specific price conditions using COUNTIF.
- Extract Day, Country Code, and Month from Product IDs.
- Develop practical Excel skills for data exploration and analysis.

## 🛠️ Tools Used

- Microsoft Excel
- Excel Formulas and Functions
- Product Dataset

## 📂 Dataset Description

The dataset contains the following columns:

| Column | Description |
|---|---|
| Product ID | Unique or identifying code associated with a product record |
| Product Name | Name of the product |
| Brand Name | Brand or manufacturer |
| Price ($) | Price of the product in dollars |
| Quantity | Available or recorded product quantity |
| Category | Product category |

**Dataset size:** 34 records and 6 original columns.

**Note:** The dataset contains repeated records. All calculations in this project include every row provided, including duplicates.

---

## 1. ➕ SUM, COUNT, and AVERAGE

Perform basic calculations using Excel functions.

### A. Calculate the Total Price of All Products

Use the `SUM` function to calculate the total of all product prices.

**Excel Formula:**

```excel
=SUM(D2:D35)
```

**Expected Result:** `$10,100`

**Explanation:** Adds all numeric values in the Price column.

### B. Determine the Number of Products

Use `COUNTA` to count the product records using Product ID.

**Excel Formula:**

```excel
=COUNTA(A2:A35)
```

**Expected Result:** `34 records`

**Explanation:** Counts all non-empty Product ID cells.

Alternatively, use the `COUNT` function on the numeric Price column:

```excel
=COUNT(D2:D35)
```

Expected result: `34`

### C. Calculate the Average Price

Use the `AVERAGE` function to calculate the average price of the products.

**Excel Formula:**

```excel
=AVERAGE(D2:D35)
```

**Expected Result:** `$297.06`

**Explanation:** Calculates the sum of the prices divided by the number of numeric price entries.

---

## 2. 📉 Minimum and Maximum Values

Use Excel functions to identify the lowest and highest product prices.

### A. Identify the Minimum Price

**Excel Formula:**

```excel
=MIN(D2:D35)
```

**Expected Result:** `$30`

**Product:** T-shirt — Adidas

### B. Identify the Maximum Price

**Excel Formula:**

```excel
=MAX(D2:D35)
```

**Expected Result:** `$1,000`

**Product:** Laptop — Dell

**Purpose:** These functions help identify the least expensive and most expensive products in the dataset.

---

## 3. 🧠 Logical Function – IF

Create a new column named **Price Range** to classify products according to their price.

### Business Rules

- If Price is greater than or equal to $500, classify it as `High Price`.
- If Price is less than $500, classify it as `Standard Price`.

### Excel Formula

Add `Price Range` as a new header in cell G1.

Enter the following formula in cell G2:

```excel
=IF(D2>=500,"High Price","Standard Price")
```

Drag the formula down to cell G35.

### Example Output

| Product Name | Price ($) | Price Range |
|---|---:|---|
| Laptop | 1000 | High Price |
| Sneakers | 80 | Standard Price |
| Coffee Maker | 130 | Standard Price |
| Smartphone | 900 | High Price |
| Tablet | 500 | High Price |
| T-shirt | 30 | Standard Price |

**Purpose:** The IF function helps categorize products based on business rules and makes the dataset easier to analyze.

---

## 4. 🧮 Conditional Functions – SUMIF and COUNTIF

### A. SUMIF – Total Price of Electronics Products

Calculate the total price of all products belonging to the Electronics category.

**Excel Formula:**

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

**Expected Result:** `$8,050`

**Explanation:**

- `F2:F35` — Category range.
- `"Electronics"` — Condition to match.
- `D2:D35` — Price values to sum.

**Purpose:** This function helps analyze the total price of products within a specific category.

### B. COUNTIF – Count Products Priced Below $100

Determine how many product records have a price below $100.

**Excel Formula:**

```excel
=COUNTIF(D2:D35,"<100")
```

**Expected Result:** `11 records`

**Explanation:** Counts every record whose Price value is less than $100.

**Purpose:** This function helps identify and count lower-priced products.

---

## 5. 🔤 Text Formatting – LEFT, RIGHT, and MID

Use Excel text functions to extract information from the Product ID column.

The Product ID format in this dataset is:

`DD-MMM-CC`

For example: `28-JAN-US`

- `DD` represents the day code.
- `MMM` represents the month code.
- `CC` represents the country code.

### A. LEFT – Create the Day Column

Create a new column named **Day** in cell H1.

Extract the first two characters from each Product ID.

**Excel Formula:**

```excel
=LEFT(A2,2)
```

**Example:**

| Product ID | Day |
|---|---|
| 28-JAN-US | 28 |
| 15-FEB-US | 15 |
| 03-MAR-US | 03 |
| 11-APR-US | 11 |

**Explanation:** The LEFT function extracts the first two characters from the Product ID.

### B. RIGHT – Create the Country Code Column

Create a new column named **Country Code** in cell I1.

Extract the last two characters from each Product ID.

**Excel Formula:**

```excel
=RIGHT(A2,2)
```

**Example:**

| Product ID | Country Code |
|---|---|
| 28-JAN-US | US |
| 15-FEB-US | US |
| 07-JUN-UK | UK |
| 17-JUN-IN | IN |
| 08-DEC-DE | DE |

**Explanation:** The RIGHT function extracts the final two characters, representing the country code in this dataset.

### C. MID – Create the Month Column

Create a new column named **Month** in cell J1.

Extract the 4th to 6th characters from each Product ID.

**Excel Formula:**

```excel
=MID(A2,4,3)
```

**Example:**

| Product ID | Month |
|---|---|
| 28-JAN-US | JAN |
| 15-FEB-US | FEB |
| 03-MAR-US | MAR |
| 11-APR-US | APR |
| 22-MAY-US | MAY |
| 07-JUN-UK | JUN |

**Explanation:** The MID function starts at character position 4 and extracts 3 characters, returning the month abbreviation.

### Final Dataset Structure

After adding the new columns, the dataset will contain the following fields:

| Column | Field Name |
|---|---|
| A | Product ID |
| B | Product Name |
| C | Brand Name |
| D | Price ($) |
| E | Quantity |
| F | Category |
| G | Price Range |
| H | Day |
| I | Country Code |
| J | Month |

Copy the formulas down through row 35 to apply them to all 34 records.

---

## 6. 📋 Summary of Excel Functions

| Function | Purpose | Formula |
|---|---|---|
| SUM | Total product prices | `=SUM(D2:D35)` |
| COUNT | Count numeric price values | `=COUNT(D2:D35)` |
| COUNTA | Count non-empty Product IDs | `=COUNTA(A2:A35)` |
| AVERAGE | Average product price | `=AVERAGE(D2:D35)` |
| MIN | Minimum product price | `=MIN(D2:D35)` |
| MAX | Maximum product price | `=MAX(D2:D35)` |
| IF | Classify products by price | `=IF(D2>=500,"High Price","Standard Price")` |
| SUMIF | Total price of Electronics products | `=SUMIF(F2:F35,"Electronics",D2:D35)` |
| COUNTIF | Count products priced below $100 | `=COUNTIF(D2:D35,"<100")` |
| LEFT | Extract the first two characters | `=LEFT(A2,2)` |
| RIGHT | Extract the last two characters | `=RIGHT(A2,2)` |
| MID | Extract the month abbreviation | `=MID(A2,4,3)` |

---

## 7. 📊 Key Findings

Based on the supplied dataset, the following results were obtained:

| Metric | Result |
|---|---:|
| Total price | $10,100 |
| Number of product records | 34 |
| Average price | $297.06 |
| Minimum price | $30 |
| Maximum price | $1,000 |
| Electronics total price | $8,050 |
| Products priced below $100 | 11 |

**Note:** Total price and average price are calculated from the Price column only, not by multiplying Price by Quantity. Duplicate records are included in all calculations.

---

## 8. 🎓 Key Learnings

Through this project, I practiced:

- Basic statistical analysis using SUM, COUNT, AVERAGE, MIN, and MAX.
- Logical decision-making using IF.
- Conditional calculations using SUMIF and COUNTIF.
- Text extraction using LEFT, RIGHT, and MID.
- Adding calculated columns to an existing dataset.
- Applying Excel formulas to a practical product dataset.
- Summarizing data to identify useful insights.

## 9. 🚀 Future Improvements

- Analyze total inventory value using Price × Quantity.
- Compare product prices across categories and brands.
- Explore sorting, filtering, and conditional formatting.
- Create PivotTables and PivotCharts.
- Build an interactive Excel dashboard.
- Extend the analysis using SQL, Python, and Power BI.

## 👨‍💻 Author

**A Lakshmana Narayanan**
**Aspiring Data Analyst**

This project is part of my Data Analytics learning journey, focusing on Excel functions, data exploration, and practical analysis.

## ⭐ Acknowledgment

Created for educational purposes to develop practical Microsoft Excel skills and establish a foundation for more advanced data analytics projects.
