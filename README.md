# 20 Beginner Excel Data Analyst Practice Questions

## Project Overview

A beginner-friendly Excel practice project focused on answering common sales-analysis questions using formulas, filters, sorting, date functions, PivotTables, PivotCharts, and a basic dashboard.

## Dataset

The `sales-dataset.xlsx` workbook contains one sheet, `Raw_Sales_Data`, with 50 sales records and 10 raw columns:

- `Order_ID`
- `Order_Date`
- `Customer`
- `City`
- `Category`
- `Product`
- `Quantity`
- `Unit_Price`
- `Salesperson`
- `Payment_Method`

The dataset covers January–August 2026 and includes orders from Karachi, Lahore, and Islamabad; Electronics, Furniture, and Clothing products; four salespeople; and Card, Cash, and Bank Transfer payments.

The dataset intentionally does not include calculated fields. Create `Revenue` using `Quantity × Unit_Price` as part of the first exercise. Add any other helper columns required by the questions during practice.

## Practice Questions

### Level 1 — Basic Excel Analysis

1. **Calculate Total Revenue:** Calculate revenue for each order using `Quantity × Unit_Price`.
2. **Calculate Total Sales:** Sum the revenue for all orders using `SUM()`.
3. **Calculate Total Quantity Sold:** Sum all units sold using `SUM()`.
4. **Calculate Average Order Revenue:** Calculate average revenue per order using `AVERAGE()`.
5. **Find the Highest Order Value:** Find the maximum order revenue using `MAX()`.

### Level 2 — Basic Data Analysis

6. **Find the Lowest Order Value:** Find the minimum order revenue using `MIN()`.
7. **Count Total Orders:** Count the records using `COUNT()` and/or `COUNTA()`.
8. **Count Orders by City:** Count orders from Karachi, Lahore, and Islamabad using `COUNTIF()`.
9. **Calculate Revenue by City:** Sum revenue for each city using `SUMIF()`.
10. **Calculate Revenue by Category:** Sum revenue for Electronics, Furniture, and Clothing using `SUMIF()`.

### Level 3 — Filtering and Conditional Analysis

11. **Find High-Value Orders:** Filter orders with revenue greater than 50,000.
12. **Count High-Value Orders:** Count orders with revenue greater than 50,000 using `COUNTIF()`.
13. **Find Electronics Orders:** Filter the dataset to show only Electronics orders.
14. **Calculate Electronics Revenue:** Sum revenue for Electronics using `SUMIF()`.
15. **Identify Large Orders:** Add an `Order_Size` column: `Large` for revenue ≥ 100,000, `Medium` for revenue ≥ 50,000 but below 100,000, and `Small` for revenue below 50,000. Use `IF()` or nested `IF()`.

### Level 4 — Practical Data Analyst Questions

16. **Find the Best Salesperson:** Calculate revenue by salesperson using `SUMIF()`, then sort or compare totals to identify the highest.
17. **Analyze Payment Methods:** Count Card, Cash, and Bank Transfer orders using `COUNTIF()` and identify the most frequently used method.
18. **Calculate Average Revenue by Category:** Calculate average order revenue for each category using `AVERAGEIF()`.
19. **Analyze Monthly Sales:** Extract the month from `Order_Date` and calculate revenue by month using date functions with `SUMIF()` or a PivotTable.
20. **Create a Basic Sales Dashboard:** Build KPI summaries for Total Revenue, Total Orders, Total Units Sold, and Average Order Value. Add charts for revenue by category, city, month, and salesperson using PivotTables and PivotCharts.

## Skills Practiced

| Excel Skill | Questions |
|---|---|
| Basic formulas and calculated columns | 1–7 |
| `SUM()` | 2–3 |
| `AVERAGE()` | 4 |
| `MAX()` / `MIN()` | 5–6 |
| `COUNT()` / `COUNTA()` | 7 |
| `COUNTIF()` | 8, 12, 17 |
| `SUMIF()` | 9, 10, 14, 16, 19 |
| `AVERAGEIF()` | 18 |
| `IF()` / nested `IF()` | 15 |
| Filtering | 11, 13 |
| Sorting | 16 |
| Date analysis | 19 |
| PivotTables and PivotCharts | 20 |
| Dashboard design | 20 |

## Recommended Progression

Work through the exercises in order:

1. Questions 1–7: Basic formulas
2. Questions 8–10: Conditional aggregation
3. Questions 11–15: Filtering and logical conditions
4. Questions 16–19: Business-focused analysis
5. Question 20: Build a mini sales dashboard

The goal is to translate business questions into calculations and insights, rather than simply memorizing Excel functions.
