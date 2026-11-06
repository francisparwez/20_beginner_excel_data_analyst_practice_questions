# 20 Beginner Excel Data Analyst Practice Questions

## Project Overview

A beginner-friendly Excel practice project focused on answering common sales-analysis questions using formulas, filters, sorting, date functions, PivotTables, PivotCharts, and a basic dashboard.

## Dataset

The `sales-dataset.xlsx` workbook contains two sheets: `Raw_Sales_Data`, with 50 sales records and the original 10 source columns, and `Analysis`, which contains the calculated revenue, summary metrics, revenue breakdowns, and filtered order results. The original 10 source columns are:

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

The `Analysis` sheet includes a calculated `Revenue` column and summary metrics and revenue breakdowns by city and category. The dataset covers January–August 2026 and includes orders from Karachi, Lahore, and Islamabad; Electronics, Furniture, and Clothing products; four salespeople; and Card, Cash, and Bank Transfer payments.

`Revenue` is calculated for each order using `=Raw_Sales_Data!H2*Raw_Sales_Data!G2` and filled down for all 50 orders. The `Analysis` sheet calculates total revenue with `=SUM(Analysis!A2:A51)`, total quantity sold with `=SUM(Raw_Sales_Data!G2:G51)`, average order revenue with `=AVERAGE(Analysis!A:A)`, the highest order value with `=MAX(Analysis!A:A)`, and the lowest order value with `=MIN(A2:A51)`. It also counts orders with `=COUNT(A2:A51)` and `=COUNTA(A2:A51)`, and counts orders by city with `=COUNTIF(Raw_Sales_Data!D:D,Analysis!C23)` (adjusting the criteria reference for each city). Revenue by city is calculated with `=SUMIF(Raw_Sales_Data!D:D,C28,Analysis!A:A)` and the criteria reference is adjusted for each city. Revenue by category is calculated with `=SUMIF(Raw_Sales_Data!E$2:E$51,B33,Analysis!A$2:A$51)` and filled down for each category. The verified results are **5,189,800 total revenue**, **300 units sold**, **103,796 average order revenue**, **324,000 highest order value**, **25,000 lowest order value**, **50 total orders**, city counts of **20 Karachi**, **17 Lahore**, and **13 Islamabad**, city revenue of **2,174,500 Karachi**, **1,906,000 Lahore**, and **1,109,300 Islamabad**, and category revenue of **2,460,900 Electronics**, **2,039,000 Furniture**, and **689,900 Clothing**.

## Practice Questions

### Level 1 — Basic Excel Analysis

1. **Calculate Total Revenue:** Calculate revenue for each order using `Quantity × Unit_Price`. **Completed** — the `Revenue` column uses `=Raw_Sales_Data!H2*Raw_Sales_Data!G2` and is filled down for all 50 orders.
2. **Calculate Total Sales:** Sum the revenue for all orders using `SUM()`. **Completed** — `=SUM(Analysis!A2:A51)` returns 5,189,800.
3. **Calculate Total Quantity Sold:** Sum all units sold using `SUM()`. **Completed** — `=SUM(Raw_Sales_Data!G2:G51)` returns 300.
4. **Calculate Average Order Revenue:** Calculate average revenue per order using `AVERAGE()`. **Completed** — `=AVERAGE(Analysis!A:A)` returns 103,796.
5. **Find the Highest Order Value:** Find the maximum order revenue using `MAX()`. **Completed** — `=MAX(Analysis!A:A)` returns 324,000.

### Level 2 — Basic Data Analysis

6. **Find the Lowest Order Value:** Find the minimum order revenue using `MIN()`. **Completed** — `=MIN(A2:A51)` returns 25,000.
7. **Count Total Orders:** Count the records using `COUNT()` and/or `COUNTA()`. **Completed** — `=COUNT(A2:A51)` and `=COUNTA(A2:A51)` both return 50.
8. **Count Orders by City:** Count orders from Karachi, Lahore, and Islamabad using `COUNTIF()`. **Completed** — `=COUNTIF(Raw_Sales_Data!D:D,Analysis!C23)` (adjusting the criteria reference for each city) returns 20 for Karachi, 17 for Lahore, and 13 for Islamabad.
9. **Calculate Revenue by City:** Sum revenue for each city using `SUMIF()`. **Completed** — `=SUMIF(Raw_Sales_Data!D:D,C28,Analysis!A:A)` (adjusting the criteria reference for each city) returns 2,174,500 for Karachi, 1,906,000 for Lahore, and 1,109,300 for Islamabad.
10. **Calculate Revenue by Category:** Sum revenue for Electronics, Furniture, and Clothing using `SUMIF()`. **Completed** — `=SUMIF(Raw_Sales_Data!E$2:E$51,B33,Analysis!A$2:A$51)` (filled down for each category) returns 2,460,900 for Electronics, 2,039,000 for Furniture, and 689,900 for Clothing; the grand total is 5,189,800.

### Level 3 — Filtering and Conditional Analysis

11. **Find High-Value Orders:** Filter orders with revenue greater than 50,000. **Completed** — `=FILTER(Raw_Sales_Data!A2:J51,Analysis!A2:A51>50000,"No orders found")` returns all 10 source fields for the 37 orders whose revenue exceeds 50,000. The formula is entered in `Analysis!D2` and spills across the results area.
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

| Excel Skill                           | Questions         |
| ------------------------------------- | ----------------- |
| Basic formulas and calculated columns | 1–7               |
| `SUM()`                               | 2–3               |
| `AVERAGE()`                           | 4                 |
| `MAX()` / `MIN()`                     | 5–6               |
| `COUNT()` / `COUNTA()`                | 7                 |
| `COUNTIF()`                           | 8, 12, 17         |
| `SUMIF()`                             | 9, 10, 14, 16, 19 |
| `AVERAGEIF()`                         | 18                |
| `IF()` / nested `IF()`                | 15                |
| Filtering                             | 11, 13            |
| Sorting                               | 16                |
| Date analysis                         | 19                |
| PivotTables and PivotCharts           | 20                |
| Dashboard design                      | 20                |

## Recommended Progression

Work through the exercises in order:

1. Questions 1–5: Level 1 — Basic Excel Analysis — completed
2. Questions 6–10: Basic formulas, conditional counting, and conditional aggregation — completed
3. Questions 11–15: Filtering and logical conditions
4. Questions 16–19: Business-focused analysis
5. Question 20: Build a mini sales dashboard

The goal is to translate business questions into calculations and insights, rather than simply memorizing Excel functions.
