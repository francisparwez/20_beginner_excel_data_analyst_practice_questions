# 20 Beginner Excel Data Analyst Practice Questions

## Project Overview

A beginner-friendly Excel practice project focused on answering common sales-analysis questions using formulas, filters, sorting, date functions, PivotTables, PivotCharts, and a basic dashboard.

## Dataset

The `sales-dataset.xlsx` workbook contains four sheets: `Raw_Sales_Data`, with 50 sales records and the 10 original source columns plus a calculated `Revenue` column; `Analysis_01`, which contains summary metrics, revenue breakdowns, and filtered order results; `Analysis_02`, which contains Electronics revenue and order-size classifications; and `Level_4_Practical_Data_Analysis`, which calculates revenue by salesperson. The 10 original source columns are:

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

The `Raw_Sales_Data` sheet includes the calculated `Revenue` column. `Analysis_01` contains summary metrics, revenue breakdowns by city and category, filtered high-value order results, and filtered Electronics order results. `Analysis_02` contains total Electronics revenue and an `Order_Size` classification for each order. `Level_4_Practical_Data_Analysis` lists the four distinct salespeople and calculates each person's total revenue. The dataset covers January–August 2026 and includes orders from Karachi, Lahore, and Islamabad; Electronics, Furniture, and Clothing products; four salespeople; and Card, Cash, and Bank Transfer payments.

`Revenue` is calculated in `Raw_Sales_Data!K2` using `=Raw_Sales_Data!H2*Raw_Sales_Data!G2` and filled down for all 50 orders. `Analysis_01` calculates total revenue with `=SUM(Raw_Sales_Data!K2:K51)`, total quantity sold with `=SUM(Raw_Sales_Data!G2:G51)`, average order revenue with `=AVERAGE(Raw_Sales_Data!K:K)`, the highest order value with `=MAX(Raw_Sales_Data!K:K)`, and the lowest order value with `=MIN(Raw_Sales_Data!K2:K51)`. It also counts orders with `=COUNT(Raw_Sales_Data!K2:K51)` and `=COUNTA(Raw_Sales_Data!K2:K51)`, and counts orders by city with `=COUNTIF(Raw_Sales_Data!D:D,Analysis_01!B23)` (adjusting the criteria reference for each city). Revenue by city is calculated with `=SUMIF(Raw_Sales_Data!D:D,B28,Raw_Sales_Data!K:K)` and the criteria reference is adjusted for each city. Revenue by category is calculated with `=SUMIF(Raw_Sales_Data!E$2:E$51,A33,Raw_Sales_Data!K$2:K$51)` and filled down for each category. `Analysis_02` calculates Electronics revenue with `=SUMIF(Raw_Sales_Data!E2:E51,"Electronics",Raw_Sales_Data!K2:K51)` and classifies each order using nested `IF()` formulas. The verified results are **5,189,800 total revenue**, **300 units sold**, **103,796 average order revenue**, **324,000 highest order value**, **25,000 lowest order value**, **50 total orders**, city counts of **20 Karachi**, **17 Lahore**, and **13 Islamabad**, city revenue of **2,174,500 Karachi**, **1,906,000 Lahore**, and **1,109,300 Islamabad**, category revenue of **2,460,900 Electronics**, **2,039,000 Furniture**, and **689,900 Clothing**, **37 high-value orders** with revenue greater than 50,000, and **19 Electronics orders**. Electronics revenue is **2,460,900**. Salesperson revenue totals are **1,789,800 Ahmed**, **1,662,000 Bilal**, **1,056,500 Hamza**, and **681,500 Sana**; Ahmed generated the highest revenue.

## Practice Questions

### Level 1 — Basic Excel Analysis

1. **Calculate Total Revenue:** Calculate revenue for each order using `Quantity × Unit_Price`. **Completed** — the `Revenue` column uses `=Raw_Sales_Data!H2*Raw_Sales_Data!G2` and is filled down for all 50 orders.
2. **Calculate Total Sales:** Sum the revenue for all orders using `SUM()`. **Completed** — `=SUM(Raw_Sales_Data!K2:K51)` returns 5,189,800.
3. **Calculate Total Quantity Sold:** Sum all units sold using `SUM()`. **Completed** — `=SUM(Raw_Sales_Data!G2:G51)` returns 300.
4. **Calculate Average Order Revenue:** Calculate average revenue per order using `AVERAGE()`. **Completed** — `=AVERAGE(Raw_Sales_Data!K:K)` returns 103,796.
5. **Find the Highest Order Value:** Find the maximum order revenue using `MAX()`. **Completed** — `=MAX(Raw_Sales_Data!K:K)` returns 324,000.

### Level 2 — Basic Data Analysis

6. **Find the Lowest Order Value:** Find the minimum order revenue using `MIN()`. **Completed** — `=MIN(Raw_Sales_Data!K2:K51)` returns 25,000.
7. **Count Total Orders:** Count the records using `COUNT()` and/or `COUNTA()`. **Completed** — `=COUNT(Raw_Sales_Data!K2:K51)` and `=COUNTA(Raw_Sales_Data!K2:K51)` both return 50.
8. **Count Orders by City:** Count orders from Karachi, Lahore, and Islamabad using `COUNTIF()`. **Completed** — `=COUNTIF(Raw_Sales_Data!D:D,Analysis_01!B23)` (adjusting the criteria reference for each city) returns 20 for Karachi, 17 for Lahore, and 13 for Islamabad.
9. **Calculate Revenue by City:** Sum revenue for each city using `SUMIF()`. **Completed** — `=SUMIF(Raw_Sales_Data!D:D,B28,Raw_Sales_Data!K:K)` (adjusting the criteria reference for each city) returns 2,174,500 for Karachi, 1,906,000 for Lahore, and 1,109,300 for Islamabad.
10. **Calculate Revenue by Category:** Sum revenue for Electronics, Furniture, and Clothing using `SUMIF()`. **Completed** — `=SUMIF(Raw_Sales_Data!E$2:E$51,A33,Raw_Sales_Data!K$2:K$51)` (filled down for each category) returns 2,460,900 for Electronics, 2,039,000 for Furniture, and 689,900 for Clothing; the grand total is 5,189,800.

### Level 3 — Filtering and Conditional Analysis

11. **Find High-Value Orders:** Filter orders with revenue greater than 50,000. **Completed** — `=FILTER(Raw_Sales_Data!A2:J51,Raw_Sales_Data!K2:K51>50000,"No orders found")` returns all 10 original source fields for the 37 orders whose revenue exceeds 50,000. The formula is entered in `Analysis_01!C2` and spills across the results area.
12. **Count High-Value Orders:** Count orders with revenue greater than 50,000 using `COUNTIF()`. **Completed** — `=COUNTIF(Raw_Sales_Data!K2:K51,">50000")` returns 37.
13. **Find Electronics Orders:** Filter the dataset to show only Electronics orders. **Completed** — `=VSTACK(Raw_Sales_Data!A1:J1,FILTER(Raw_Sales_Data!A2:J51,Raw_Sales_Data!E2:E51="Electronics","No Electronics orders found"))` returns the headers and all 10 original source fields for the 19 Electronics orders. The formula is entered in `Analysis_01!B44`.
14. **Calculate Electronics Revenue:** Sum revenue for Electronics using `SUMIF()`. **Completed** — `=SUMIF(Raw_Sales_Data!E2:E51,"Electronics",Raw_Sales_Data!K2:K51)` in `Analysis_02!A2` returns 2,460,900.
15. **Identify Large Orders:** Add an `Order_Size` column: `Large` for revenue ≥ 100,000, `Medium` for revenue ≥ 50,000 but below 100,000, and `Small` for revenue below 50,000. Use `IF()` or nested `IF()`. **Completed** — `=IF(Raw_Sales_Data!K2>=100000,"Large",IF(Raw_Sales_Data!K2>=50000,"Medium","Small"))` in `Analysis_02!B2` is filled down through `B51`.

### Level 4 — Practical Data Analyst Questions

16. **Find the Best Salesperson:** Calculate revenue by salesperson using `SUMIF()`, then sort or compare totals to identify the highest. **Completed** — `=UNIQUE(Raw_Sales_Data!I2:I51)` in `Level_4_Practical_Data_Analysis!A2` returns the distinct salespeople, and `=SUMIF(Raw_Sales_Data!I2:I51,A2,Raw_Sales_Data!K2:K51)` in `B2` is filled down through `B5` to calculate each person's revenue. Comparing the totals identifies Ahmed as the top salesperson with 1,789,800.
17. **Analyze Payment Methods:** Count Card, Cash, and Bank Transfer orders using `COUNTIF()` and identify the most frequently used method.
18. **Calculate Average Revenue by Category:** Calculate average order revenue for each category using `AVERAGEIF()`.
19. **Analyze Monthly Sales:** Extract the month from `Order_Date` and calculate revenue by month using date functions with `SUMIF()` or a PivotTable.
20. **Create a Basic Sales Dashboard:** Build KPI summaries for Total Revenue, Total Orders, Total Units Sold, and Average Order Value. Add charts for revenue by category, city, month, and salesperson using PivotTables and PivotCharts.

## Skills Practiced

| Excel Skill                           | Questions         |
| ------------------------------------- | ----------------- |
| Basic formulas and calculated columns | 1–7, 15           |
| `SUM()`                               | 2–3               |
| `AVERAGE()`                           | 4                 |
| `MAX()` / `MIN()`                     | 5–6               |
| `COUNT()` / `COUNTA()`                | 7                 |
| `COUNTIF()`                           | 8, 12, 17         |
| `SUMIF()`                             | 9, 10, 14, 16, 19 |
| `AVERAGEIF()`                         | 18                |
| `IF()` / nested `IF()`                | 15                |
| `UNIQUE()`                            | 16                |
| Filtering                             | 11, 13            |
| Date analysis                         | 19                |
| PivotTables and PivotCharts           | 20                |
| Dashboard design                      | 20                |

## Recommended Progression

Work through the exercises in order:

1. Questions 1–5: Level 1 — Basic Excel Analysis — completed
2. Questions 6–10: Basic formulas, conditional counting, and conditional aggregation — completed
3. Questions 11–15: Filtering and logical conditions — completed
4. Question 16: Find the Best Salesperson — completed; Questions 17–19: remaining business-focused analysis
5. Question 20: Build a mini sales dashboard

The goal is to translate business questions into calculations and insights, rather than simply memorizing Excel functions.
