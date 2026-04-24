<h1 align="center">Superstore Sales Analysis - Excel & Looker Studio</h1>

<p align="center">
<img width="1707" height="960" alt="chrome_Ihr749FO55" src="https://github.com/user-attachments/assets/132c7957-c548-4550-af82-87d35104e0fc" />
  <img width="1920" height="797" alt="EXCEL_Fw59VNkrT4" src="https://github.com/user-attachments/assets/3f3d2c06-2d75-42eb-a0bf-b3b92dd07e91" />
  <img width="1629" height="979" alt="EXCEL_kOygmOSviu" src="https://github.com/user-attachments/assets/cded3a5a-5cd8-4caa-ae16-2948365e9df6" />
<img width="1600" height="745" alt="EXCEL_CkJlrhZrQD" src="https://github.com/user-attachments/assets/c3f1f6b3-ea51-4834-a2e0-8b3aa980258a" />

</p>

*Gervon Alcide*

A full sales analysis of a US-based retail superstore covering 4 years of transaction data. I cleaned and structured the raw data in Excel, performed exploratory analysis using pivot tables, scatter plots, and graphs, built an interactive Excel dashboard with VBA automation, and a companion Looker Studio dashboard.

## Setup
**Tools:** Microsoft Excel, Google Looker Studio<br>
**Dataset:** [Superstore Dataset — Kaggle](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final) <br>
**Records:** 9,994 rows, 21 columns <br>
**Period:** January 2014 – December 2017 <br>
**Live Dashboard:** [View Interactive Dashboard](https://datastudio.google.com/reporting/a41c5dbd-f9b7-4bfd-a46b-1b02b15ad303) <br>
**Excel File:** [Superstore Dataver1.xlsm](https://github.com/user-attachments/files/27033504/Superstore.Dataver1.xlsm) <br>


---

## Data Cleaning
<img width="1344" height="613" alt="Pasted image 20260421083127" src="https://github.com/user-attachments/assets/ebcac47f-8d17-49c6-946d-e3e00fc2250f" />


### Steps for cleaning
- Verified no blank cells across all 21 columns:
```excel
=COUNTBLANK(A1:U9995)
```
- Confirmed no duplicate rows
- Verified all 9,994 order and ship dates are valid numbers:
```excel
=SUM(--ISNUMBER(C2:C9995))
```
- Confirmed no shipment arrived before its order date:
```excel
=SUM(IF(C2:C9995>D2:D9995, 1, 0))
```
- Verified no negative discounts and no quantities at or below zero:
```excel
=SUM(IF(S2:S9995 <= 0, 1, 0))
```
- Trimmed whitespace from 6 text columns (Product Name, Product ID, City, State, Customer Name, Customer ID) using `TRIM()`
- Formatted Discount column as percentage, Sales and Profit as currency
- Applied data validation to key columns:
  - `Order Date` and `Ship Date`: dates only
  - `Ship Mode`, `Segment`, `Country`, `City`, `State`, `Region`, `Category`, `Sub-Category`: dynamic dropdown lists
  - `Sales`, `Discount`, `Profit`: decimals only
  - `Quantity`: whole numbers only
- Converted dataset to a formal Excel Table for dynamic referencing
  
---

## Analysis

Added Expenses column
	=[@Sales] - [@Profit]
Added Cumulative Profit column
	Sort the table by order date low to high
	=IFERROR([@Profit] + $W1, [@Profit])
Added Profit Margin column
	=[@Profit] / [@Sales]

Pivot tables, scatter plots, trendlines, and forecasts built to answer each one.

**Pivot tables created:**
Average Profit:
1. Category
2. Ship Mode
3. Segment
4. Region
5. Sub-Category
6. State
<p align="center">
<img width="1777" height="472" alt="EXCEL_JiwirpUANd" src="https://github.com/user-attachments/assets/490a4d07-c719-41c7-9012-267a0e1c0ed4" />
</p>

Profit Margin:
1. Discount
2. Expenses (Grouped by 1,000s)
3. Quantity
4. Sub-Category
<p align="center">
<img width="1600" height="745" alt="EXCEL_CkJlrhZrQD" src="https://github.com/user-attachments/assets/4adb2e74-71ad-43ca-b78d-dfdbb92e4d15" />
</p>

**Customer return analysis:**
Built a separate table using `UNIQUE()`, `COUNTIFS()` to compare return rates for discounted vs non-discounted customers. 
  Customers who received discounts returned at an average rate of 6.5x vs 6x without discounts.
<p align="center">
  <img width="836" height="89" alt="EXCEL_hgV6UE0CdT" src="https://github.com/user-attachments/assets/8198b869-2794-4039-93e1-222eea4ce3b9" />
</p>

**Top and bottom month analysis:**
Used `LARGE()`, `INDEX()`, and `MATCH()` to identify the 10 highest and lowest performing months by profit and cross-referenced against discount averages.
<p align="center">
<img width="1204" height="532" alt="EXCEL_hytvKpEKBm" src="https://github.com/user-attachments/assets/29522549-9d61-49a2-8e54-9dbd2fc86508" />
</p>


**Forecast:**
Built 2 projections — cumulative profit and month-by-month profit.
<p align="center">
<img width="1769" height="598" alt="EXCEL_sjhALfocnz" src="https://github.com/user-attachments/assets/c0cd36e6-1388-4e94-aa35-7da9c075864c" />
</p>

<p align="center">

</p>

---

## Insights

1. Technology was the most profitable category. Copiers led all sub-categories by a significant margin.
2. The West region was the most profitable. The Central region was the least.
3. Standard Class was the most profitable ship mode.
4. 5 of the 10 worst profit months coincided with the 10 highest average discount months. High discounts reliably hurt performance but low discounts alone don't guarantee great performing months.
5. Transactions at or above a 30% discount rate consistently produced losses. Removing all 30%+ discount transactions would have added approximately $135k to total profit resulting to a total profit increase from $286 to $422k.
7. The Tables sub-category lost money every single month without exception. Removing Tables entirely would have recovered ~$18,000, bringing total profit from $286k to $304k.
8. Cumulative profit trended positively across all 4 years. Month over month profit is volatile, however it still shows positive growth over time, indicating overall business growth despite the drag from high discounts and underperforming sub-categories.

---

## Dashboard

Built an interactive Excel dashboard across 5 sheets with a navigatable homepage and top navigation bar on all pages.

**Visuals:**
- USA map: green states profitable, red states at a loss
- Donut chart: share of transactions with 30%+ discounts
- Donut char: profit lost to 30%+ discount transactions vs total profit
- Sub-category by profit bar chart
- Sub-category by units sold bar chart
- Monthly profit and cumulative profit line chart

**Interactive Controls:**
- Date slicer
- Category slicer
- Segment slicer
- Region slicer
- Ship Mode slicer





**VBA Automation:**
- Down arrow button on the data sheet automatically highlights the next empty row for fast data entry
- Up arrow button returns to the top of the dataset instantly

Built a companion Looker Studio dashboard with 5 KPI cards (Total Sales, Total Expenses, Total Profit, Quantity Sold, Profit Margin) and filters for Month, Year, Segment, Category, and Region.

**[→ View Live Looker Studio Dashboard](https://datastudio.google.com/reporting/a41c5dbd-f9b7-4bfd-a46b-1b02b15ad303)**

---

## Recommendations

1. **Cap discounts at 29%.** Losses begin consistently at 30%+. A hard cap on discount approvals is the single highest-leverage change available — worth an estimated $135,000 in recovered profit.
2. **Review the Tables sub-category.** Tables bled money every month for 4 years. Whether to discontinue, reprice, or restrict discounts specifically on Tables is worth a serious internal conversation.
3. **Investigate the Central region.** Underperformance there may reflect discount policy, product mix, or regional demand — worth isolating before drawing conclusions.
