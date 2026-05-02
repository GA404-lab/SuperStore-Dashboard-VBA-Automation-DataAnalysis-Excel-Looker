<h1 align="center">Superstore Sales Analysis - Excel & Looker Studio</h1>
<p align="center">
<img width="1920" height="797" alt="ezgif-4" src="https://github.com/user-attachments/assets/f02cfe45-99c3-47d7-b40a-b1c90aa578f2" />
<img width="1707" height="960" alt="chrome_Ihr749FO55" src="https://github.com/user-attachments/assets/132c7957-c548-4550-af82-87d35104e0fc" />
<img width="1662" height="977" alt="EXCEL_lbkwzKQwNm" src="https://github.com/user-attachments/assets/ec0c134e-3a2e-4a0a-8c1e-432a824826b7" />
<img width="1014" height="488" alt="EXCEL_mBlDHeAsf8" src="https://github.com/user-attachments/assets/043487ed-db65-4cad-ae80-9f8d28dd6f85" />
<img width="1600" height="745" alt="EXCEL_CkJlrhZrQD" src="https://github.com/user-attachments/assets/c3f1f6b3-ea51-4834-a2e0-8b3aa980258a" />

</p>

*Gervon Alcide*


A full sales analysis of a US-based retail superstore covering 4 years of real transaction data. I cleaned and structured the raw data in Excel, performed exploratory analysis using pivot tables, scatter plots, and graphs, built an interactive Excel dashboard with VBA automation, and a companion Looker Studio dashboard.

## Setup
**Tools:** Microsoft Excel, Google Looker Studio<br>
**Dataset:** [Superstore Dataset Kaggle](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final) <br>
**Live Dashboard:** [View Interactive Dashboard](https://datastudio.google.com/reporting/a41c5dbd-f9b7-4bfd-a46b-1b02b15ad303) <br>
**Excel File:** [Superstore Dataset .xlsm](https://github.com/user-attachments/files/27057821/Superstore.Dataset.xlsm)

---

# Data Cleaning
<p align="center">

<img width="1344" height="613" alt="Pasted image 20260421083127" src="https://github.com/user-attachments/assets/ebcac47f-8d17-49c6-946d-e3e00fc2250f" />
</p>


## Steps for cleaning

9,994 rows, 21 columns<br>
Data spans January 2014 to December 2017 <br>

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
- Converted dataset to a formal Excel Table and added named ranges for dynamic referencing
  
---

# Analysis

- Added Expenses column
- Added Cumulative Profit column
	- Sort the table by order date low to high
```excel
 =IFERROR([@Profit] + $W1, [@Profit])
 ```
- Added Profit Margin column

Pivot tables, scatter plots, trendlines, and forecasts built to answer each one.

## **Pivot tables, scatterplots, trendlines and forecasts:**

- **Various metrics against Average Profit:** 
	- Category
	- Ship Mode
	- Segment
	- Region
	- Sub-Category
	- State
<p align="center">
<img width="1777" height="472" alt="EXCEL_JiwirpUANd" src="https://github.com/user-attachments/assets/490a4d07-c719-41c7-9012-267a0e1c0ed4" />
</p>

<br>

- **Various metrics against Profit Margin:**
	- Discount
	- Expenses (Grouped by 1,000s)
	- Quantity
	- Sub-Category

<p align="center">
<img width="1600" height="745" alt="EXCEL_CkJlrhZrQD" src="https://github.com/user-attachments/assets/4adb2e74-71ad-43ca-b78d-dfdbb92e4d15" />
</p>
<br>

---

- **Customer return analysis:**
	- Built a separate table using to compare return rates for discounted vs non-discounted customers.
 
- Column 1:
	- A list of all customers without duplicates
```excel
	{=UNIQUE(Customer_ID)}
```
- Column 2:
	- Counts how many times a customer showed up when they received a discount
```excel
	=COUNTIFS(Customer_ID, Sheet2!C2, Discount, ">" & 0)
```

- Column 3:
	- This counts how many times a customer showed up when they didn't receive a discount
```excel
	=COUNTIFS(Customer_ID, Sheet2!C2, Discount, "="&0)
```
**Customers who received discounts returned at an average rate of 6.5x vs 6x without discounts.**
<p align="center">
  <img width="836" height="89" alt="EXCEL_hgV6UE0CdT" src="https://github.com/user-attachments/assets/8198b869-2794-4039-93e1-222eea4ce3b9" />
</p>

---

- Made 3 scatterplots for analysis.
	- Profit and Avg Discount
		- Shows a negative correlation with discount rate and profit
	- Table and Avg Profit
  		- Shows theres **no** correlation between more profit and more sales of tables.
	- Copier and Avg Profit
		- Shows that there is a correlation between more profit and more sales of copiers.

<p align="center">
<img width="1014" height="488" alt="EXCEL_mBlDHeAsf8" src="https://github.com/user-attachments/assets/36d56b9b-c16e-41f8-a418-ad37da998555" />
<img width="916" height="485" alt="EXCEL_GPryFVqAc4" src="https://github.com/user-attachments/assets/6438eee6-fcf7-4b5c-a23b-de9236bcefda" />
<img width="922" height="485" alt="EXCEL_4dtHy1W0yo" src="https://github.com/user-attachments/assets/e215e6f2-041d-4363-888a-00b783da25aa" />
</p>

---

**Top and bottom month analysis:**<br>
Used `LARGE()` / `SMALL()`, `INDEX()`, and `MATCH()` to identify the 10 highest and lowest performing months by profit and average discount, then compared the lists for commonalities.
```excel
=LARGE($F$128:$F$175, H128)
```
```excel
=INDEX($E$128:$E$175,MATCH($I128,$F$128:$F$175,0))
```

<p align="center">
<img width="1204" height="532" alt="EXCEL_hytvKpEKBm" src="https://github.com/user-attachments/assets/42b3ee17-c449-430e-8ccc-5e7c27b65ee1" />
</p>

---

**Forecast:**<br>
Built 2 projections:
- cumulative profit
- month-by-month profit
<p align="center">
	<img width="1769" height="598" alt="EXCEL_sjhALfocnz" src="https://github.com/user-attachments/assets/a9ba6d87-1452-473b-93de-87df17667530" />
</p>

---

**Formatting**<br>
Pivot tables, charts, and supporting analysis are organized across dedicated sheets by topic for easy navigation and reference.
<p align="center">
<img width="1920" height="797" alt="ezgif-1" src="https://github.com/user-attachments/assets/83e0c541-751f-43cb-bbf2-06627a22e5e3" />

</p>

---
---

## Insights

1. Technology was the most profitable category. Copiers led all sub-categories by a significant margin.
2. The West region was the most profitable. The Central region was the least.
3. Texas was the worst perfoming state, losing the company $25k in total. California was the best performing state.
4. 5 of the 10 worst profit months were also part of the top 10 highest average discount months. High discounts reliably hurt performance but low discounts alone don't guarantee great performing months.
5. Customers are almost just as likely to be recurring customers with or without discounts.
6. Transactions at or above a 30% discount rate consistently produced losses. Removing all 30%+ discount transactions would have added approximately $135k to total profit resulting to a total profit increase from $286k to $422k.
7. The Tables sub-category lost money every single month without exception. Removing Tables entirely would have recovered ~$18,000, bringing total profit from $286k to $304k.
8. Cumulative profit trended positively across all 4 years. Month over month profit is volatile, however it still shows positive growth over time, indicating overall business growth despite the drag from high discounts and underperforming sub-categories.

### Most Important Insight
Discount rates 30% and higher are extremely detrimental to the businesses, consistently making a loss. <br>
Keeping an eye on discount related metrics are important and will be included in the dashboard.
<p align="center">
<img width="461" height="719" alt="EXCEL_dDqHyRi6WK" src="https://github.com/user-attachments/assets/c11cfc4b-a141-404d-9e65-8c6c6a936a55" />
</p>

---

# Dashboard, Automation and Navigation

## Dashboard
Built an interactive Excel dashboard with a navigatable homepage and top navigation bar on all pages.

**KPI:**
- Sales
- Expenses
- Profit
	- Added conditional formatting to make the text red when its negative and green when its positive
 - Quantity Sold
 - AVG Profit Margin
 	- Added conditional formatting to make the text red when under 10% and green when above 10%

**Visuals:**
- USA map: green states profitable, red states at a loss
- Donut chart: share of discounts with a 30%+ rate
- Donut chart: profit lost to 30%+ discount transactions vs total profit
- Sub-category by profit bar chart
- Sub-category by units sold bar chart
- Monthly profit and cumulative profit line chart

**Interactive Controls:**
- Date slicer
- Category slicer
- Segment slicer
- Region slicer
- Ship Mode slicer
<p align="center">
<img width="1920" height="1010" alt="ezgif-ExcelDashboard2" src="https://github.com/user-attachments/assets/dacccd7c-5d13-4d4f-b14b-18698662c058" />

</p>

*full picture of the dashboard is at the top of the file*

---

## VBA Automation
- Down arrow button on the data sheet automatically highlights the next empty row for swifter data entry
- Up arrow button returns to the top of the dataset instantly
```vba
Sub NavigatingDown()
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub

Sub NavigatingUP()
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("SuperStoreData[[#Headers],[Row ID]]").Select
End Sub
```
<p align="center">
<img width="1920" height="797" alt="ezgif-3" src="https://github.com/user-attachments/assets/454b47bf-8c9d-4e19-9aee-66ce5b828845" />
</p>

---

## Navigation
The workbook is designed to feel like a simple app rather than a raw Excel file.

- Dedicated Homepage serves as the entry point to the workbook
- All pages feature a top navigation bar for jumping between sheets instantly
- The Data Entry sheet includes a formatted navigation bar alongside the VBA automation buttons
- Key sheets are locked to prevent accidental edits while keeping slicers and navigation fully functional

<p align="center">
<img width="1920" height="797" alt="ezgif-4" src="https://github.com/user-attachments/assets/933ea04d-7379-4b6c-b07a-86832fe29d75" />

</p>

---
---

# Looker Studio Dashboard

Built a companion Looker Studio dashboard alongside the Excel workbook.

**KPIs:**
- Total Sales
- Total Expenses
- Total Profit
- Quantity Sold
- Profit Margin

**Visuals:**
- Donut chart: total profit by category
- Quantity sold and total profit by sub-category
- Profit margin and sales by sub-category
- USA map: higher profit states appear blue, lower profit states appear purple
- Profit per month and cumulative profit line chart

**Filters:**
- Month
- Year
- Segment
- Category
- Region

<p align="center">
<img width="1708" height="960" alt="ezgif-5" src="https://github.com/user-attachments/assets/2b5795d9-8da6-4784-b842-8d8d3d4f0e20" />

</p>

**[View Live Looker Studio Dashboard Here](https://datastudio.google.com/reporting/a41c5dbd-f9b7-4bfd-a46b-1b02b15ad303)**

---

# Recommendations

1. **Cap discounts at 29%.** Losses begin consistently at 30%+. A hard cap on discount approvals is the single highest-leverage change available. Worth an estimated $135k in profit.
2. **Review the Tables sub-category.** Tables bled money every month for 4 years. Whether to discontinue, reprice, or restrict discounts specifically on Tables is worth a serious discussion.
3. **Investigate the Central region.** Underperformance in the central region, especially Texas, may reflect discount policy, product mix, or regional demand. Worth looking deeper before drawing conclusions.
