# Process
<p align="center">
<img width="3485" height="1678" alt="Picture1" src="https://github.com/user-attachments/assets/5266e1b6-22f0-42ec-adc0-b0e2382a7782" />
<img width="1866" height="877" alt="EXCEL_xb72hHQA9s" src="https://github.com/user-attachments/assets/f203e689-79be-4bc0-a00e-9ebd1236c082" />
</p>

*Gervon Alcide*

# Table of Contents
- [Methodology](#methodology)
- [Data Cleaning](#data-cleaning)
- [Analysis](#analysis)
  - [Automated Data Entry](#automated-data-entry)

[View Insights, Recommendations, and Visualisations](README.md)

---

# Methodology
This document covers the full technical process behind the analysis. For insights, recommendations, and dashboards, [see the README](README.md).

- I started by creating a backup. I clean anything I can, then reference a personal cleaning checklist to ensure thoroughness.
- For analysis, the main question I am answering is what drives or harms profit. I studied the relationships of various metrics and profit to draw insights.
- After analysis, I used VBA Automation to create buttons for swifter data entry.

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
	- Sort the table by order date ascending
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
<img width="1866" height="877" alt="EXCEL_xb72hHQA9s" src="https://github.com/user-attachments/assets/3b47c60b-2785-484c-9138-8285188c5567" />
</p>
<br>

---

- **Customer return analysis:**
	- Built a separate table to compare return rates for discounted vs non-discounted customers.
 
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
  		- Shows there's **no** correlation between more profit and more sales of tables.
	- Copier and Avg Profit
		- Shows that there is a correlation between more profit and more sales of copiers.

<p align="center">
<img width="3485" height="1678" alt="Picture1" src="https://github.com/user-attachments/assets/791cd4c4-61bd-47be-a625-a9002d70d5c1" />
<img width="916" height="485" alt="EXCEL_GPryFVqAc4" src="https://github.com/user-attachments/assets/6438eee6-fcf7-4b5c-a23b-de9236bcefda" />
<img width="3169" height="1667" alt="Picture2" src="https://github.com/user-attachments/assets/2f1d0733-68e8-4f1a-962f-a0f5ec2b4cb4" />
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

## Automated Data Entry
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
