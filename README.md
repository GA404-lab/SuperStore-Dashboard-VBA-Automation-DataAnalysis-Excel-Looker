<h1 align="center">Superstore Sales Analysis - Excel & Looker Studio</h1>
<p align="center">
<img width="1920" height="797" alt="ezgif-4" src="https://github.com/user-attachments/assets/f02cfe45-99c3-47d7-b40a-b1c90aa578f2" />
<img width="1707" height="960" alt="chrome_Ihr749FO55" src="https://github.com/user-attachments/assets/132c7957-c548-4550-af82-87d35104e0fc" />
<img width="1662" height="977" alt="EXCEL_lbkwzKQwNm" src="https://github.com/user-attachments/assets/ec0c134e-3a2e-4a0a-8c1e-432a824826b7" />
</p>

*Gervon Alcide*


A full sales analysis of a US-based retail superstore covering 4 years of real transaction data. I cleaned and structured the raw data in Excel, performed exploratory analysis using pivot tables, scatter plots, and graphs, built an interactive Excel dashboard with VBA automation, and a companion Looker Studio dashboard.

## Setup
**Tools:** Microsoft Excel, Google Looker Studio<br>
**Dataset:** [Superstore Dataset Kaggle](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final) <br>
**Live Dashboard:** [View Interactive Dashboard](https://datastudio.google.com/reporting/a41c5dbd-f9b7-4bfd-a46b-1b02b15ad303) <br>
**Excel File:** [Superstore Dataset .xlsm](https://github.com/user-attachments/files/27057821/Superstore.Dataset.xlsm)

---
# Table of Contents

- [Methodology](#methodology)
- [Insights](#insights)
- [Recommendations](#recommendations)
- [Dashboard, Automation and Navigation](#dashboard-automation-and-navigation)
- [Looker Studio Dashboard](#looker-studio-dashboard)

[View full documented process](Process.md)

---
# Methodology

This analysis examined the relationships of various metrics to answer: what is actually driving profit, and what is reducing it. I explored that through product performance, regional trends, discount behaviour, and customer patterns using pivot tables, scatter plots, trendlines, and forecasting. You can view the full documented process of cleaning, analysing, and creating automation [here](Process.md).
<p align="center">
<img width="1920" height="797" alt="ezgif-1" src="https://github.com/user-attachments/assets/83e0c541-751f-43cb-bbf2-06627a22e5e3" />
</p>

---
# Insights

1. Technology was the most profitable category. Copiers led all sub-categories by a significant margin.
2. The West region was the most profitable. The Central region was the least.
3. Texas was the worst performing state, losing the company $25k in total. California was the best performing state.
4. 5 of the 10 worst profit months were also part of the top 10 highest average discount months. High discounts reliably hurt performance but low discounts alone don't guarantee great performing months.
5. Customers are almost just as likely to be recurring customers with or without discounts.
6. Transactions at or above a 30% discount rate consistently produced losses. Removing all 30%+ discount transactions would have added approximately $135k to total profit resulting to a total profit increase from $286k to $422k.
7. The Tables sub-category lost money every single month without exception. Removing Tables entirely would have recovered ~$18,000, bringing total profit from $286k to $304k.
8. Cumulative profit trended positively across all 4 years. Month over month profit is volatile, however it still shows positive growth over time, indicating overall business growth despite the drag from high discounts and underperforming sub-categories.

### Most Important Insight
Discount rates 30% and higher are extremely detrimental to the businesses, consistently making a loss. <br>
Keeping an eye on discount related metrics are important and will be included in the dashboard.
<p align="center">
<img width="461" height="719" alt="newnewnewnewnwnewrnewrwe" src="https://github.com/user-attachments/assets/435d2a05-2aae-476a-98fe-08dabdb3b5e5" />
</p>

---

# Recommendations

1. **Cap discounts at 29%.** Losses begin consistently at 30%+. A hard cap on discount approvals is the single highest-leverage change available. Worth an estimated $135k in profit.
2. **Review the Tables sub-category.** Tables bled money every month for 4 years. Whether to discontinue, reprice, or restrict discounts specifically on Tables is worth a serious discussion.
3. **Investigate the Central region.** Underperformance in the central region, especially Texas, may reflect discount policy, product mix, or regional demand. Worth looking deeper before drawing conclusions.

---

# Dashboard, Automation and Navigation

## Dashboard
Built an interactive Excel dashboard with a navigatable homepage and top navigation bar on all pages. Users can filter by date, category, segment, region and ship mode.

<p align="center">
<img width="1920" height="1010" alt="ezgif-ExcelDashboard2" src="https://github.com/user-attachments/assets/dacccd7c-5d13-4d4f-b14b-18698662c058" />

</p>

*still image of the dashboard is at the top of the file*

---

## VBA Automation

Buttons that takes you to the next empty row for swift data entry.

<p align="center">
<img width="1920" height="797" alt="ezgif-3" src="https://github.com/user-attachments/assets/454b47bf-8c9d-4e19-9aee-66ce5b828845" />
</p>

---

## Navigation
The workbook is designed to feel like a simple app rather than a raw Excel file.

- Dedicated Homepage.
- All pages feature a top navigation bar.
- Key sheets are locked to prevent accidental edits while keeping slicers and navigation fully functional.

<p align="center">
<img width="1920" height="797" alt="ezgif-4" src="https://github.com/user-attachments/assets/933ea04d-7379-4b6c-b07a-86832fe29d75" />

</p>

---

## Looker Studio Dashboard

Built a companion interactive Looker Studio web dashboard for sharing results without needing an Excel workbook.

<p align="center">
<img width="1708" height="960" alt="ezgif-5" src="https://github.com/user-attachments/assets/2b5795d9-8da6-4784-b842-8d8d3d4f0e20" />
</p>

**[View Live Looker Studio Dashboard Here](https://datastudio.google.com/reporting/a41c5dbd-f9b7-4bfd-a46b-1b02b15ad303)**

*still image of the dashboard is at the top of the file*
