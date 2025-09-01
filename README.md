# automated-financial-tracker-vba-powerbi

# Automated Financial Tracker (Excel VBA + Power BI)

📊 Automated financial performance tracker using Excel (Power Query & VBA) and Power BI.

## Project Overview
This project demonstrates:
- Excel Power Query for data cleaning and transformation
- VBA macros to automate monthly refresh and reporting
- Pivot Tables for summaries
- Power BI dashboards for interactive insights

## 📂 Repo Structure
automated-financial-tracker-vba-powerbi/
├── data/                   # synthetic dataset
│   └── financial_data.csv
├── excel/                  # Excel tracker with VBA macros
│   └── financial_tracker.xlsm
├── powerbi/                # Power BI dashboard
│   └── financial_dashboard.pbix
├── images/                 # screenshots
│   └── (screenshots go here)
└── README.md




## 🛠️ Dataset Plan

File: /data/financial_data.csv

Columns:

- Date → Monthly dates from Jan 2023 to Dec 2024
- Department → Sales, Marketing, IT, HR, Operations
- Revenue → Random between 100,000 – 500,000 (varies by department)
- Expenses → 50–90% of Revenue (to simulate realistic costs)
- Budget → Revenue ± 10–20% (to simulate over/under performance)

This gives ~120 rows (24 months × 5 departments).

Here’s a preview of first few rows:

``` 
Date,Department,Revenue,Expenses,Budget
2023-01-01,Sales,452310,354827,490128
2023-01-01,Marketing,212480,161234,231256
2023-01-01,IT,178342,121220,195150
2023-01-01,HR,95480,74322,103750
2023-01-01,Operations,305600,228457,325789
2023-02-01,Sales,468920,376527,430112
```

## Status
- [ ] Dataset ready
- [ ] Excel template built
- [ ] VBA automation added
- [ ] Power BI dashboard created
- [ ] Documentation finalized
