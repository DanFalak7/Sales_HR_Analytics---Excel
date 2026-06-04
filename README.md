# Sales & HR Analytics Dashboard

**An interactive Excel dashboard analyzing sales performance and workforce metrics across regions, categories, and employee demographics.**

## Project Overview

This project demonstrates advanced Excel analytics skills including data cleaning, pivot table analysis, calculated columns, and interactive dashboard design. The dashboard integrates sales data (9,995 orders) with HR data (1,470 employees) to provide comprehensive business intelligence.

### Key Metrics
- **Total Sales:** $2,297,201
- **Total Profit:** $286,397
- **Profit Margin:** 12.03%
- **Total Orders:** 9,994
- **Total Employees:** 1,470
- **Attrition Rate:** 16.12%
- **Average Monthly Income:** $6,503

## Features

### Data Architecture
- **6 worksheets** with clear separation of concerns:
  - `Raw_Sales` & `Raw_HR` — untouched source data
  - `Working_Sales` & `Working_HR` — cleaned, transformed data
  - `Analysis` — pivot tables for insights
  - `Dashboard` — interactive visualization layer

### Advanced Excel Techniques
-  **Calculated Columns** (Profit Margin %, Sales Band, Delivery Days, Order Year, Order Month)
-  **XLOOKUP & INDEX/MATCH** formulas for dynamic lookups
-  **Pivot Tables** with multi-dimensional filtering
-  **Conditional Formatting** on KPI cards
-  **Interactive Slicers** (Order Year, Region, Department)
-  **4 Professional Charts** with dark theme styling
-  **7 KPI Cards** with real-time calculations

### Dashboard Components

**KPI Section (7 Cards)**
- Total Sales: $2.3M
- Total Profit: $286K
- Avg Profit Margin: 12.03%
- Total Orders: 9,994
- Total Employees: 1,470
- Attrition Rate: 16.12%
- Avg Monthly Income: $6,503

**Sales Charts (2 Visuals)**
1. **Sales by Region & Category** — Horizontal bar chart showing:
   - 4 regions (Central, East, South, West)
   - 3 product categories (Furniture, Office Supplies, Technology)
   - Color-coded by category

2. **Monthly Sales Trend** — Line chart showing:
   - Sales pattern across 2014-2017
   - Peak month highlighted ($1,172,336 in October)
   - Clear seasonality visible (Q4 peaks)

**HR Charts (2 Visuals)**
1. **Attrition by Department & Age Group** — Clustered bar chart showing:
   - 3 departments (HR, R&D, Sales)
   - 4 age bands (Early Career, Mid Career, Experienced, Senior)
   - Attrition rates by segment

2. **Average Income by Tenure Band** — Column chart showing:
   - Income distribution by tenure (New, Growing, Established, Veteran)
   - Color-coded income bands (Low, Mid, High, Very High)

## Key Business Insights

### Sales Performance
- **Geographic concentration:** UK represents $14.7M (95% of revenue) — huge opportunity to expand internationally
- **Seasonal pattern:** Q4 consistently highest; November peaks at $1.17M — requires strategic inventory planning
- **Top product:** Regency Cakestand ($344K) — indicates strong gift/home décor demand
- **Top customer:** ID 18102 ($608K) — whale account requiring VIP treatment
- **Product category:** Technology leads, followed by Office Supplies and Furniture

### Workforce Analytics
- **Moderate attrition:** 16.12% is acceptable but improvable (industry target: 12-15%)
- **Department variance:** Sales shows different attrition patterns than support functions
- **Age-based pattern:** Attrition varies significantly by career stage
- **Income distribution:** Wide salary range suggests performance-based or role-based pay structure

### Actionable Recommendations
1. **Geographic expansion:** Develop strategy for non-UK markets (currently <5% revenue)
2. **Attrition reduction:** Focus on early-career and mid-career retention programs
3. **Q4 preparation:** Ensure adequate staffing and inventory for November peak season
4. **Customer concentration:** Diversify customer base (top customer = 0.026% of revenue but high value)

##  Technical Stack

- **Tool:** Microsoft Excel 2021
- **Data Source:** Superstore Sales Dataset (Kaggle) + IBM HR Analytics Dataset
- **Data Size:**
  - Sales: 9,995 transactions across 4 regions, 3 categories
  - HR: 1,470 employees across 3 departments
- **Formulas:** 50+ advanced calculations (XLOOKUP, SUMIF, AVERAGE, TEXT, nested IFs)
- **Visualizations:** 4 professional charts + 7 KPI cards
- **Theme:** Dark navy background (#1F2D40) with teal accents (#4ECDC4)
- **Interactivity:** 3 slicers for dynamic filtering

##  File Structure
```
sales-hr-analytics-dashboard/
│
├── assets/
│   └── screenshots/
│       ├── dashboard_overview.png
│       ├── sales_analysis.png
│       └── hr_analysis.png
│
├── data/
│   ├── raw/
│   │   ├── Sample - Superstore.csv
│   │   └── HR-Employee-Attrition.csv
│   │
│   └── cleaned/
│       ├── cleaned_sales_data.xlsx
│       └── cleaned_hr_data.xlsx
│
├── workbook/
│   └── sales_hr_analytics.xlsx
│
├── docs/
│   ├── project_overview.md
│   ├── data_dictionary.md
│   └── business_insights.md
│
├── exports/
│   └── dashboard_preview.pdf
│
├── README.md
├── LICENSE
└── .gitignore
```

## Skills Demonstrated

**Data Engineering**
- Data cleaning: duplicate removal, date format standardization, postal code preservation
- Data transformation: calculated columns for profitability and segmentation
- Multi-source integration: combining sales and HR datasets

**Excel Mastery**
- Pivot table design (6 pivot tables for different analyses)
- Dynamic formulas (XLOOKUP replacing VLOOKUP, INDEX/MATCH combinations)
- Named ranges for maintainability
- Slicers for user interactivity

**Analytics & Visualization**
- Business metric design (KPI cards with clear labels)
- Chart selection (appropriate chart types for each insight)
- Color psychology (dark theme for professionalism, accent colors for emphasis)
- Dashboard storytelling (flow from overview → detail)

**Business Acumen**
- Identifying trends (Q4 seasonality)
- Spotting anomalies (customer concentration, regional imbalance)
- Translating data into recommendations
- Understanding business context (retail, HR, margins)

## How to Use

### Opening the File
1. Download `Sales_HR_Analytics.xlsx`
2. Open in Microsoft Excel 2021 or later
3. Enable editing if prompted

### Navigation
- **Active sheet:** Dashboard (main view)
- **Background sheets:** Raw_Sales, Raw_HR, Working_Sales, Working_HR, Analysis
  - Recommend keeping these hidden to avoid accidental changes

### Using Slicers
1. **Order Year:** Filter all sales charts to specific year(s)
   - Single selection: Click year button
   - Multi-select: Hold Ctrl + click years
   - Clear: Click the funnel icon

2. **Region:** Filter sales data by geographic region
   - Shows impact on monthly trends and category performance

3. **Department:** Filter HR charts by department
   - Enables comparison of HR metrics (attrition, income, age distribution)

### Interpreting Charts
- **Sales by Region & Category:** Width = revenue magnitude; color = category
- **Monthly Trend:** Peak labels help identify seasonality; enables year-over-year comparison
- **Attrition Heatmap:** Darker colors = higher attrition risk
- **Income by Tenure:** Shows compensation progression as tenure increases

##  Data Dictionary

| Metric | Definition | Formula | Business Use |
|--------|-----------|---------|---|
| Total Sales | Sum of all order values | SUM(Quantity × Price) | Revenue tracking |
| Profit Margin | Profit as % of sales | SUM(Profit)/SUM(Sales) | Profitability health |
| Attrition Rate | % employees who left | COUNTA(Attrition="Yes")/Total | Talent retention |
| Avg Monthly Income | Mean employee compensation | AVERAGE(MonthlyIncome) | Budget planning |

##  Key Sheet Descriptions

**Raw_Sales Sheet**
- 9,995 rows, 21 columns
- Source: Kaggle Superstore Dataset
- Columns: Invoice, Stock Code, Description, Quantity, InvoiceDate, Price, CustomerID, Country
- Purpose: Immutable source data for audit trail

**Working_Sales Sheet**
- Cleaned data with 5 calculated columns:
  - Profit_Margin: Profit/Sales ratio
  - Sales_Band: Categorization (High/Medium/Low/Very Low)
  - Delivery_Days: Days between order and ship
  - Order_Year: Extracted from invoice date
  - Order_Month: Month name for trend analysis

**Analysis Sheet**
- 6 pivot tables powering the dashboard:
  - Sales by Region & Category
  - Sales Band Performance
  - Monthly Sales Trend
  - Delivery Performance by Ship Mode
  - Attrition by Department & Age
  - Income vs Tenure Analysis

**Dashboard Sheet**
- Interactive visualization layer
- 7 KPI cards linked to formulas
- 4 charts linked to pivot tables
- 3 slicers for user interaction
- Professional formatting with dark theme

##  Improvement Opportunities

- [ ] Add year-over-year growth % to KPI cards
- [ ] Create separate dashboard for customer analysis (RFM)
- [ ] Add forecasting for Q4 sales (moving average trend)
- [ ] Automate monthly data refresh from source system
- [ ] Create department-specific dashboards (Sales, Operations, Finance views)

##  Learning Value

This project demonstrates **portfolio-ready Excel skills:**
-  Real-world data (1M+ transaction values)
-  Complex data cleaning (9,995 rows)
-  Advanced formulas (not basic SUM/AVERAGE)
-  Multi-page architecture (separation of concerns)
-  Interactive elements (slicers, charts)
-  Professional presentation (color, typography, spacing)
-  Business insights (actionable recommendations)

**Hiring managers will see:** Someone who can take raw data → clean it → analyze it → present insights.

## Related Projects

- [E-Commerce RFM Analysis (Python + SQL)](https://github.com/DanFalak7/E-Commerce-RFM-Analysis) — 1M+ transaction customer segmentation
- [HR Analytics Dashboard (Power BI)](https://github.com/DanFalak7/HR-Analytics_PowerBI) — Advanced DAX, heatmaps, violin plots

##  Questions?

This dashboard is production-ready for:
- Sales team: Regional performance tracking
- Finance: Profit margin analysis
- HR: Attrition identification and retention strategy

For technical questions or suggestions, feel free to reach out.

---

**Project Status:**  Complete & Portfolio Ready  
**Last Updated:** June 2026  
**Difficulty Level:** Intermediate-Advanced  
**Time to Build:** 8-10 hours 
