# 🪔 Diwali Sales Analysis – EDA & Dashboard in Excel

An end-to-end **Exploratory Data Analysis in Microsoft Excel** of 11,239 Diwali-season sales transactions from India: data audit, KPI formulas, PivotTables, charts and an interactive slicer-driven dashboard, with insights on **who buys, what sells and where**.

## 📌 Project Overview
| | |
|---|---|
| **Goal** | Identify the customer segments, products and regions that drive Diwali revenue |
| **Data** | `diwali_data.xlsx` – 11,239 rows × 13 columns, no blank cells |
| **Tool** | Microsoft Excel 365 (Tables, formulas, PivotTables, PivotCharts, Slicers) |
| **Deliverables** | Interactive dashboard, report (.docx), presentation (.pptx) |

## 🗂️ Repository Structure
```
├── data/
│   ├── diwali_data.xlsx                # raw data
│   └── diwali_data_dashboard.xlsx      # Excel project (Analysis, Analysis_2, Dashboard)
├── images/                             # charts + dashboard screenshot
├── docs/
│   ├── Diwali_EDA_Report.docx
│   └── Diwali_EDA_Presentation.pptx
└── README.md
```

## 📖 Data Dictionary
| Column | Description |
|---|---|
| User_ID, Cust_name | Customer identifier and name |
| Product_ID | Product code (2,350 unique) |
| Gender | F / M |
| Age Group, Age | 7 age bands (0-17 … 55+) and exact age (12–92) |
| Marital_Status | 0 = unmarried, 1 = married |
| State, Zone | 16 states, 5 zones (Central, Southern, Western, Northern, Eastern) |
| Occupation | 15 occupations |
| Product_Category | 18 categories |
| Orders | Units ordered in the transaction (1–4) |
| Amount | Transaction value (₹) |

## 🛠️ Excel Workflow
1. **Import & structure** – raw range converted to an Excel Table named `Diwali_Data`.
2. **Audit** – checked blanks, duplicates, text spacing, Age vs Age Group, outliers.
3. **KPI cells (`Analysis` sheet)**
   - Total Revenue `=SUM(Diwali_Data[Amount])`
   - Total Orders `=SUM(Diwali_Data[Orders])`
   - Total Customers `=COUNTA(UNIQUE(Diwali_Data[User_ID]))`
   - Average Order Value `=SUM(Diwali_Data[Amount])/SUM(Diwali_Data[Orders])`
4. **PivotTables (`Analysis_2` sheet)** – 6 pivots: revenue by Gender, Age Group, State, Occupation, Product Category; orders by Product Category.
5. **Charts** – doughnut (gender), pie (age group), bar/column charts (state, occupation, category).
6. **Dashboard** – charts linked to pivots with 5 slicers (Gender, Age Group, State, Occupation, Product Category).

## 🔑 Key Findings
| KPI | Value |
|---|---|
| Total revenue | ₹106.25M |
| Units ordered | 27,981 |
| Transactions | 11,239 |
| Customers (unique User_ID) | 3,752 |
| Revenue per unit ordered | ₹3,797 |
| Average / median transaction | ₹9,454 / ₹8,109 |

1. **Women drive sales:** 70% of revenue (₹74.3M vs ₹31.9M); spend per transaction is nearly identical for both genders.
2. **Ages 26–35 are the top group** (40.1%); ages 18–45 together give 77%. Unmarried buyers give 58.5%.
3. **Food is #1 category** (31.9%); Food, Clothing, Electronics and Footwear together = 77% of revenue.
4. **Top states:** Uttar Pradesh (18.2%), Maharashtra (13.6%), Karnataka (12.7%), Delhi (10.9%). Central zone = 39.2%.
5. **Top occupations:** IT Sector (13.9%), Healthcare (12.3%), Aviation (11.9%), Banking (10.1%).
6. Order quantity, age and marital status have almost no correlation with transaction amount (|r| < 0.04).

![Gender](images/01_gender.png)
![Category](images/06_category.png)

## ⚠️ Data Quality Notes
| Issue | Count | Suggested fix |
|---|---|---|
| Exact duplicate rows | 8 (₹70K) | Data ▸ Remove Duplicates |
| Age Group "36-45" but Age 26–35 | 22 rows | Recompute group from Age |
| User_IDs with two genders | 502 | Customer count is approximate; analyse at transaction level |
| `Andhra Pradesh` contains a non-breaking space | 811 rows | `=SUBSTITUTE(x,CHAR(160)," ")` |
| Dashboard cells A2, AK11 show `#VALUE!` | 2 | Repair the formulas |
| Slicer filters saved as active | – | Clear all slicers before saving so pivots show full data |

## 💡 Recommendations
1. Target women aged 26–35 with personalised festive campaigns.
2. Stock and promote Food, Apparel, Electronics and Footwear early.
3. Test gifting offers to grow male and 46+ segments.
4. Focus on top states; expand in Punjab, Rajasthan and Telangana.
5. Run corporate offers for IT, Healthcare, Aviation and Banking professionals.

## ▶️ How to Use
1. Open `data/diwali_data_dashboard.xlsx` in **Excel 365** (`UNIQUE()` needs 365/2021).
2. Go to the **Dashboard** sheet and use the slicers.
3. Right-click any pivot ▸ *Refresh* after changing the data.

## ⚠️ Limitations
- Single festive season; no dates, so no time-trend analysis.
- No cost or profit data, so results are revenue-only.
- Customer identity is unreliable (see data quality notes).

## 🚀 Future Work
Add date/time data, RFM segmentation, profit analysis, and a Power BI or Python version of this analysis.

## 👤 Author
Vaibhav Choudhary



