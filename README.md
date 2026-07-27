# 📦 Amazon India — Sales & Product Analysis

A comprehensive end-to-end data analysis of Amazon India product listings, covering data cleaning, descriptive statistics, correlation analysis, regression modelling, hypothesis testing, and data visualisation — all performed in Microsoft Excel.

---

## 🎯 Project Overview

**Main Aim:**
To analyse Amazon India product listings and determine what factors — price, discount, and product category — influence customer satisfaction (ratings), using the full data analysis pipeline from raw data to actionable insights.

**Business Questions Answered:**
- What does Amazon India's product and pricing landscape look like?
- Do price and discount percentages affect how customers rate products?
- Which product categories perform best in terms of price, discount, and satisfaction?
- Are heavily discounted products rated significantly differently from less discounted ones?

---

## 📁 Repository Structure

```
Amazon-Indian-Sales---Analysis/
│
├── amazon.xlsx                        # Main Excel workbook (all analysis inside)
│   ├── amazon                         # Original raw dataset (untouched)
│   ├── Working                        # Cleaned working dataset
│   ├── Unique Products                # Deduplicated product sheet + all analysis
│   ├── Pivot Tables                   # Pivot tables and charts
│   └── Dashboard                      # Visual dashboard
│
├── a1.png                             # Chart screenshot 1
├── a2.png                             # Chart screenshot 2
└── README.md                          # Project documentation
```

---

## 📊 Dataset Information

| Attribute | Detail |
|---|---|
| **Dataset Name** | Amazon India Product Reviews & Pricing |
| **Source** | [Kaggle — karkavelrajaj](https://www.kaggle.com/datasets/karkavelrajaj/amazon-sales-dataset) |
| **Collection Method** | Web scraping from amazon.in |
| **Time Period** | 2022 |
| **Total Records** | 1,465 rows |
| **Unique Products** | 1,351 |
| **Total Columns (Original)** | 16 |
| **Columns Used for Analysis** | 9 cleaned columns |

### Key Variables

| Column | Type | Description |
|---|---|---|
| `product_id` | Text | Unique product identifier |
| `main_category` | Text | Primary product category (extracted) |
| `clean_discounted_price` | Number | Selling price after discount (Rs.) |
| `clean_actual_price` | Number | Original price before discount (Rs.) |
| `clean_discount_pct` | Number | Percentage discount applied |
| `clean_rating` | Number | Average customer rating (out of 5) |
| `clean_rating_count` | Number | Total number of customer ratings |

### Product Categories

| Category | Products |
|---|---|
| Electronics | 490 |
| Home & Kitchen | 448 |
| Computers & Accessories | 375 |
| Office Products | 31 |
| Others (Car, Musical, Toys, Health) | 7 |
| **Total** | **1,351** |

---

## 🔧 Tools Used

| Tool | Purpose |
|---|---|
| **Microsoft Excel Online** | Data cleaning, analysis, pivot tables, charts |
| **Excel Formulas** | SUBSTITUTE, VALUE, TRIM, CLEAN, IFERROR, LEFT, FIND |
| **Excel Functions** | AVERAGE, MEDIAN, MODE, STDEV, VAR, MIN, MAX, QUARTILE, CORREL, SLOPE, INTERCEPT, RSQ, T.TEST, COUNTIF, AVERAGEIF |
| **Excel Features** | Pivot Tables, Conditional Formatting, Remove Duplicates, Charts |

---

## 🧹 Data Cleaning Steps

The following issues were identified and resolved before analysis:

1. **Currency symbol in prices** — Removed ₹ symbol using `=VALUE(SUBSTITUTE(SUBSTITUTE(D2,"₹",""),",",""))`
2. **Percentage symbol in discounts** — Removed % using `=VALUE(SUBSTITUTE(F2,"%",""))`
3. **Comma-formatted numbers** — Removed commas from rating counts using `=VALUE(SUBSTITUTE(H2,",",""))`
4. **Scraping artefacts in ratings** — Fixed pipe characters using `=IFERROR(VALUE(TRIM(CLEAN(G2))),"")`
5. **Nested category paths** — Extracted main category using `=LEFT(C2,FIND("|",C2)-1)`
6. **Duplicate product IDs** — Removed 114 duplicates using Data → Remove Duplicates
7. **Missing values** — 2 missing values in rating_count handled automatically by VALUE() formula
8. **Irrelevant columns** — Removed 8 columns (URLs, user details, review text, descriptions)

---

## 📈 Analysis Performed

### 1. Descriptive Statistics

Computed across discounted price, actual price, discount percentage, and customer rating:

| Metric | Disc. Price (Rs.) | Discount % | Rating |
|---|---|---|---|
| Mean | 3,304.80 | 46.69% | 4.09 |
| Median | 899 | 49% | 4.10 |
| Mode | 299 | 50% | 4.10 |
| Std Deviation | 7,173.97 | 21.63 | 0.297 |
| Min | 39 | 0% | 2.0 |
| Max | 77,990 | 94% | 5.0 |
| Q1 | 349 | 31% | 3.9 |
| Q3 | 2,174 | 62% | 4.3 |
| IQR | 1,825 | 31 | 0.4 |

### 2. Outlier Detection (IQR Method)

| Metric | Value |
|---|---|
| Lower Bound | Rs. -2,388.50 |
| Upper Bound | Rs. 4,911.50 |
| Outlier Count | 209 products (15.5%) |

### 3. Category-Level Analysis (Pivot Table)

| Category | Avg Price (Rs.) | Avg Discount% | Avg Rating |
|---|---|---|---|
| Electronics | 6,225.87 | 49.9% | 4.08 |
| Computers & Accessories | 947.49 | 53.2% | 4.15 |
| Home & Kitchen | 2,330.62 | 40.1% | 4.04 |
| Office Products | 301.58 | 12.4% | 4.31 |

### 4. Correlation Analysis

| Variable Pair | r Value | Strength |
|---|---|---|
| Discount% vs Rating Count | 0.004 | No correlation |
| Price vs Rating | 0.127 | Little correlation |
| Discount% vs Rating | -0.162 | Little correlation |
| Actual Price vs Discount% | -0.112 | Little correlation |
| Price vs Rating Count | -0.024 | No correlation |

### 5. Regression Analysis

**Predicting Customer Rating from Discount Percentage:**

```
Rating = 4.196 + (-0.00223 × Discount%)
```

| Metric | Value |
|---|---|
| Slope (b) | -0.00223 |
| Intercept (a) | 4.196 |
| R-squared | 0.026 (2.6%) |

### 6. Hypothesis Testing (T-Test)

> **Question:** Do products with discounts above 50% receive significantly different ratings from products with discounts of 50% or below?

| Component | Detail |
|---|---|
| H₀ | High discount rating = Low discount rating |
| H₁ | High discount rating ≠ Low discount rating |
| Alpha | 0.05 |
| High Discount Mean (>50%) | 4.052 (608 products) |
| Low Discount Mean (≤50%) | 4.125 (743 products) |
| p-value | 0.0000111 |
| **Decision** | **Reject H₀ — Statistically significant difference** |

---

## 📉 Visualisations

All charts were built in Excel using Pivot Tables and the Chart Wizard:

| Chart | Type | Insight |
|---|---|---|
| Products by Category | Column Chart | Electronics & Home & Kitchen dominate |
| Average Rating by Category | Column Chart | Office Products highest rated (4.31) |
| Average Price by Category | Column Chart | Electronics far more expensive than others |
| Discount Distribution | Histogram | Near-normal, peaking at 47%–54% |
| Price vs Rating | Scatter Plot | Weak positive relationship confirmed |

---

## 🔑 Key Findings

1. **Price distribution is right-skewed** — Mean (Rs.3,304) far exceeds Median (Rs.899) due to 209 premium outlier products
2. **Amazon India relies heavily on discounting** — Average discount is 47%; most common is 50%
3. **Electronics dominates by price** (Rs.6,226 avg) but Office Products leads in satisfaction (4.31 rating)
4. **Price and discount do NOT strongly predict ratings** — All correlations below 0.3; R² = 2.6%
5. **High discounts are linked to lower ratings** — Statistically significant at p = 0.0000111
6. **Customer satisfaction is consistently high** — Overall average rating of 4.09 / 5.0

---

## 💡 Recommendations

- Target the **Rs.299–Rs.899 price range** where demand is highest
- **Avoid discounts above 50%** — associated with marginally lower customer ratings
- **Office Products** demonstrate that niche, quality products succeed without heavy discounting
- Invest in **product quality and accurate descriptions** over price competition
- Expand analysis with **time-series data** to track price and rating trends over time

---

## 🗂️ How to Use This Repository

1. Download `amazon.xlsx`
2. Open in Microsoft Excel (desktop or online)
3. Navigate through the sheet tabs:
   - Start with **amazon** to see the raw data
   - Go to **Working** to see the cleaned data and all added columns
   - Go to **Unique Products** to see the statistics, correlation, regression, and hypothesis testing tables
   - Go to **Pivot Tables** to see category analysis and charts
   - Go to **Dashboard** for the visual overview
4. All formulas are live — you can explore and interact with the data directly

---

## 📄 Report

A full written report following the standard data analysis format (Introduction → Data → Insights → Conclusion) is available as a Word document in this repository.

---

## 👤 Author

**Kashb-shielah**
GitHub: [https://github.com/Kashb-shielah](https://github.com/Kashb-shielah)

---

## 📌 Acknowledgements

- Dataset: [Amazon India Product Dataset on Kaggle](https://www.kaggle.com/datasets/karkavelrajaj/amazon-sales-dataset)
- Analysis framework based on a structured Data Analysis course curriculum covering:
  - Introduction to Data Analysis
  - Data Collection
  - Data Challenges
  - Data Cleaning
  - Data Analysis
  - Data Visualisation

---

*This project is part of an ongoing data analysis portfolio. A Healthcare Dataset Analysis is coming next.*
