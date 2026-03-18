# 📊 Excel → Python → SQL → Power BI | End-to-End ETL Pipeline

> **Transform raw Excel data into professional business dashboards using a production-grade ETL pipeline.**

---

## 🗺️ Pipeline Architecture

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│    Excel    │────▶│   Python    │────▶│     SQL     │────▶│  Power BI   │
│  Raw Data   │     │  Clean &    │     │  Store &    │     │  Visualize  │
│  (Source)   │     │  Transform  │     │  Warehouse  │     │  & Analyze  │
└─────────────┘     └─────────────┘     └─────────────┘     └─────────────┘
  Operations /        Data Analyst        Data Engineer       Business Analyst
  Sales Team
```

---

## 🚨 Why NOT Direct Excel → Power BI?

| Problem | Impact |
|---|---|
| ❌ Messy data (duplicates, nulls, bad formats) | Incorrect dashboards |
| ❌ Performance issues on large datasets | Slow or crashing reports |
| ❌ Temporary connections | Breaks if file is moved or renamed |
| ❌ No automation | Manual effort every time |

---

## ✅ Professional Solution: ETL Pipeline

| Tool | Role | Responsibility |
|---|---|---|
| 🟢 **Excel** | Raw Data Source | Operations / Sales data entry |
| 🐍 **Python** | Data Cleaning & Transformation | Analyst automates cleaning logic |
| 🛢️ **SQL (MySQL)** | Data Warehouse | Permanent, scalable storage |
| 📊 **Power BI** | Visualization | Business insights & dashboards |

---

## 📁 Dataset Overview

- **Source:** Excel file with **520 rows × 7 columns**
- **Columns:** `Order ID`, `Order Date`, `Customer`, `Region`, `Product`, `Sales`, `Cost`
- **Issues:** Duplicates, missing dates, null values

---

## 🐍 Step 1 — Python: Data Cleaning & Transformation

### Install Dependencies
```bash
pip install pandas openpyxl sqlalchemy pymysql
```

### Load Excel File
```python
import pandas as pd

df = pd.read_excel('sales_data.xlsx')
print(df.shape)  # (520, 7)
```

### Clean the Data
```python
# Remove duplicate rows
df = df.drop_duplicates()  # 520 → 500 rows

# Fix Order Date column — convert to datetime safely
df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce')

# Drop rows with missing Order Date (NaT)
df = df.dropna(subset=['Order Date'])  # 500 → 482 rows

# Fill other missing values
df['Region'].fillna('NA', inplace=True)
df['Product'].fillna('NA', inplace=True)
```

### Feature Engineering
```python
# Add new calculated columns
df['Profit'] = df['Sales'] - df['Cost']
df['Year']   = df['Order Date'].dt.year
df['Month']  = df['Order Date'].dt.month

print(df.shape)  # (482, 10)
```

### Save Cleaned Data
```python
df.to_csv('sales_clean.csv', index=False)
```

> **Why CSV instead of Excel?**
> Avoids Excel formatting quirks and type issues when loading into SQL.

---

## 🛢️ Step 2 — SQL: Store Data in MySQL

### Create Database
```sql
CREATE DATABASE sales_db;
```

### Load Data via Python (SQLAlchemy)
```python
import pandas as pd
from sqlalchemy import create_engine

df = pd.read_csv('sales_clean.csv')

# Create engine — replace @ in password with %40
engine = create_engine('mysql+pymysql://username:password@localhost:3306/sales_db')

# Upload DataFrame to SQL table
df.to_sql(name='sales_data', con=engine, if_exists='replace', index=False)

print("✅ Data uploaded successfully!")
```

### Verify in MySQL
```sql
SHOW TABLES;
SELECT * FROM sales_data LIMIT 5;
SELECT COUNT(*) FROM sales_data;  -- Should return 482
```

> **Why SQL over Excel for Power BI?**
> SQL connections are **permanent** — unaffected by file moves or renames. Excel connections **break** the moment the file path changes.

---

## 📊 Step 3 — Power BI: Build the Dashboard

### Connect Power BI to MySQL
```
Get Data → More → MySQL Database
→ Server: localhost
→ Database: sales_db
→ Authentication: Username / Password
```

### Calculated Column (DAX)
```dax
Profit % = DIVIDE([Profit], [Sales])
```

### Dashboard Components

| Visual | Purpose |
|---|---|
| 📌 **Card** | KPIs — Total Sales, Profit, Profit % |
| 📈 **Line Chart** | Monthly Sales trend by Order Date |
| 📊 **Bar Chart** | Regional Profit comparison |
| 🔲 **Column Chart** | Product-wise Sales breakdown |
| 🎛️ **Slicers** | Filter by Year and Region |

---

## 📦 Project Structure

```
project/
│
├── data/
│   ├── sales_data.xlsx          # Raw Excel source
│   └── sales_clean.csv          # Cleaned output
│
├── scripts/
│   ├── 01_clean_data.py         # Python cleaning script
│   └── 02_load_to_sql.py        # Python → MySQL loader
│
├── powerbi/
│   └── sales_dashboard.pbix     # Power BI report file
│
└── README.md
```

---

## 🔄 Data Cleaning Summary

| Step | Before | After |
|---|---|---|
| Remove duplicates | 520 rows | 500 rows |
| Remove missing dates | 500 rows | 482 rows |
| Add Profit column | 7 columns | 10 columns |
| Add Year & Month | — | Extracted from Order Date |

---

## 🛠️ Tech Stack

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![MySQL](https://img.shields.io/badge/MySQL-4479A1?style=for-the-badge&logo=mysql&logoColor=white)
![Power BI](https://img.shields.io/badge/Power_BI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)

---

## 🎯 Key Learnings

- ✅ **Never use Excel directly as a Power BI source** in production — connections are fragile
- ✅ **Python automates** data cleaning at scale — no manual effort
- ✅ **SQL provides permanent, reliable** data connections for Power BI
- ✅ **ETL thinking** separates raw data from analysis-ready data
- ✅ This pipeline covers **4 real-world roles**: Excel Operator → Data Analyst → Data Engineer → Business Analyst

---

## 🔮 Coming Next

- [ ] **Incremental data updates** — append new data without rebuilding the pipeline
- [ ] **Automated scheduling** — refresh data on a schedule
- [ ] **Error logging** — track pipeline failures gracefully

---

## 🙌 Contributing

Pull requests welcome! For major changes, please open an issue first to discuss what you'd like to change.

---

## 📄 License

MIT License — feel free to use this pipeline as a template for your own projects.
