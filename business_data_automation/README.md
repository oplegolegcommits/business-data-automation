# Business Data Automation

A small Python automation tool that transforms monthly e-commerce sales CSV files into a ready-to-use Excel business report.

The goal of the project is simple:

> **Drop CSV files into a folder → run one command → get an automated business report.**

## What it does

The script automatically:

- loads all monthly CSV files from `data/`;
- cleans and validates the data;
- removes duplicates and invalid records;
- recalculates revenue from `price × quantity`;
- aggregates sales data;
- calculates key business metrics;
- generates an Excel report;
- creates charts for management-level analysis.

## Business Metrics

The generated report includes:

- Total Revenue
- Total Orders
- Units Sold
- Average Order Value (AOV)
- Monthly Revenue
- Month-to-Month Revenue Growth
- Top Products
- Revenue by Category
- Revenue by Region
- Regional AOV
- Revenue Share

## Project Structure

```text
business-data-automation/
│
├── data/
│   ├── january.csv
│   ├── february.csv
│   ├── march.csv
│   └── april.csv
│
├── output/
│   └── report.xlsx
│
├── generate_report.py
├── requirements.txt
└── README.md
```

## Input Data

Each CSV file contains sales transactions with the following fields:

```text
date
product
category
price
quantity
revenue
customer
region
```

Example:

```csv
date,product,category,price,quantity,revenue,customer,region
2026-01-03,Wireless Mouse,Accessories,24.90,2,49.80,C0042,West
2026-01-03,Mechanical Keyboard,Accessories,79.90,1,79.90,C0178,North
2026-01-04,USB-C Hub,Accessories,39.90,3,119.70,C0091,Central
```

The original `revenue` column is not blindly trusted. The script recalculates it from:

```text
revenue = price × quantity
```

This makes the pipeline more robust against incorrect source data.

## Data Cleaning

The project demonstrates several basic data-quality checks:

- conversion of dates to proper datetime format;
- conversion of numeric fields;
- removal of invalid records;
- removal of duplicate transactions;
- handling of missing values;
- validation of positive prices and quantities;
- recalculation of revenue.

The sample dataset intentionally contains several data-quality issues to demonstrate this process.

## Excel Report

Running the script creates:

```text
output/report.xlsx
```

The workbook contains:

### Dashboard

A management-level overview with:

- Revenue
- Orders
- Units Sold
- Average Order Value
- Monthly performance
- Revenue trend
- Product and regional charts

### Products

Product-level performance:

- Revenue
- Units Sold
- Orders
- Revenue Share

### Regions

Regional performance:

- Revenue
- Orders
- Units Sold
- Average Order Value
- Revenue Share

### Categories

Category-level performance:

- Revenue
- Orders
- Units Sold
- Revenue Share

### Data

The cleaned dataset used for the analysis.

## Visualization

The report automatically generates four charts:

1. Revenue by Month
2. Top 10 Products by Revenue
3. Revenue by Region
4. Revenue by Category

## Technologies

- **Python**
- **Pandas** — data processing and aggregation
- **OpenPyXL** — Excel report generation and formatting

## Installation

Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Place the monthly CSV files into:

```text
data/
```

Then run:

```bash
python generate_report.py
```

The report will be generated automatically:

```text
output/report.xlsx
```

Example console output:

```text
=======================================================
BUSINESS DATA AUTOMATION
=======================================================
Found 4 CSV files.
  Loading april.csv...
  Loading february.csv...
  Loading january.csv...
  Loading march.csv...

Cleaning data...
  Removed 17 invalid/duplicate rows.
  Final rows: 1383

Calculating business metrics...
Creating Excel report...

=======================================================
REPORT GENERATED SUCCESSFULLY
=======================================================
File: output/report.xlsx

Revenue:       $XXX,XXX.XX
Orders:        1,XXX
Units sold:    X,XXX
Average order: $XX.XX
=======================================================
```

## Business Use Case

This type of automation can be used by a small or medium-sized business that regularly receives sales data in CSV/Excel files and needs recurring performance reports.

Instead of manually:

```text
Open CSV
    ↓
Clean data
    ↓
Calculate metrics
    ↓
Build pivot tables
    ↓
Create charts
    ↓
Prepare Excel report
```

the entire workflow becomes:

```text
CSV files
    ↓
python generate_report.py
    ↓
report.xlsx
```

This project demonstrates how repetitive reporting tasks can be converted into a simple, repeatable Python workflow.