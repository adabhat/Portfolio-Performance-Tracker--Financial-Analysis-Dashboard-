
# 📈 Portfolio Performance Tracker Dashboard

An interactive Excel-based portfolio tracking dashboard designed to analyze stock holdings, monitor portfolio performance, and visualize historical trends through dynamic charts and financial metrics.

Built using advanced Excel functions including **XLOOKUP, SUMIFS, INDEX, IFERROR, and dynamic arrays**, the dashboard provides a centralized view of investment performance and stock comparisons.

---

## 🚀 Features

### Portfolio Overview

* Track current stock holdings
* Monitor portfolio market value
* View sector/industry classification
* Analyze portfolio allocation across investments

### Performance Analysis

* Gain/Loss tracking for individual holdings
* Daily portfolio change analysis
* Market value calculations
* Allocation (%) breakdown

### Stock Metrics

* Current Price
* Previous Close
* Day Change
* 52-Week High
* 52-Week Low

### Historical Analysis

* Dynamic stock selection
* Historical performance tracking for:

  * 1 Month
  * 3 Months
  * 6 Months
  * 12 Months
* Interactive trend visualization

### Comparative Analysis

* Compare two stocks over a selected time period
* Dynamic comparison charts
* Secondary-axis visualization for stocks with different price ranges

---

# 📊 Dashboard Components

### Holdings Summary

Displays:

* Stock Ticker
* Industry
* Quantity Held
* Current Price
* Market Value
* Gain/Loss
* Allocation Percentage

### Portfolio Allocation

Calculates each stock's contribution to total portfolio value, helping identify concentration risk and portfolio diversification.

### Historical Trend Analysis

Allows users to:

1. Select a stock
2. Select a time period (1, 3, 6, or 12 months)
3. View historical price movement through dynamic charts

### Stock Comparison

Allows users to compare two stocks over the same selected time period to evaluate relative performance.

---

# 🛠️ Excel Functions Used

## XLOOKUP

Used to dynamically retrieve:

* Industry information
* Historical stock data
* Comparison stock data

Example:

=XLOOKUP(
C4,
Historical_Data[[#Headers],[RELIANCE]:[ICICIBANK]],
Historical_Data[[RELIANCE]:[ICICIBANK]]
)

---

## SUMIFS

Used for:

* Quantity calculations
* Investment cost calculations
* Portfolio aggregation

Example:

=SUMIFS(...)

---

## INDEX

Used to dynamically retrieve the last N months of historical data based on user selection.

Example:

=INDEX($I$10#,ROWS($I$10#)-$C$5+K2)

---

## ROWS

Used to determine the total number of historical records available.

Example:

=ROWS($I$10#)

---

## IFERROR

Used to suppress reference errors and return blank cells for cleaner chart generation.

Example:

=IFERROR(...,"")

---

# 📈 Dynamic Chart Logic

Excel 2021 does not support the `STOCKHISTORY()` and `TAKE()` functions used in many modern portfolio dashboards.

To overcome this limitation:

1. Historical price data was stored in a dedicated `Historical_Data` sheet.
2. `XLOOKUP` was used to retrieve the selected stock's complete price history.
3. Helper columns and `INDEX` formulas were used to dynamically extract the last 1, 3, 6, or 12 months of data.
4. Dynamic ranges were used as chart sources to automatically update visualizations when user selections change.

---

# 📂 Project Structure


├── Transactions
│   ├── Buy/Sell Records
│   ├── Previous Close Prices
│
├── Dashboard
│   ├── Portfolio Summary
│   ├── Performance Metrics
│   ├── Allocation Analysis
│   ├── Historical Trend Chart
│   └── Stock Comparison Chart
│
└── Historical_Data
    ├── Monthly Stock Prices
    ├── Date Series
    └── Lookup Data
```

---

# 🎯 Key Learning Outcomes

* Portfolio performance analysis
* Financial dashboard development in Excel
* Dynamic chart creation
* Stock benchmarking and comparison
* Portfolio allocation analysis
* Advanced lookup and aggregation techniques
* Building interactive dashboards without VBA

---

# 🧰 Tools Used

* Microsoft Excel 2021
* Financial Analysis Concepts
* Dynamic Charts
* Data Validation Dropdowns
* Excel Formulas & Functions
This version is recruiter-friendly, technically accurate, and aligned with what you actually built during the project. It also explains the Excel 2021 workaround, which is arguably the most interesting part of the project.
