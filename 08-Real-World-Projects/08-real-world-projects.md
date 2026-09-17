# Module 8: Real-World Excel Projects

> **Level:** Beginner → Advanced (applied)
> **Duration:** 40–60 hours (8 projects × 5–7 hours each)
> **Prerequisites:** Modules 1–7 (all core Excel skills)

---

## Table of Contents

1. [Project 1: Personal Budget Tracker](#project-1-personal-budget-tracker)
2. [Project 2: Sales Dashboard](#project-2-sales-dashboard)
3. [Project 3: Inventory Management System](#project-3-inventory-management-system)
4. [Project 4: Employee Schedule & Leave Tracker](#project-4-employee-schedule--leave-tracker)
5. [Project 5: Invoice Generator Template](#project-5-invoice-generator-template)
6. [Project 6: Data Cleaning Pipeline](#project-6-data-cleaning-pipeline)
7. [Project 7: Financial Model](#project-7-financial-model)
8. [Project 8: Automated Report Generator](#project-8-automated-report-generator)
9. [Industry Use Cases](#industry-use-cases)
10. [Excel Certification Paths](#excel-certification-paths)
11. [Excel vs Google Sheets](#excel-vs-google-sheets)
12. [Recommended Templates & Add-ins](#recommended-templates--add-ins)
13. [YouTube Video References](#youtube-video-references)
14. [Sources](#sources)

---

# Project 1: Personal Budget Tracker

## Overview

Build a comprehensive personal finance tracker that records income and expenses by category, calculates running balances, and presents spending insights through a visual dashboard. This is the ideal first project because it uses foundational skills in a context everyone understands.

**Difficulty:** ⭐⭐ Beginner
**Time:** 5–6 hours
**Skills Used:** SUM, SUMIF, SUMIFS, charts (pie, bar, line), conditional formatting, data validation, named ranges, basic formatting

## Step-by-Step Instructions

### Step 1 — Set Up the Workbook Structure

Create four sheets:

| Sheet Name | Purpose |
|---|---|
| `Transactions` | Raw data entry for every income/expense |
| `Categories` | Reference list of income & expense categories |
| `Monthly Summary` | Aggregated totals by month and category |
| `Dashboard` | Charts and KPIs |

### Step 2 — Build the Categories Sheet

In the `Categories` sheet, create two tables:

**Income Categories (Column A):**

| Category |
|---|
| Salary |
| Freelance |
| Investments |
| Gifts |
| Other Income |

**Expense Categories (Column C):**

| Category |
|---|
| Housing |
| Utilities |
| Groceries |
| Transportation |
| Dining Out |
| Entertainment |
| Healthcare |
| Insurance |
| Clothing |
| Education |
| Savings |
| Debt Payments |
| Miscellaneous |

**Named Ranges:**
- Select Income Categories → Formulas → Define Name → `IncomeCategories`
- Select Expense Categories → Define Name → `ExpenseCategories`

### Step 3 — Build the Transactions Sheet

Set up headers in Row 1:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| **Date** | **Type** | **Category** | **Description** | **Amount** | **Account** | **Running Balance** |

**Data Validation (dropdowns):**

1. **Type column (B):** Select B2:B5000 → Data → Data Validation → List → Source: `Income,Expense`
2. **Category column (C):** Use a dependent dropdown:
   ```excel
   =IF(B2="Income", IncomeCategories, ExpenseCategories)
   ```
   (Data Validation → List → Source: the formula above)
3. **Account column (F):** Create a named range `Accounts` with: `Checking,Savings,Credit Card,Cash`

**Running Balance Formula (G2):**
```excel
=IF(B2="Income", G1+E2, G1-E2)
```
Drag down. Set G1 to your starting balance.

### Step 4 — Build the Monthly Summary Sheet

Create a summary table:

| | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | **Month** | **Total Income** | **Total Expenses** | **Net Savings** | **Savings Rate** |
| 2 | Jan 2025 | | | | |

**Formulas (assuming Transactions data in rows 2–5000):**

```excel
B2 (Total Income):
=SUMIFS(Transactions!E:E, Transactions!B:B, "Income",
        Transactions!A:A, ">="&DATE(2025,1,1),
        Transactions!A:A, "<"&DATE(2025,2,1))

C2 (Total Expenses):
=SUMIFS(Transactions!E:E, Transactions!B:B, "Expense",
        Transactions!A:A, ">="&DATE(2025,1,1),
        Transactions!A:A, "<"&DATE(2025,2,1))

D2 (Net Savings): =B2-C2

E2 (Savings Rate): =IF(B2>0, D2/B2, 0)  → Format as percentage
```

For category breakdowns, add a second table:

| Category | Budget | Actual | Difference | % of Budget |
|---|---|---|---|---|

```excel
Actual (for "Groceries" in cell A10):
=SUMIFS(Transactions!E:E, Transactions!C:C, A10,
        Transactions!B:B, "Expense",
        Transactions!A:A, ">="&DATE(2025,1,1),
        Transactions!A:A, "<"&DATE(2025,2,1))

Difference: =B10-C10
% of Budget: =IF(B10>0, C10/B10, 0)
```

### Step 5 — Apply Conditional Formatting

**Budget vs Actual (Difference column):**

1. Select the Difference column
2. Home → Conditional Formatting → New Rule → Format cells that contain
3. Cell Value < 0 → Red fill (over budget)
4. Cell Value > 0 → Green fill (under budget)
5. Cell Value = 0 → Yellow fill (on target)

**Spending Alerts:**
Select the % of Budget column → Conditional Formatting → Color Scale:
- Green (0%) → Yellow (80%) → Red (100%+)

### Step 6 — Build the Dashboard

**Chart 1 — Monthly Income vs Expenses (Clustered Bar)**
1. Select Month, Total Income, Total Expenses columns from Monthly Summary
2. Insert → Chart → Clustered Bar
3. Title: "Monthly Income vs Expenses"

**Chart 2 — Expense Breakdown (Pie/Donut)**
1. Select Category and Actual from the category breakdown table
2. Insert → Chart → Doughnut
3. Title: "Spending by Category"
4. Add data labels with percentages

**Chart 3 — Savings Trend (Line)**
1. Select Month and Net Savings
2. Insert → Chart → Line with Markers
3. Title: "Monthly Savings Trend"
4. Add a trendline (Linear)

**Chart 4 — Budget Gauge (for single month)**
Use a stacked bar showing spent vs remaining:
```
=BudgetUsed  |  =BudgetRemaining
```

**KPI Cards (top of dashboard):**
Use merged cells with large font:
```excel
Total Income:    =SUMIFS(Transactions!E:E, Transactions!B:B, "Income")
Total Expenses:  =SUMIFS(Transactions!E:E, Transactions!B:B, "Expense")
Net Savings:     =TotalIncome-TotalExpenses
Savings Rate:    =NetSavings/TotalIncome
```

## Sample Data

| Date | Type | Category | Description | Amount | Account |
|---|---|---|---|---|---|
| 2025-01-01 | Income | Salary | January paycheck | 4500.00 | Checking |
| 2025-01-03 | Expense | Housing | Rent payment | 1200.00 | Checking |
| 2025-01-05 | Expense | Groceries | Weekly groceries | 85.50 | Credit Card |
| 2025-01-06 | Expense | Transportation | Gas | 45.00 | Credit Card |
| 2025-01-08 | Expense | Dining Out | Restaurant dinner | 62.30 | Credit Card |
| 2025-01-10 | Income | Freelance | Web design project | 750.00 | Savings |
| 2025-01-12 | Expense | Utilities | Electric bill | 95.00 | Checking |
| 2025-01-15 | Expense | Insurance | Car insurance | 150.00 | Checking |
| 2025-01-18 | Expense | Entertainment | Streaming services | 25.99 | Credit Card |
| 2025-01-20 | Expense | Healthcare | Doctor copay | 30.00 | Cash |
| 2025-01-22 | Expense | Groceries | Weekly groceries | 92.15 | Credit Card |
| 2025-01-25 | Income | Investments | Dividend payment | 125.00 | Savings |
| 2025-01-28 | Expense | Clothing | Winter jacket | 89.99 | Credit Card |
| 2025-01-30 | Expense | Savings | Transfer to savings | 500.00 | Checking |

## Expected Output

- A `Transactions` sheet with 50+ entries over 3+ months
- A `Monthly Summary` showing income, expenses, net savings per month with % savings rate
- Category breakdown with budget vs actual comparison
- A `Dashboard` sheet with 4 charts and KPI cards
- Conditional formatting highlighting over-budget categories in red
- Dropdown menus for Type, Category, and Account

## Tips & Tricks

- **Use Excel Tables** (Ctrl+T) for the Transactions range — formulas auto-expand
- **Freeze Row 1** on the Transactions sheet for scrolling
- Use `TEXT(A2, "MMM YYYY")` to group transactions by month in pivot-style summaries
- Add a **"Quick Add"** section at the top of Transactions for rapid data entry
- Use `SUBTOTAL(9, range)` instead of `SUM` if you filter transactions

## Common Pitfalls

- ❌ **Forgetting to lock the date range** in SUMIFS (use absolute references for the summary sheet)
- ❌ **Circular references** in running balance if row 1 formula references itself
- ❌ **Not using consistent category names** — always use the dropdown, never type manually
- ❌ **Including the header row** in SUMIFS ranges
- ❌ **Date format mismatches** — ensure all dates are actual Excel dates, not text

---

# Project 2: Sales Dashboard

## Overview

Build an interactive sales analytics dashboard using PivotTables, PivotCharts, slicers, and timelines. Analyze sales data by region, product, salesperson, and time period. Create KPI cards and dynamic charts that respond to user selections.

**Difficulty:** ⭐⭐⭐ Intermediate
**Time:** 6–7 hours
**Skills Used:** PivotTables, PivotCharts, slicers, timeline, GETPIVOTDATA, INDEX/MATCH, conditional formatting, dashboard layout, sparklines

## Step-by-Step Instructions

### Step 1 — Prepare the Sales Data

Create a `SalesData` sheet with the following columns:

| Column | Header | Format |
|---|---|---|
| A | OrderID | Text (e.g., ORD-0001) |
| B | OrderDate | Date |
| C | Region | Text |
| D | Salesperson | Text |
| E | Product | Text |
| F | Category | Text |
| G | Quantity | Number |
| H | UnitPrice | Currency |
| I | Revenue | Currency |
| J | Cost | Currency |
| K | Profit | Currency |
| L | ProfitMargin | Percentage |

**Key Formulas:**
```excel
I2: =G2*H2          (Revenue)
K2: =I2-J2          (Profit)
L2: =IF(I2>0, K2/I2, 0)  (Profit Margin)
```

Convert to Excel Table (Ctrl+T) and name it `SalesTable`.

### Step 2 — Create the PivotTable

1. Insert → PivotTable → From Table/Range
2. Place on a new sheet called `PivotAnalysis`
3. **Configure the PivotTable:**
   - **Rows:** Region, then Salesperson (nested)
   - **Columns:** Product Category
   - **Values:** Sum of Revenue, Sum of Profit, Count of OrderID

4. **Right-click the PivotTable** → PivotTable Options:
   - Check "Show items with no data"
   - Set "For empty cells show: 0"

### Step 3 — Add Calculated Fields

In the PivotTable:
1. PivotTable Analyze → Fields, Items & Sets → Calculated Field
2. Add `AvgOrderValue` = `Revenue / Quantity`
3. Add `ProfitMargin` = `Profit / Revenue`

### Step 4 — Create PivotCharts

**Chart 1 — Revenue by Region (Clustered Bar)**
1. Click inside PivotTable → Insert → PivotChart → Clustered Bar
2. Move chart to `Dashboard` sheet

**Chart 2 — Monthly Revenue Trend (Line)**
1. Create a second PivotTable with OrderDate grouped by Month
2. Values: Sum of Revenue, Sum of Profit
3. Insert Line Chart with markers

**Chart 3 — Product Category Mix (Pie)**
1. Create a PivotChart → Pie Chart
2. Show percentage data labels

**Chart 4 — Salesperson Performance (Horizontal Bar)**
1. Rows: Salesperson, Values: Sum of Revenue
2. Sort descending by revenue

### Step 5 — Add Slicers and Timeline

**Slicers:**
1. Click PivotTable → PivotTable Analyze → Insert Slicer
2. Add slicers for: Region, Category, Salesperson
3. **Connect to all PivotTables:** Right-click each slicer → Report Connections → check all PivotTables
4. Arrange slicers at the top of the Dashboard

**Timeline:**
1. PivotTable Analyze → Insert Timeline → OrderDate
2. Set to show by Quarter (dropdown in timeline)
3. Connect to all PivotTables

### Step 6 — Build KPI Cards

Above the charts, create KPI cards using GETPIVOTDATA:

```excel
Total Revenue:
=GETPIVOTDATA("Revenue", PivotAnalysis!$A$3)

Total Orders:
=GETPIVOTDATA("OrderID", PivotAnalysis!$A$3, "OrderID", "")

Avg Order Value:
=TotalRevenue/TotalOrders

Top Region:
=INDEX(RegionList, MATCH(MAX(RegionRevenue), RegionRevenue, 0))
```

Format KPIs with:
- Large bold numbers
- Icon indicators (▲/▼) for comparison to previous period
- Color coding (green for up, red for down)

### Step 7 — Dashboard Layout and Formatting

1. Set the `Dashboard` sheet to a fixed view: View → Page Layout
2. Arrange elements in a grid:
   - **Row 1-2:** Title + KPI cards
   - **Row 3:** Slicers and Timeline
   - **Row 4-5:** Main charts
   - **Row 6:** Detail tables
3. Use consistent colors: pick a 4-color palette
4. Remove chart clutter: no gridlines, minimal legends
5. Add a company logo placeholder

## Sample Data (first 15 rows)

| OrderID | OrderDate | Region | Salesperson | Product | Category | Qty | UnitPrice | Cost |
|---|---|---|---|---|---|---|---|---|
| ORD-0001 | 2024-01-05 | North | Alice Chen | Widget Pro | Hardware | 25 | 49.99 | 625.00 |
| ORD-0002 | 2024-01-08 | South | Bob Martinez | Service Plan | Services | 10 | 199.00 | 800.00 |
| ORD-0003 | 2024-01-12 | East | Carol Davis | Widget Lite | Hardware | 50 | 29.99 | 900.00 |
| ORD-0004 | 2024-01-15 | West | David Kim | Premium Suite | Software | 5 | 499.00 | 1250.00 |
| ORD-0005 | 2024-01-20 | North | Alice Chen | Service Plan | Services | 15 | 199.00 | 1200.00 |
| ORD-0006 | 2024-02-03 | South | Eve Johnson | Widget Pro | Hardware | 30 | 49.99 | 900.00 |
| ORD-0007 | 2024-02-10 | East | Frank Wilson | Basic Suite | Software | 20 | 99.00 | 1000.00 |
| ORD-0008 | 2024-02-14 | West | Grace Lee | Widget Lite | Hardware | 100 | 29.99 | 1800.00 |
| ORD-0009 | 2024-03-01 | North | Henry Patel | Premium Suite | Software | 8 | 499.00 | 2000.00 |
| ORD-0010 | 2024-03-05 | South | Alice Chen | Service Plan | Services | 12 | 199.00 | 960.00 |

## Expected Output

- Raw data table with 200+ rows across 4 quarters
- 3–4 PivotTables analyzing different dimensions
- 4 PivotCharts on the Dashboard
- 3 slicers + 1 timeline, all interconnected
- KPI cards showing Total Revenue, Total Orders, Avg Order Value, Top Region
- Charts update dynamically when slicers change
- Professional color scheme and layout

## Tips & Tricks

- **Alt + F1** creates an instant chart from selected data
- Use **PivotTable → Repeat All Item Labels** to avoid blank cells
- **Slicer shortcuts:** Ctrl+click to multi-select, Alt+S to clear filter
- Use **sparklines** (Insert → Sparklines) for inline mini-charts in tables
- Name your PivotTable ranges for cleaner GETPIVOTDATA formulas

## Common Pitfalls

- ❌ **Slicers not connected** to all PivotTables (must do manually)
- ❌ **Dates not recognized** — ensure column is formatted as Date, not Text
- ❌ **PivotTable doesn't refresh** when source data changes (right-click → Refresh)
- ❌ **Overcrowding the dashboard** — leave white space between elements
- ❌ **GETPIVOTDATA errors** when PivotTable field names change

---

# Project 3: Inventory Management System

## Overview

Build a product inventory database with automated stock tracking, reorder alerts, and in/out transaction logging. Use lookup functions to pull product details, conditional formatting to flag low stock, and basic VBA to automate repetitive tasks.

**Difficulty:** ⭐⭐⭐ Intermediate
**Time:** 6–7 hours
**Skills Used:** XLOOKUP, INDEX/MATCH, data validation, conditional formatting, VBA macros, COUNTIF, SUMIFS, running totals

## Step-by-Step Instructions

### Step 1 — Create the Product Database Sheet

Headers for `Products` sheet:

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| SKU | Product Name | Category | Supplier | Unit Cost | Sell Price | Current Stock | Reorder Level | Status |

**Status Formula (I2):**
```excel
=IF(G2<=0, "OUT OF STOCK",
  IF(G2<=H2, "REORDER NOW",
    IF(G2<=H2*1.5, "LOW STOCK", "IN STOCK")))
```

**Named Ranges:**
- `SKUList` — Column A of Products
- `ProductTable` — entire Products table

### Step 2 — Add Reorder Conditional Formatting

Select the Status column (I2:I500):

1. **OUT OF STOCK:** Red fill, white bold text
   - Formula: `=$I2="OUT OF STOCK"`
2. **REORDER NOW:** Orange fill, dark text
   - Formula: `=$I2="REORDER NOW"`
3. **LOW STOCK:** Yellow fill, dark text
   - Formula: `=$I2="LOW STOCK"`
4. **IN STOCK:** Green fill, dark text
   - Formula: `=$I2="IN STOCK"`

**Additional: Highlight entire row when stock < reorder level:**
Select A2:I500 → New Rule → Use formula:
```excel
=$G2<$H2
```
Format: Light red fill

### Step 3 — Create the Transaction Log Sheet

Headers for `Transactions` sheet:

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| TransID | Date | SKU | Type | Quantity | Unit Cost | Total | Notes |

**Type column (D):** Data Validation list: `Stock In,Stock Out,Adjustment,Return`

**Auto-fill Product Info using XLOOKUP:**
```excel
E2 (Product Name lookup — in a helper column or use XLOOKUP):
=XLOOKUP(C2, Products!A:A, Products!B:B, "Unknown SKU")

F2 (Unit Cost):
=XLOOKUP(C2, Products!A:A, Products!E:E, 0)

G2 (Total):
=IF(D2="Stock In", E2*F2, -E2*F2)
```

### Step 4 — Build the Stock Summary with Running Totals

Create a `StockSummary` sheet using SUMIFS:

```excel
For SKU "WGT-001":

Total Stock In:
=SUMIFS(Transactions!E:E, Transactions!C:C, "WGT-001",
        Transactions!D:D, "Stock In") +
 SUMIFS(Transactions!E:E, Transactions!C:C, "WGT-001",
        Transactions!D:D, "Return")

Total Stock Out:
=SUMIFS(Transactions!E:E, Transactions!C:C, "WGT-001",
        Transactions!D:D, "Stock Out") +
 SUMIFS(Transactions!E:E, Transactions!C:C, "WGT-001",
        Transactions!D:D, "Adjustment")*−1

Current Stock: =TotalIn - TotalOut
```

### Step 5 — Build the Dashboard

**Chart 1 — Stock by Category (Bar)**
Group products by category, show total current stock.

**Chart 2 — Top 10 Products by Value (Horizontal Bar)**
```excel
Stock Value: =CurrentStock * UnitCost
```

**Chart 3 — Transactions Over Time (Line)**
Sum of stock in and stock out by month.

**KPI Cards:**
- Total SKUs: `=COUNTA(Products!A:A)-1`
- Items to Reorder: `=COUNTIF(Products!I:I, "REORDER NOW")+COUNTIF(Products!I:I, "OUT OF STOCK")`
- Total Inventory Value: `=SUMPRODUCT(Products!G:G, Products!E:E)`

### Step 6 — Basic VBA: One-Click Stock Update

Press **Alt+F11** → Insert → Module:

```vba
Sub UpdateStockLevels()
    Dim wsProd As Worksheet, wsTrans As Worksheet
    Dim lastRow As Long, i As Long
    Dim sku As String, qty As Long, transType As String

    Set wsProd = Sheets("Products")
    Set wsTrans = Sheets("Transactions")

    ' Reset current stock to 0
    lastRow = wsProd.Cells(wsProd.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        wsProd.Cells(i, "G").Value = 0
    Next i

    ' Recalculate from transactions
    lastRow = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        sku = wsTrans.Cells(i, "C").Value
        qty = wsTrans.Cells(i, "E").Value
        transType = wsTrans.Cells(i, "D").Value

        ' Find the product row
        Dim prodRow As Range
        Set prodRow = wsProd.Columns("A").Find(sku, LookAt:=xlWhole)

        If Not prodRow Is Nothing Then
            Select Case transType
                Case "Stock In", "Return"
                    wsProd.Cells(prodRow.Row, "G").Value = _
                        wsProd.Cells(prodRow.Row, "G").Value + qty
                Case "Stock Out"
                    wsProd.Cells(prodRow.Row, "G").Value = _
                        wsProd.Cells(prodRow.Row, "G").Value - qty
                Case "Adjustment"
                    wsProd.Cells(prodRow.Row, "G").Value = _
                        wsProd.Cells(prodRow.Row, "G").Value + qty
            End Select
        End If
    Next i

    MsgBox "Stock levels updated!", vbInformation
End Sub
```

## Sample Data

**Products (first 10):**

| SKU | Product Name | Category | Supplier | Unit Cost | Sell Price | Stock | Reorder Lvl |
|---|---|---|---|---|---|---|---|
| WGT-001 | Widget Pro | Hardware | Acme Corp | 15.00 | 49.99 | 150 | 50 |
| WGT-002 | Widget Lite | Hardware | Acme Corp | 8.00 | 29.99 | 30 | 100 |
| SFT-001 | Basic Suite | Software | TechSoft | 40.00 | 99.00 | 200 | 50 |
| SFT-002 | Premium Suite | Software | TechSoft | 150.00 | 499.00 | 15 | 20 |
| SRV-001 | Service Plan | Services | InHouse | 50.00 | 199.00 | 500 | 100 |
| CAB-001 | USB Cable 6ft | Cables | WireCo | 2.50 | 12.99 | 8 | 200 |
| CAB-002 | HDMI Cable | Cables | WireCo | 5.00 | 24.99 | 180 | 100 |
| ACC-001 | Mouse Pad | Accessories | PadCo | 3.00 | 9.99 | 0 | 50 |
| ACC-002 | Webcam HD | Accessories | CamTech | 25.00 | 79.99 | 45 | 30 |
| PKG-001 | Starter Pack | Bundles | InHouse | 30.00 | 89.99 | 75 | 25 |

## Expected Output

- Product database with 50+ SKUs and automatic status indicators
- Transaction log with 100+ entries showing stock movements
- Conditional formatting: red rows for out-of-stock, orange for reorder
- Stock summary with running totals matching physical counts
- Dashboard with inventory value, reorder alerts, category breakdown
- VBA macro that recalculates all stock levels in one click

## Tips & Tricks

- Use **Data Validation → Error Alert** to prevent negative stock entries
- **Ctrl+Shift+L** to toggle filters on the product list
- Create a **"Quick Stock In"** form using VBA UserForm for faster data entry
- Use `SUMPRODUCT` for weighted average cost calculations
- Protect the Products sheet to prevent accidental formula edits

## Common Pitfalls

- ❌ **XLOOKUP returns #N/A** for new SKUs — always use the `if_not_found` argument
- ❌ **Stock going negative** — add a check: `=MAX(0, calculated_stock)`
- ❌ **Transaction types inconsistent** — always use the dropdown, never free-text
- ❌ **Not backing up before running VBA** — macros can't be undone (Ctrl+Z)
- ❌ **Forgetting to refresh** stock levels after manual transaction edits

---

# Project 4: Employee Schedule & Leave Tracker

## Overview

Create a workforce management tool that handles weekly schedules, tracks leave balances, and integrates company holidays. Use date functions to calculate working days, conditional formatting to visualize availability, and formulas to manage accruals.

**Difficulty:** ⭐⭐⭐ Intermediate
**Time:** 5–6 hours
**Skills Used:** NETWORKDAYS, WORKDAY, EOMONTH, conditional formatting, data validation, date functions, COUNTIF, DATEDIF

## Step-by-Step Instructions

### Step 1 — Set Up the Employee Database

Create an `Employees` sheet:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| EmpID | Name | Department | HireDate | Annual Leave | Sick Leave | Personal Days |
| EMP001 | John Smith | Engineering | 2020-03-15 | 15 | 10 | 3 |
| EMP002 | Maria Garcia | Marketing | 2021-07-01 | 15 | 10 | 3 |

### Step 2 — Create the Holiday Calendar

Create a `Holidays` sheet with a named range `HolidayList`:

| Date | Holiday Name |
|---|---|
| 2025-01-01 | New Year's Day |
| 2025-01-20 | MLK Day |
| 2025-02-17 | Presidents Day |
| 2025-05-26 | Memorial Day |
| 2025-07-04 | Independence Day |
| 2025-09-01 | Labor Day |
| 2025-11-27 | Thanksgiving |
| 2025-11-28 | Day After Thanksgiving |
| 2025-12-25 | Christmas Day |

### Step 3 — Build the Weekly Schedule Grid

Create a `Schedule` sheet:

**Header Row:** Employee Name | Mon 1/6 | Tue 1/7 | Wed 1/8 | ... | Fri 1/10

**Data Validation for shift cells:** List → `Day Shift,Night Shift,Off,Leave - Annual,Leave - Sick,Leave - Personal,Training,Remote`

**Conditional Formatting for the grid:**
- `Day Shift` → Light blue fill
- `Night Shift` → Dark blue fill, white text
- `Off` → Gray fill
- `Leave - Annual` → Green fill
- `Leave - Sick` → Orange fill
- `Leave - Personal` → Purple fill
- Holiday dates → Red fill with holiday name

### Step 4 — Build the Leave Tracker

Create a `LeaveTracker` sheet:

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| EmpID | Name | Leave Type | Start Date | End Date | Working Days | Status | Approved By | Notes |

**Working Days Formula (F2):**
```excel
=NETWORKDAYS(D2, E2, HolidayList)
```
This automatically excludes weekends and company holidays.

**Leave Balance Summary (separate table):**

| EmpID | Name | Annual Entitled | Annual Used | Annual Remaining | Sick Entitled | Sick Used | Sick Remaining |
|---|---|---|---|---|---|---|---|

```excel
Annual Used:
=SUMIFS(LeaveTracker!F:F, LeaveTracker!A:A, A2,
        LeaveTracker!C:C, "Leave - Annual",
        LeaveTracker!G:G, "Approved")

Annual Remaining: =C2-E2
```

### Step 5 — Create the Monthly Calendar View

Build a visual calendar for each month:

1. Create a grid: Columns = days 1–31, Rows = employees
2. Use WORKDAY and EOMONTH to auto-generate dates:
   ```excel
   =WORKDAY(EOMONTH(DATE(2025,1,1),-1), 1)  → First workday of Jan
   ```
3. Use INDEX/MATCH to pull schedule data for each employee/day
4. Apply conditional formatting based on shift type

### Step 6 — Dashboard: Attendance Overview

**Charts:**
1. **Leave Usage by Type (Pie):** Annual vs Sick vs Personal usage
2. **Monthly Absences (Bar):** Count of leave days per month
3. **Department Attendance (Stacked Bar):** Present vs Absent by dept
4. **Leave Balance Summary (Horizontal Bar):** Remaining days per employee

**KPIs:**
- Total employees on leave today: `=COUNTIF(today_column, "*Leave*")`
- Average attendance rate: `=(total_workdays - total_absences) / total_workdays`
- Upcoming holidays: Next holiday from HolidayList

## Sample Data

**Leave Requests:**

| EmpID | Name | Leave Type | Start Date | End Date | Working Days | Status |
|---|---|---|---|---|---|---|
| EMP001 | John Smith | Leave - Annual | 2025-03-10 | 2025-03-14 | 5 | Approved |
| EMP002 | Maria Garcia | Leave - Sick | 2025-02-03 | 2025-02-04 | 2 | Approved |
| EMP003 | James Wilson | Leave - Annual | 2025-04-21 | 2025-04-25 | 5 | Pending |
| EMP001 | John Smith | Leave - Personal | 2025-05-02 | 2025-05-02 | 1 | Approved |
| EMP004 | Sarah Lee | Leave - Annual | 2025-12-22 | 2025-12-31 | 6 | Approved |

## Expected Output

- Employee database with leave entitlements
- Holiday calendar with named range
- Weekly schedule grid with color-coded shifts
- Leave tracker calculating working days automatically
- Leave balance summary per employee showing used and remaining
- Monthly calendar view with conditional formatting
- Dashboard with attendance KPIs and charts

## Tips & Tricks

- **NETWORKDAYS(start, end, holidays)** is your best friend for leave calculations
- Use `TODAY()` in formulas to highlight the current day on the schedule
- Create a **printable weekly schedule** with Page Layout view
- Use **Data Validation Input Message** to explain what each shift code means
- Protect cells with formulas so only schedule cells are editable

## Common Pitfalls

- ❌ **Holiday list not updated** annually — set a reminder for January
- ❌ **NETWORKDAYS including the holiday list range but missing new holidays**
- ❌ **Leave spanning month boundaries** — ensure SUMIFS covers the full year
- ❌ **Time zones not considered** for remote workers in different zones
- ❌ **Not distinguishing between "Pending" and "Approved"** leave in balance calculations

---

# Project 5: Invoice Generator Template

## Overview

Build a professional, print-ready invoice template that auto-calculates subtotals, taxes, and totals. Include a customer database, product catalog with lookup, and professional formatting suitable for client-facing documents.

**Difficulty:** ⭐⭐ Beginner-Intermediate
**Time:** 4–5 hours
**Skills Used:** Named ranges, VLOOKUP/XLOOKUP, templates, print setup, cell formatting, IF, ROUND, page layout

## Step-by-Step Instructions

### Step 1 — Create Supporting Sheets

**Sheet 1: `CompanyInfo`**

| Field | Value |
|---|---|
| Company Name | Your Company LLC |
| Address | 123 Business Ave, Suite 100 |
| City, State ZIP | New York, NY 10001 |
| Phone | (555) 123-4567 |
| Email | billing@yourcompany.com |
| Website | www.yourcompany.com |
| Tax ID | XX-XXXXXXX |
| Bank Name | First National Bank |
| Account Number | XXXX-XXXX-1234 |
| Routing Number | XXXXXXXXX |

**Sheet 2: `CustomerDB`**

| CustID | Company Name | Contact | Address | City | State | ZIP | Email | Phone |
|---|---|---|---|---|---|---|---|---|
| C001 | Acme Corp | John Doe | 456 Client St | Boston | MA | 02101 | john@acme.com | (555) 234-5678 |

**Sheet 3: `ProductCatalog`**

| ProdID | Description | Unit Price | Taxable |
|---|---|---|---|
| P001 | Widget Pro | 49.99 | Yes |
| P002 | Consulting (hourly) | 150.00 | Yes |
| P003 | Shipping | 15.00 | No |

**Named Ranges:**
- `CustomerTable` — CustomerDB data
- `ProductTable` — ProductCatalog data
- `CompanyName`, `CompanyAddress`, etc. — from CompanyInfo

### Step 2 — Design the Invoice Layout

Create the `Invoice` sheet with the following layout:

**Rows 1-5: Header**
```
[Company Logo Placeholder]     YOUR COMPANY LLC
                               123 Business Ave, Suite 100
                               New York, NY 10001
                               (555) 123-4567

                               INVOICE
```

**Rows 7-12: Invoice Details**
```
Invoice #:     [INV-2025-001]          Bill To:
Date:          [Auto: =TODAY()]         [Customer Name from lookup]
Due Date:      [Auto: +30 days]        [Address]
Payment Terms: Net 30                  [City, State ZIP]
                                       [Email]
```

**Customer Lookup:**
```excel
Customer Name: =XLOOKUP(InvoiceCustomerID, CustomerDB!A:A, CustomerDB!B:B)
Address:       =XLOOKUP(InvoiceCustomerID, CustomerDB!A:A, CustomerDB!D:D)
```

**Rows 14-25: Line Items Table**

| # | Description | Qty | Unit Price | Line Total |
|---|---|---|---|---|
| 1 | [Product lookup] | 10 | [Auto from catalog] | [Qty × Price] |
| 2 | | | | |
| 3 | | | | |
| ... | | | | |

**Formulas:**
```excel
Description: =XLOOKUP(B14, ProductCatalog!A:A, ProductCatalog!B:B, "")
Unit Price:  =XLOOKUP(B14, ProductCatalog!A:A, ProductCatalog!C:C, 0)
Line Total:  =C14*D14
```

**Rows 27-31: Totals**
```
                            Subtotal:     =SUM(E14:E25)
                            Tax (8.875%): =Subtotal * TaxRate
                            Shipping:     [manual entry]
                            TOTAL DUE:    =Subtotal + Tax + Shipping
```

**Rows 33-37: Footer**
```
Notes: [Optional text]
Payment Instructions: Please make check payable to Your Company LLC
Bank: First National Bank | Acct: XXXX-XXXX-1234 | Routing: XXXXXXXXX
Thank you for your business!
```

### Step 3 — Professional Formatting

- **Company name:** 18pt bold, dark blue
- **"INVOICE" title:** 24pt bold, centered
- **Line items header:** Bold, dark blue fill, white text
- **Alternating row colors:** Light gray on even rows
- **TOTAL DUE:** Bold, 14pt, top border, light blue fill
- **Currency format:** `$#,##0.00` for all money fields
- **Date format:** `MMMM D, YYYY` (e.g., "January 15, 2025")

### Step 4 — Print Setup

1. Page Layout → Size: Letter
2. Margins: Narrow (0.5" all sides)
3. Orientation: Portrait
4. Print Area: Select the invoice area → Set Print Area
5. Header/Footer: Add page number ("Page &P of &N")
6. Scale: Fit to 1 page wide by 1 page tall
7. Gridlines: Uncheck (for clean print)
8. **Test print** to PDF before using

### Step 5 — Make It Reusable

1. **Save as template:** File → Save As → Excel Template (.xltx)
2. **Invoice number auto-increment:** Use a helper cell:
   ```excel
   ="INV-" & TEXT(YEAR(TODAY()), "0000") & "-" & TEXT(NextInvoiceNum, "000")
   ```
3. **Due date auto-calc:** `=InvoiceDate + PaymentTermsDays`

## Sample Data

**Product Catalog:**

| ProdID | Description | Unit Price | Taxable |
|---|---|---|---|
| P001 | Widget Pro | 49.99 | Yes |
| P002 | Widget Lite | 29.99 | Yes |
| P003 | Consulting - Standard | 150.00 | Yes |
| P004 | Consulting - Premium | 250.00 | Yes |
| P005 | Service Plan - Annual | 1,999.00 | Yes |
| P006 | Setup Fee | 500.00 | Yes |
| P007 | Shipping - Standard | 15.00 | No |
| P008 | Shipping - Express | 35.00 | No |
| P009 | Training (per session) | 300.00 | Yes |
| P010 | Custom Development (hourly) | 175.00 | Yes |

## Expected Output

- Professional-looking invoice template
- Auto-populating company and customer info
- Product catalog lookup for line items
- Automatic subtotal, tax, and total calculations
- Print-ready layout that fits on one page
- Saved as .xltx for reuse
- Invoice numbering system

## Tips & Tricks

- Use **conditional formatting** to highlight the TOTAL row
- Add a **"Paid" stamp** as a WordArt or text box overlay
- Create a **duplicate sheet** for credit notes (negative quantities)
- Use `ROUND(tax_calc, 2)` to avoid fractional cent errors
- Save multiple templates for different tax rates or currencies

## Common Pitfalls

- ❌ **Tax calculated on non-taxable items** — use SUMPRODUCT with a taxable flag
- ❌ **Print area not set** — prints blank pages or cuts off content
- ❌ **Currency formatting inconsistent** — apply to entire columns at once
- ❌ **Invoice number not unique** — use a separate counter sheet
- ❌ **Forgetting to update the year** in invoice numbering

---

# Project 6: Data Cleaning Pipeline

## Overview

Build a repeatable data cleaning workflow using Power Query and text functions. Import messy CSV data, fix encoding issues, remove duplicates, standardize formats, and output a clean dataset ready for analysis. This project teaches one of the most valuable Excel skills: transforming dirty data.

**Difficulty:** ⭐⭐⭐ Intermediate
**Time:** 5–6 hours
**Skills Used:** Power Query, TRIM, CLEAN, SUBSTITUTE, PROPER, UPPER, LOWER, TEXT, LEFT, RIGHT, MID, Find & Replace, Remove Duplicates, Text to Columns

## Step-by-Step Instructions

### Step 1 — Import Messy Data

**Source:** Download or create a messy CSV with common data issues:

1. Data → Get Data → From File → From Text/CSV
2. Select the file → Preview dialog appears
3. Check encoding (UTF-8 usually correct) → Load

### Step 2 — Identify Data Quality Issues

Open Power Query Editor (Data → Get Data → Launch Power Query Editor):

**Common issues to look for:**
- Leading/trailing spaces
- Mixed case ("new york", "NEW YORK", "New York")
- Duplicates (exact and fuzzy)
- Missing values (blank, N/A, "null", "-")
- Inconsistent date formats (MM/DD/YYYY vs DD/MM/YYYY)
- Phone numbers in different formats
- Addresses with abbreviations
- Special characters and encoding artifacts

### Step 3 — Power Query Transformations

**In Power Query Editor, apply these steps:**

**3a. Remove extra spaces:**
```
Transform → Format → Trim
```

**3b. Standardize text case:**
```
For names:   Transform → Format → Capitalize Each Word
For emails:  Transform → Format → Lowercase
For codes:   Transform → Format → Uppercase
```

**3c. Remove special characters:**
```
Transform → Replace Values
Value to Find: specific characters
Replace With: (empty)
```

**3d. Handle missing values:**
```
Transform → Replace Values
Value: null → Replace with: (blank or "N/A")
Value: "-" → Replace with: (blank)
Value: "N/A" → Replace with: (blank)
```

**3e. Remove duplicates:**
```
Home → Remove Rows → Remove Duplicates
Select key columns (e.g., Email, FullName)
```

**3f. Standardize dates:**
```
Transform → Data Type → Date
If mixed formats, use: Date.From([DateString], "MM/dd/yyyy")
```

**3g. Parse and clean phone numbers:**
```
Add Column → Custom Column:
= Text.Select([Phone], {"0".."9"})
Then format: = "(" & Text.Range(CleanPhone,0,3) & ") " & ...
```

**3h. Split/merge columns:**
```
FullName → Split Column by Delimiter (space) → First, Last
Address → Split Column by Delimiter (comma) → Street, City, State ZIP
```

### Step 4 — Advanced: Fuzzy Matching for Deduplication

Power Query supports fuzzy matching:

```
Home → Remove Rows → Remove Duplicates
→ Right-click column → Remove Duplicates
→ Options: Similarity threshold = 0.8
```

Or use a Merge query with fuzzy matching:
```
Home → Merge Queries
Select join columns → Check "Use fuzzy matching"
Set similarity threshold
```

### Step 5 — Text Function Alternatives (Without Power Query)

If not using Power Query, use these formulas on a helper column:

```excel
Remove spaces:        =TRIM(A2)
Remove non-printable: =CLEAN(A2)
Capitalize names:     =PROPER(A2)
Remove commas:        =SUBSTITUTE(A2, ",", "")
Extract first name:   =LEFT(A2, FIND(" ", A2)-1)
Extract last name:    =RIGHT(A2, LEN(A2)-FIND(" ", A2))
Standardize phone:    ="("&MID(A2,1,3)&") "&MID(A2,4,3)&"-"&MID(A2,7,4)
Clean email:          =LOWER(TRIM(SUBSTITUTE(A2, " ", "")))
```

### Step 6 — Output Clean Dataset

1. In Power Query: Home → Close & Load → To new worksheet
2. Rename the output sheet: `CleanData`
3. **Verify:** Spot-check 10-20 rows against the original
4. **Document changes:** Create a `CleaningLog` sheet listing all transformations applied

### Step 7 — Build a Reusable Cleaning Template

1. Save the Power Query steps as a reusable query
2. Create parameters for file path so you can change the source
3. Add a "Refresh" button that re-runs the entire pipeline
4. Document the expected input format

## Sample Data (Messy)

| FullName | Email | Phone | City | State | ZipCode | DateJoined |
|---|---|---|---|---|---|---|
| john smith | John.Smith@Gmail.COM | (555)123-4567 | new york | NY | 10001 | 01/15/2023 |
| JANE DOE | jane.doe@yahoo.com | 555.234.5678 | Los Angeles | ca | 90001 | 2023-03-20 |
| bob  Johnson | BOB@company.org | 5553456789 | chicago | Illinois | 60601 | 03/15/2023 |
| John Smith | john.smith@gmail.com | (555) 123-4567 | New York | NY | 10001 | 01/15/2023 |
| Alice Brown | alice.brown@ | 555.456.7890 |  | TX | 75001 | - |
| null | - | N/A | Miami | FL | 33101 | 2023-06-01 |
| María García | maria@email.com | +1-555-567-8901 | San Antonio | TX | 78201 | 07/04/2023 |
| Charlie Wilson | CHARLIE.W@WORK.COM | (555)678-9012 | seattle | WA | 98101 | 08/30/2023 |

**Issues present:**
- Mixed case names, emails
- Extra spaces (bob  Johnson)
- Duplicate person (john smith / John Smith)
- Different phone formats
- Missing values (null, -, N/A, blank)
- Different date formats
- Invalid email (alice.brown@)
- Inconsistent state (ca vs CA vs California)

## Expected Output

- Clean dataset with all issues resolved
- Consistent capitalization (Proper Case for names, lowercase for emails)
- Standardized phone format: (XXX) XXX-XXXX
- All dates in MM/DD/YYYY format
- Duplicates removed (2 → 1)
- Missing values handled consistently
- A cleaning log documenting all transformations
- Reusable Power Query pipeline

## Tips & Tricks

- **Always keep the original data** — clean into a new sheet/table
- Use `TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," ")))` for stubborn spaces (non-breaking)
- **Flash Fill** (Ctrl+E) can auto-detect patterns for name splitting
- Power Query's **Applied Steps** pane is your undo history — review each step
- Use `=IFERROR(VALUE(A2), A2)` to convert text-numbers to numbers

## Common Pitfalls

- ❌ **Deleting rows** instead of cleaning them — you lose data
- ❌ **Not checking for fuzzy duplicates** — "Jon Smith" vs "John Smith"
- ❌ **Changing data types before cleaning** — clean first, then set types
- ❌ **Power Query case sensitivity** — "New York" ≠ "new york" in some operations
- ❌ **Not refreshing the query** after source data changes

---

# Project 7: Financial Model

## Overview

Build a 5-year financial projection model with revenue forecasting, expense modeling, break-even analysis, loan amortization, and sensitivity analysis. This is the most formula-intensive project and is directly applicable to business planning and investment analysis.

**Difficulty:** ⭐⭐⭐⭐ Advanced
**Time:** 7–8 hours
**Skills Used:** NPV, IRR, PMT, PPMT, IPMT, FV, PV, Data Tables (1-var, 2-var), Scenario Manager, Goal Seek, INDEX/MATCH, conditional formatting

## Step-by-Step Instructions

### Step 1 — Model Structure

Create five sheets:

| Sheet | Purpose |
|---|---|
| `Assumptions` | All input variables in one place |
| `Revenue` | Revenue projections |
| `Expenses` | Cost and expense model |
| `Financials` | Income statement, cash flow |
| `Analysis` | NPV, IRR, break-even, sensitivity |

### Step 2 — Build the Assumptions Sheet

**Revenue Assumptions:**

| Parameter | Value | Notes |
|---|---|---|
| Starting Customers | 100 | Year 1 beginning |
| Monthly Growth Rate | 5% | New customer acquisition |
| Churn Rate | 2% | Monthly customer loss |
| Avg Revenue Per User (ARPU) | $50/month | |
| ARPU Growth | 3%/year | Annual price increase |

**Cost Assumptions:**

| Parameter | Value |
|---|---|
| Fixed Costs (monthly) | $10,000 |
| Variable Cost per Customer | $15 |
| Cost Inflation | 2%/year |
| Initial Investment | $250,000 |
| Loan Amount | $200,000 |
| Loan Interest Rate | 6% annually |
| Loan Term (years) | 5 |
| Tax Rate | 25% |

**Name each cell** for easy reference (e.g., `GrowthRate`, `ChurnRate`, `ARPU`)

### Step 3 — Revenue Projections

**Monthly Customer Model:**
```excel
Month 1 Customers: =StartingCustomers
Month 2 Customers: =Month1 * (1 + GrowthRate - ChurnRate)
Month N Customers: =Month(N-1) * (1 + GrowthRate - ChurnRate)
```

**Monthly Revenue:**
```excel
=Customers * ARPU
```

**Annual Revenue:** Sum of monthly revenues for the year.

### Step 4 — Expense Model

**Monthly Expenses:**
```excel
Fixed:   =FixedCosts * (1 + CostInflation) ^ (YearNum - 1)
Variable: =Customers * VariableCostPerCustomer
Total:   =Fixed + Variable
```

### Step 5 — Income Statement

| | Year 1 | Year 2 | Year 3 | Year 4 | Year 5 |
|---|---|---|---|---|---|
| Revenue | | | | | |
| COGS (Variable) | | | | | |
| Gross Profit | | | | | |
| Operating Expenses (Fixed) | | | | | |
| EBITDA | | | | | |
| Depreciation | | | | | |
| EBIT | | | | | |
| Interest Expense | | | | | |
| EBT | | | | | |
| Taxes | | | | | |
| Net Income | | | | | |

**Key Formulas:**
```excel
Gross Profit:    =Revenue - COGS
EBITDA:          =Gross Profit - Operating Expenses
EBIT:            =EBITDA - Depreciation
Interest:        (from amortization schedule)
EBT:             =EBIT - Interest
Taxes:           =MAX(0, EBT * TaxRate)
Net Income:      =EBT - Taxes
```

### Step 6 — Loan Amortization Schedule

| Month | Beginning Balance | Payment | Principal | Interest | Ending Balance |
|---|---|---|---|---|---|

```excel
Monthly Payment (cell B1): =PMT(Rate/12, Term*12, -LoanAmount)
Beginning Balance (Month 1): =LoanAmount
Interest: =BeginningBalance * (Rate/12)
Principal: =Payment - Interest
Ending Balance: =BeginningBalance - Principal
```

**Built-in functions for specific periods:**
```excel
Principal in month 13: =PPMT(Rate/12, 13, Term*12, -LoanAmount)
Interest in month 13:  =IPMT(Rate/12, 13, Term*12, -LoanAmount)
```

### Step 7 — NPV and IRR Analysis

**Net Present Value:**
```excel
Cash Flows: Year 0 = -InitialInvestment, Years 1-5 = Net Income + Depreciation
=NPV(DiscountRate, Year1:Year5) + Year0
```

**Internal Rate of Return:**
```excel
=IRR(Year0:Year5)
```

**Payback Period:** Use cumulative cash flow to find when it turns positive.

### Step 8 — Break-Even Analysis

**Break-Even Point (units):**
```excel
=FixedCosts / (ARPU - VariableCostPerCustomer)
```

**Break-Even with Goal Seek:**
1. Data → What-If Analysis → Goal Seek
2. Set cell: Net Income (Year 1)
3. To value: 0
4. By changing: Growth Rate
5. Click OK → Excel finds the minimum growth rate for profitability

### Step 9 — Sensitivity Analysis with Data Tables

**1-Variable Data Table (Revenue sensitivity to growth rate):**

Set up:
| Growth Rate | Revenue |
|---|---|
| 1% | =RevenueFormula |
| 2% | |
| 3% | |
| ... | |

1. Select the table range (including the formula cell)
2. Data → What-If Analysis → Data Table
3. Column input cell: GrowthRate cell
4. Excel fills in revenue for each growth rate

**2-Variable Data Table (Revenue sensitivity to growth rate AND ARPU):**

| | $30 ARPU | $40 ARPU | $50 ARPU | $60 ARPU |
|---|---|---|---|---|
| 2% growth | =RevenueFormula | | | |
| 4% growth | | | | |
| 6% growth | | | | |
| 8% growth | | | | |

1. Select the table range
2. Data → What-If Analysis → Data Table
3. Row input cell: ARPU cell
4. Column input cell: GrowthRate cell

**Conditional Formatting on the 2-var table:**
- Color scale: Red (low) → Green (high)
- Highlight profitable combinations in green

### Step 10 — Scenario Manager

1. Data → What-If Analysis → Scenario Manager
2. Create three scenarios:

| Scenario | Growth Rate | ARPU | Variable Cost |
|---|---|---|---|
| Best Case | 8% | $60 | $12 |
| Base Case | 5% | $50 | $15 |
| Worst Case | 2% | $40 | $20 |

3. Show each scenario to see the impact on Net Income and NPV
4. Create a Scenario Summary report

## Sample Data

**Assumptions for Base Case:**

| Input | Value |
|---|---|
| Starting Customers | 100 |
| Monthly Growth Rate | 5% |
| Monthly Churn Rate | 2% |
| ARPU | $50/month |
| Fixed Costs | $10,000/month |
| Variable Cost/Customer | $15/month |
| Initial Investment | $250,000 |
| Loan | $200,000 @ 6% for 5 years |
| Discount Rate | 10% |
| Tax Rate | 25% |

## Expected Output

- 5-year revenue projection with monthly granularity
- Complete income statement (Revenue → Net Income)
- Loan amortization schedule (60 months)
- NPV and IRR calculations
- Break-even point calculation
- 1-variable and 2-variable sensitivity data tables
- 3 scenarios (Best/Base/Worst) with summary
- Dashboard showing key financial metrics over 5 years

## Tips & Tricks

- **Keep all assumptions on one sheet** — never hardcode numbers in formulas
- Use **named ranges** for every assumption cell
- Color code: Blue font = input, Black font = formula (financial modeling convention)
- **Ctrl+`** (backtick) toggles formula view — useful for auditing
- Use `IFERROR` to handle division by zero in margin calculations

## Common Pitfalls

- ❌ **Circular references** in interest calculations (interest depends on balance, which depends on payments)
- ❌ **Not using absolute references** for assumptions in data tables
- ❌ **Forgetting to add back depreciation** in cash flow (it's non-cash)
- ❌ **Data tables recalculating slowly** — use Ctrl+Alt+F9 to force recalc
- ❌ **Inconsistent time periods** — don't mix monthly and annual figures

---

# Project 8: Automated Report Generator

## Overview

Build a VBA-powered macro that automatically refreshes data, generates PivotTables and charts, exports the report as a PDF, and emails it to a distribution list. This is the capstone project that ties together everything learned.

**Difficulty:** ⭐⭐⭐⭐⭐ Advanced
**Time:** 7–8 hours
**Skills Used:** VBA macros, PivotTables, charts, Outlook automation, PDF export, error handling, user forms, progress indicators

## Step-by-Step Instructions

### Step 1 — Design the Report Template

Create a `Report` sheet with:
- Header section: Title, date, prepared by
- KPI summary section
- PivotTable section
- Chart section
- Footer: Page numbers, confidentiality notice

### Step 2 — Create the Data Connection

1. Data → Get Data → From your data source (SQL, CSV, SharePoint, etc.)
2. Load to PivotTable cache
3. Name the query: `ReportData`

### Step 3 — Build the VBA Macro

Press **Alt+F11** → Insert → Module:

```vba
Option Explicit

Sub GenerateReport()
    ' ============================================
    ' Automated Report Generator
    ' Refreshes data, updates charts, exports PDF,
    ' and emails the report
    ' ============================================

    On Error GoTo ErrorHandler

    Dim wsReport As Worksheet
    Dim pvt As PivotTable
    Dim reportDate As String
    Dim pdfPath As String
    Dim startTime As Double

    startTime = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- Step 1: Update status ---
    UpdateStatus "Starting report generation..."

    ' --- Step 2: Refresh all data connections ---
    UpdateStatus "Refreshing data..."
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    ' --- Step 3: Update report date ---
    Set wsReport = Sheets("Report")
    wsReport.Range("ReportDate").Value = Format(Date, "MMMM D, YYYY")

    ' --- Step 4: Refresh PivotTables ---
    UpdateStatus "Updating PivotTables..."
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        For Each pvt In ws.PivotTables
            pvt.RefreshTable
            pvt.Update
        Next pvt
    Next ws

    ' --- Step 5: Export to PDF ---
    UpdateStatus "Exporting PDF..."
    reportDate = Format(Date, "YYYY-MM-DD")
    pdfPath = ThisWorkbook.Path & "\Reports\Report_" & reportDate & ".pdf"

    ' Create Reports folder if it doesn't exist
    If Dir(ThisWorkbook.Path & "\Reports", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\Reports"
    End If

    wsReport.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' --- Step 6: Send email ---
    UpdateStatus "Sending email..."
    Call SendReportEmail(pdfPath)

    ' --- Step 7: Complete ---
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    UpdateStatus "Report generated in " & _
        Format(Timer - startTime, "0.0") & " seconds"
    MsgBox "Report generated and emailed successfully!" & vbNewLine & _
           "Saved to: " & pdfPath, vbInformation, "Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Report Generation Failed"
End Sub

Sub SendReportEmail(pdfPath As String)
    ' Requires Microsoft Outlook
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim recipientList As String

    ' Get recipient list from Recipients sheet
    recipientList = ""
    Dim wsRecip As Worksheet
    Set wsRecip = Sheets("Recipients")
    Dim i As Long
    For i = 2 To wsRecip.Cells(wsRecip.Rows.Count, "A").End(xlUp).Row
        If wsRecip.Cells(i, "B").Value = "Active" Then
            recipientList = recipientList & wsRecip.Cells(i, "A").Value & ";"
        End If
    Next i

    ' Create email
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)

    With mailItem
        .To = Left(recipientList, Len(recipientList) - 1)
        .Subject = "Weekly Report - " & Format(Date, "MMMM D, YYYY")
        .Body = "Hello," & vbNewLine & vbNewLine & _
                "Please find attached this week's report." & vbNewLine & _
                vbNewLine & "Best regards," & vbNewLine & _
                "Automated Report System"
        .Attachments.Add pdfPath
        .Display  ' Use .Send to auto-send
    End With

    Set mailItem = Nothing
    Set outlookApp = Nothing
End Sub

Sub UpdateStatus(statusMsg As Application.StatusBar)
    Application.StatusBar = statusMsg
    DoEvents
End Sub
```

### Step 4 — Create the Recipients Sheet

| Email | Status | Department |
|---|---|---|
| manager@company.com | Active | Management |
| team-lead@company.com | Active | Operations |
| analyst@company.com | Active | Finance |
| former-employee@company.com | Inactive | Former |

### Step 5 — Add a Run Button

1. Insert → Shapes → Rounded Rectangle
2. Right-click → Assign Macro → `GenerateReport`
3. Format the button with color and text "📊 Generate Report"
4. Place it at the top of the Report sheet

### Step 6 — Error Handling and Logging

Add a `Log` sheet and logging subroutine:

```vba
Sub LogAction(action As String)
    Dim wsLog As Worksheet
    Set wsLog = Sheets("Log")
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    wsLog.Cells(nextRow, 1).Value = Now()
    wsLog.Cells(nextRow, 2).Value = Application.UserName
    wsLog.Cells(nextRow, 3).Value = action
End Sub
```

### Step 7 — Schedule Automatic Execution (Optional)

To run the report on a schedule, use Windows Task Scheduler:

1. Create a .bat file:
   ```bat
   "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /r "C:\Reports\ReportGenerator.xlsm"
   ```
2. Open Task Scheduler → Create Basic Task
3. Set trigger (e.g., every Monday at 8 AM)
4. Set action to run the .bat file

## Sample Data

**Recipients (already shown above)**

**Report Content (auto-generated from source data):**
- Sales summary by region (PivotTable)
- Monthly trend chart
- Top 10 products bar chart
- KPI cards: Revenue, Orders, Avg Order Value
- Comparison to previous period

## Expected Output

- One-click report generation button
- Auto-refreshed data and PivotTables
- Professional PDF report with charts and KPIs
- Auto-email to distribution list
- Generation log with timestamps
- Error handling for common issues (no data, Outlook not running, etc.)
- Execution time displayed in status bar

## Tips & Tricks

- **Always test with `.Display` first**, then switch to `.Send` when confident
- Use `Application.StatusBar` for progress updates during long operations
- **Save as .xlsm** (macro-enabled) — regular .xlsx won't save VBA code
- Add `Application.CutCopyMode = False` after paste operations to clear clipboard
- Use `DoEvents` to keep Excel responsive during long operations

## Common Pitfalls

- ❌ **Forgetting to save as .xlsm** — macros disappear when saved as .xlsx
- ❌ **Outlook security prompts** — user may need to allow programmatic access
- ❌ **PDF export failing** if the report sheet has print area issues
- ❌ **Not handling "no data" scenarios** — add checks before PivotTable refresh
- ❌ **Macro security blocking execution** — ensure trusted location or signed macro

---

# Industry Use Cases

## Finance & Banking

| Use Case | Excel Features | Example |
|---|---|---|
| Financial reporting | PivotTables, charts, Power Query | Monthly P&L, balance sheets |
| Budgeting & forecasting | What-If Analysis, Data Tables, scenarios | Annual budget with variance analysis |
| Loan analysis | PMT, NPV, IRR, amortization | Mortgage calculator, loan comparison |
| Risk assessment | Monte Carlo (add-in), probability distributions | Portfolio risk modeling |
| Audit trails | Track Changes, comments, cell protection | Financial audit workpapers |

## Human Resources

| Use Case | Excel Features | Example |
|---|---|---|
| Payroll processing | VLOOKUP, tax tables, SUMPRODUCT | Monthly payroll calculation |
| Employee scheduling | NETWORKDAYS, conditional formatting | Shift planning, leave management |
| Recruitment tracking | Data validation, filters, COUNTIF | Candidate pipeline management |
| Performance reviews | Rating scales, radar charts | Employee evaluation scorecards |
| Headcount planning | PivotTables, trend analysis | Workforce planning dashboard |

## Marketing

| Use Case | Excel Features | Example |
|---|---|---|
| Campaign analysis | PivotTables, charts, A/B test calculations | Email campaign performance |
| Customer segmentation | PivotTables, conditional formatting | RFM analysis (Recency, Frequency, Monetary) |
| Content calendar | Date functions, conditional formatting | Monthly content planning |
| ROI tracking | SUMIFS, percentage calculations | Marketing spend vs revenue attribution |
| Social media reporting | Power Query (API import), charts | Weekly social media metrics |

## Operations & Supply Chain

| Use Case | Excel Features | Example |
|---|---|---|
| Inventory management | XLOOKUP, conditional formatting, VBA | Stock tracking, reorder alerts |
| Production planning | Solver add-in, linear programming | Optimal production mix |
| Quality control | Statistical functions, control charts | Defect tracking, SPC charts |
| Vendor management | Data validation, scoring models | Supplier evaluation matrix |
| Logistics optimization | Distance matrices, route planning | Delivery route cost analysis |

## Education

| Use Case | Excel Features | Example |
|---|---|---|
| Grade book | AVERAGE, weighted grades, conditional formatting | Student grade tracking |
| Attendance tracking | COUNTIF, NETWORKDAYS | Daily attendance with % calculations |
| Test analysis | Statistical functions, histograms | Item analysis, score distributions |
| Budget management | SUMIFS, charts | Department budget tracking |
| Scheduling | Date functions, Gantt charts | Course schedule, room allocation |

## Healthcare

| Use Case | Excel Features | Example |
|---|---|---|
| Patient tracking | Data validation, conditional formatting | Patient intake, appointment scheduling |
| Clinical trial data | Statistical analysis, data cleaning | Trial results analysis |
| Inventory management | XLOOKUP, alerts | Medical supply tracking |
| Staff scheduling | NETWORKDAYS, shift patterns | Nurse scheduling, on-call rotation |
| Financial reporting | PivotTables, charts | Department cost analysis |

---

# Excel Certification Paths

## Microsoft Office Specialist (MOS) — Excel Associate

**Exam:** MO-200 (Excel 2019) or MO-210 (Microsoft 365 Apps)

**Skills Measured:**
- Manage worksheets and workbooks (15–20%)
- Manage data cells and ranges (20–25%)
- Manage tables and table data (15–20%)
- Perform operations by using formulas and functions (20–25%)
- Manage charts (15–20%)

**Preparation:**
- Study time: 40–60 hours
- Practice with Excel 2019/365 (not older versions)
- Use GMetrix or Certiport practice tests
- Focus areas: PivotTables, formulas, charts, cell formatting

**Career Value:** Entry-level validation, good for administrative and analyst roles

## Microsoft Office Specialist (MOS) — Excel Expert

**Exam:** MO-201 (Excel 2019) or MO-211 (Microsoft 365 Apps)

**Skills Measured:**
- Manage workbook options and settings (10–15%)
- Apply custom data formats and layouts (15–20%)
- Create advanced formulas (20–25%)
- Create advanced charts and tables (25–30%)
- Manage PivotTables and PivotCharts (15–20%)

**Key Differences from Associate:**
- Advanced formulas: INDEX/MATCH, array formulas, XLOOKUP, LAMBDA
- Power Query basics
- Advanced PivotTable features (calculated fields, groups)
- Complex chart types (combo, waterfall, funnel)

**Preparation:**
- Study time: 60–80 hours
- Requires Associate certification first (recommended, not required)
- Focus on scenario-based questions

**Career Value:** Strong differentiator for data analysts and finance professionals

## Microsoft Office Specialist (MOS) — Master

**Requirements:** Pass 4 exams:
1. Excel Expert (MO-201 or MO-211)
2. Word Expert (MO-101 or MO-111)
3. PowerPoint Associate (MO-300 or MO-310)
4. Outlook Associate (MO-400 or MO-410) OR Access Expert (MO-500)

**Career Value:** Highest Office certification, signals comprehensive Microsoft 365 proficiency. Valued for executive assistants, office managers, and training roles.

## Other Relevant Certifications

| Certification | Focus | Level |
|---|---|---|
| **CFA (Chartered Financial Analyst)** | Financial modeling in Excel | Advanced |
| **FMVA (Financial Modeling & Valuation Analyst)** | Excel financial modeling | Intermediate-Advanced |
| **Google Sheets Certification** | Google Workspace equivalent | Intermediate |
| **Tableau Desktop Specialist** | Visualization (Excel skills transfer) | Intermediate |
| **Power BI Data Analyst (PL-300)** | Power BI (builds on Excel skills) | Intermediate |

---

# Excel vs Google Sheets

## Feature Comparison

| Feature | Excel | Google Sheets | Winner |
|---|---|---|---|
| **Max rows** | 1,048,576 | 10,000,000 cells total | Google |
| **Max columns** | 16,384 | 18,278 | Google |
| **Offline access** | Full (desktop app) | Limited (requires setup) | Excel |
| **Real-time collaboration** | Good (Microsoft 365) | Excellent (native) | Google |
| **PivotTables** | Full-featured | Basic but improving | Excel |
| **Power Query** | Full ETL tool | Not available | Excel |
| **VBA/Macros** | Full VBA | Apps Script (JavaScript) | Excel |
| **PivotCharts** | Yes | No (manual charts only) | Excel |
| **Conditional Formatting** | Advanced | Good | Excel |
| **Array Formulas** | Dynamic arrays (FILTER, SORT, UNIQUE) | Similar (with ARRAYFORMULA) | Tie |
| **XLOOKUP** | Yes (native) | Yes (XLOOKUP added 2022) | Tie |
| **Data Tables (What-If)** | Yes | No | Excel |
| **Solver** | Yes (add-in) | Yes (add-on) | Tie |
| **Python integration** | Yes (Excel 365) | No (external only) | Excel |
| **Collaboration** | Good | Excellent | Google |
| **Cost** | $6.99–$22/user/month (M365) | Free (personal) / $6/user/month (Workspace) | Google |
| **File size** | Can handle large files | Performance degrades >5MB | Excel |
| **Add-in ecosystem** | Massive | Growing but smaller | Excel |
| **Learning curve** | Steeper (more features) | Gentler (simpler UI) | Google |
| **Version history** | Good (OneDrive/SharePoint) | Excellent (native) | Google |
| **Scripting** | VBA (proprietary) | Apps Script (open JS) | Depends |
| **Import from web** | Power Query (very powerful) | IMPORTHTML, IMPORTXML (basic) | Excel |

## When to Use Excel

- **Financial modeling** — What-If Analysis, Data Tables, Scenario Manager
- **Large datasets** — Power Query, Power Pivot, millions of rows
- **Complex automation** — VBA is more mature and capable than Apps Script
- **Enterprise environments** — Microsoft 365 integration, compliance features
- **Advanced analytics** — Power Pivot, DAX, Python integration
- **Offline-heavy work** — Full desktop application with no internet dependency
- **Professional reports** — Better charting, formatting, and print controls

## When to Use Google Sheets

- **Team collaboration** — Real-time editing is seamless and intuitive
- **Simple data tracking** — Inventory, to-do lists, basic budgets
- **Web-based workflows** — No installation, accessible from any device
- **Cost-sensitive** — Free for personal use
- **Integration with Google ecosystem** — Forms, BigQuery, Apps Script
- **Quick sharing** — Link sharing without file attachments
- **Education** — Easy distribution to students, no license needed

## Migration Tips

**Excel → Google Sheets:**
1. Upload .xlsx to Google Drive → Open with Sheets
2. Most formulas transfer, but check: VBA (won't work), complex conditional formatting
3. Data validation, PivotTables transfer but may behave differently
4. Named ranges transfer

**Google Sheets → Excel:**
1. File → Download → .xlsx
2. Apps Script won't transfer — rewrite in VBA
3. QUERY/IMPORTDATA functions won't work — use Power Query
4. Check conditional formatting rules

---

# Recommended Templates & Add-ins

## Built-in Templates (File → New)

| Template | Use Case |
|---|---|
| Monthly Budget | Personal finance tracking |
| Invoice | Professional billing |
| Project Planner | Gantt chart-style project management |
| Calendar (any year) | Printable monthly/annual calendar |
| To-Do List | Task management |
| Loan Amortization | Loan payment schedule |
| Sales Report | Pre-built dashboard with sample data |
| Gantt Chart | Project timeline visualization |
| Blood Pressure Tracker | Health monitoring |
| Meal Planner | Weekly meal planning with grocery list |

## Recommended Add-ins

| Add-in | Purpose | Cost |
|---|---|---|
| **Power Query** (built-in 2016+) | Data import and transformation | Free |
| **Power Pivot** (built-in) | Data modeling, DAX, large datasets | Free |
| **Solver** (built-in) | Optimization, linear programming | Free |
| **Analysis ToolPak** (built-in) | Statistical analysis, histograms | Free |
| **XLMiner Analysis ToolPak** | Advanced statistics (Google Sheets) | Free |
| **Fuzzy Lookup** (Microsoft) | Match similar but not identical data | Free |
| **Geography/Stocks** (built-in 365) | Live data types for geography and stocks | Free (365) |
| **Power BI Publisher** | Export ranges to Power BI | Free |
| **Kutools for Excel** | 300+ productivity tools | $49 |
| **Ablebits** | Data cleaning, merge, dedup | $59.95 |
| **Think-Cell** | Professional charting (consulting firms) | $300+/yr |
| **@RISK** | Monte Carlo simulation | $$$ |
| **Crystal Ball** | Forecasting and simulation | $$$ |
| **Python in Excel** (365) | Run Python directly in cells | Included in 365 |

## Free Online Resources

| Resource | URL | Description |
|---|---|---|
| Excel Easy | excel-easy.com | Beginner tutorials with examples |
| Chandoo | chandoo.org | Advanced formulas, dashboards, VBA |
| ExcelJet | exceljet.net | Formula reference, shortcuts, tips |
| Contextures | contextures.com | PivotTables, data validation, VBA |
| Mr. Excel | mrexcel.com | Forum for troubleshooting |
| Leila Gharani (YouTube) | youtube.com/@LeilaGharani | Professional Excel tutorials |
| ExcelIsFun (YouTube) | youtube.com/@excelisfun | 3,000+ videos, all levels |

---

# YouTube Video References

## Beginner Projects
- "How to Create a Personal Budget in Excel" — Leila Gharani
- "Excel Budget Template Tutorial" — ExcelIsFun
- "Invoice Template in Excel" — Leila Gharani
- "Excel for Beginners Full Course" — freeCodeCamp

## Intermediate Projects
- "Excel Sales Dashboard Tutorial" — MyOnlineTrainingHub
- "PivotTable Tutorial for Beginners" — Leila Gharani
- "Excel Inventory Management System" — ExcelIsFun
- "Employee Schedule in Excel" — MyOnlineTrainingHub
- "Data Cleaning in Excel — Full Tutorial" — Leila Gharani

## Advanced Projects
- "Financial Modeling in Excel — Full Tutorial" — Aswath Damodaran
- "Excel VBA Tutorial for Beginners" — Leila Gharani
- "Power Query Tutorial" — ExcelIsFun
- "Excel Dashboard Course" — MyOnlineTrainingHub
- "Automate Excel with VBA" — WiseOwlTutorials

## Certification Prep
- "MOS Excel Associate Exam Prep" — Certiport
- "MOS Excel Expert Exam Tips" — Leila Gharani
- "Excel Certification Study Guide" — Simon Sez IT

---

# Sources

1. Microsoft. "Excel documentation and training." support.microsoft.com
2. Microsoft. "Microsoft Office Specialist (MOS) certification." learn.microsoft.com
3. ExcelJet. "Excel formulas and functions reference." exceljet.net
4. Chandoo. "Excel tips, tutorials, and templates." chandoo.org
5. Gharani, Leila. "Excel tutorials and courses." leilagharani.com
6. ExcelIsFun. "Excel video tutorials." youtube.com/excelisfun
7. MyOnlineTrainingHub. "Excel dashboard and analysis tutorials." myonlinetraininghub.com
8. Contextures. "Excel tutorials and sample files." contextures.com
9. Mr. Excel. "Excel forum and resources." mrexcel.com
10. Google. "Google Sheets documentation." support.google.com/docs
11. Investopedia. "Financial modeling in Excel." investopedia.com
12. Corporate Finance Institute. "Excel for financial analysis." corporatefinanceinstitute.com

---

> **Next Steps:** After completing all 8 projects, you should be comfortable using Excel for any professional task. Consider pursuing MOS certification, exploring Power BI for advanced visualization, or learning Python integration for data science workflows.

> **Practice Dataset Sources:**
> - Kaggle (kaggle.com/datasets) — free datasets for practice
> - data.gov — U.S. government open data
> - Google Dataset Search (datasetsearch.research.google.com)
> - UCI Machine Learning Repository — classic datasets
