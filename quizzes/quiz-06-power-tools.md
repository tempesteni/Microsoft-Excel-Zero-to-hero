# Quiz 06 - Power Query, Power Pivot & DAX

**Module:** Power Query, Power Pivot, DAX, Data Modeling
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** What is the primary purpose of Power Query?
- A) Creating PivotTables
- B) Connecting to, transforming, and loading data from various sources
- C) Writing VBA macros
- D) Creating charts

**2.** In Power Query, you apply a "Remove Duplicates" step on column A. The result:
- A) Deletes duplicate rows from the original data
- B) Creates a new query result with duplicates removed - original data unchanged
- C) Highlights duplicates with color
- D) Moves duplicates to a new sheet

**3.** What is a data model in Excel?
- A) A 3D chart template
- B) A collection of tables with relationships defined between them
- C) A VBA script that processes data
- D) A Power Query connection

**4.** In DAX, what does `CALCULATE(SUM(Sales[Amount]), Sales[Region]="East")` do?
- A) Sums all sales amounts
- B) Sums sales amounts filtered to only the "East" region
- C) Returns an error
- D) Counts East region sales

**5.** You have two tables: Orders (OrderID, CustomerID, Amount) and Customers (CustomerID, Name, City). To create a relationship, you link:
- A) OrderID to CustomerID
- B) CustomerID in Orders to CustomerID in Customers
- C) Amount to Name
- D) Any column - Excel figures it out

**6.** What is the difference between "Close & Load" and "Close & Load To..." in Power Query?
- A) There's no difference
- B) "Close & Load" puts data in a new sheet; "Close & Load To..." lets you choose the destination
- C) "Close & Load To..." saves the file
- D) "Close & Load" creates a PivotTable

**7.** In DAX, which function creates a calculated column that ranks products by sales?
- A) `RANKX(ALL(Products), Products[Sales])`
- B) `RANK.EQ(Products[Sales])`
- C) `SUMX(Products, Products[Sales])`
- D) `INDEX(Products[Sales])`

**8.** Power Pivot differs from a regular PivotTable because:
- A) It can handle millions of rows without slowing down
- B) It can use DAX measures
- C) It works with the data model
- D) All of the above

**9.** You merge two queries in Power Query using a Left Outer join. What does this return?
- A) Only matching rows from both tables
- B) All rows from the left table, matching rows from the right (null if no match)
- C) All rows from both tables
- D) Only rows that don't match

**10.** A DAX measure using `ALL(Sales[Region])` does what within CALCULATE?
- A) Filters to all regions
- B) Removes the Region filter, showing totals across all regions
- C) Returns an error
- D) Filters to only blank regions

**11.** In Power Query's M language, what does `= Table.AddColumn(Source, "Tax", each [Amount] * 0.1)` do?
- A) Multiplies all amounts by 0.1
- B) Adds a new column "Tax" that calculates 10% of each row's Amount
- C) Replaces the Amount column
- D) Creates a separate table

**12.** Which DAX function is the time-intelligence equivalent of "year-to-date total"?
- A) `TOTALYTD()`
- B) `DATESYTD()`
- C) `YTD()`
- D) `SUMYTD()`

**13.** You append (union) two queries in Power Query. Both must have:
- A) The same number of rows
- B) The same column names and compatible data types
- C) The same number of columns only
- D) Identical data

**14.** What does "Edit Queries" in a PivotTable connected to Power Query let you do?
- A) Modify the data transformation steps
- B) Change the PivotTable layout
- C) Write SQL queries
- D) Nothing - it's read-only

**15.** In DAX, `SUMX` differs from `SUM` because:
- A) SUMX is faster
- B) SUMX evaluates an expression row by row before summing
- C) SUMX only works in calculated columns
- D) There's no difference

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

**16.** Write a DAX measure that calculates the total sales amount: `Total Sales = ...`

> **Your answer:** `_________________________________`

**17.** Write a DAX measure that calculates the percentage of total sales for the current filter context: `Sales % = ...`

> **Your answer:** `_________________________________`

**18.** Describe the Power Query steps to split a column containing "Last, First" into two separate columns.

> **Your answer:** `_________________________________`

**19.** Write a DAX calculated column formula that flags orders as "High" if Amount > 1000, otherwise "Low".

> **Your answer:** `_________________________________`

**20.** Describe how to create a date table in Power Query that spans from Jan 1, 2023 to Dec 31, 2025.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | Power Query is Excel's ETL tool for data import and transformation. |
| 2 | **B** | Power Query is non-destructive - it creates a new output from transformation steps. |
| 3 | **B** | The data model stores tables and their relationships for cross-table analysis. |
| 4 | **B** | CALCULATE modifies the filter context - here it filters to Region="East". |
| 5 | **B** | Relationships require matching key columns between tables. |
| 6 | **B** | "Close & Load To..." gives options: Table, PivotTable, Connection Only, etc. |
| 7 | **A** | RANKX iterates over a table and ranks based on an expression. |
| 8 | **D** | Power Pivot leverages the data model, DAX, and handles large datasets efficiently. |
| 9 | **B** | Left Outer keeps all left rows; right rows are null when no match exists. |
| 10 | **B** | ALL removes filters - used to calculate totals ignoring the current filter. |
| 11 | **B** | Table.AddColumn adds a new computed column to the table. |
| 12 | **A** | TOTALYTD computes year-to-date totals with a date column reference. |
| 13 | **B** | Append requires compatible columns (names and types). |
| 14 | **A** | "Edit Queries" opens Power Query Editor to modify transformation steps. |
| 15 | **B** | SUMX iterates row by row evaluating an expression; SUM just adds a column. |
| 16 | `Total Sales = SUM(Sales[Amount])` | Simple aggregation measure. |
| 17 | `Sales Pct = DIVIDE(SUM(Sales[Amount]), CALCULATE(SUM(Sales[Amount]), ALL(Sales)))` | DIVIDE handles division by zero; ALL removes filters for the denominator. |
| 18 | Select the column → Transform tab → Split Column → By Delimiter → choose Comma → set "At each occurrence" → OK. This creates two columns. Rename as needed. | - |
| 19 | `Order Flag = IF(Sales[Amount] > 1000, "High", "Low")` | Calculated column with IF. |
| 20 | Home > New Source > Blank Query → In Advanced Editor, enter: `= List.Dates(#date(2023,1,1), 1096, #duration(1,0,0,0))` → Convert to Table → Add date-related columns (Year, Month, etc.) → Close & Load. | 1096 days covers 2023-01-01 to 2025-12-31. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 7
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 6
