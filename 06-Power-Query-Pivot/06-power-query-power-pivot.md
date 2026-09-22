# Module 6: Power Query & Power Pivot

> **Target Audience:** Intermediate to Advanced Excel users  
> **Prerequisites:** Modules 1-5 (Formulas, PivotTables, Advanced Functions)  
> **Excel Version:** Microsoft 365 / Excel 2019+ (Power Pivot requires Pro Plus or standalone license)  
> **Last Updated:** September 2026

---

## Table of Contents

### Part A: Power Query (Get & Transform)
1. [What is Power Query?](#1-what-is-power-query)
2. [Getting Data from Various Sources](#2-getting-data-from-various-sources)
3. [The Query Editor Interface](#3-the-query-editor-interface)
4. [Basic Transformations](#4-basic-transformations)
5. [Split, Merge, and Extract Columns](#5-split-merge-and-extract-columns)
6. [Pivot and Unpivot Columns](#6-pivot-and-unpivot-columns)
7. [Merge Queries (SQL-Style JOINs)](#7-merge-queries)
8. [Append Queries (Stacking Tables)](#8-append-queries)
9. [Custom Columns with M Language](#9-custom-columns-with-m-language)
10. [Parameters](#10-parameters)
11. [Refreshing and Updating Queries](#11-refreshing-and-updating-queries)

### Part B: Power Pivot
12. [What is Power Pivot?](#12-what-is-power-pivot)
13. [Data Model Concepts](#13-data-model-concepts)
14. [Relationships Between Tables](#14-relationships-between-tables)
15. [DAX Basics](#15-dax-basics)
16. [Calculated Columns vs Measures](#16-calculated-columns-vs-measures)
17. [Key Performance Indicators (KPIs)](#17-key-performance-indicators-kpis)
18. [Hierarchies](#18-hierarchies)

### Part C: Integration & Practice
19. [Connecting Power Query to Power Pivot](#19-connecting-power-query-to-power-pivot)
20. [Practice Exercises](#20-practice-exercises)
21. [Sources and Further Reading](#21-sources-and-further-reading)

---

# Part A: Power Query (Get & Transform)

---

## 1. What is Power Query?

### The ETL Process Explained

Power Query is Excel's built-in **ETL** (Extract, Transform, Load) tool. ETL is the standard process for preparing data for analysis:

| Stage | What It Does | Power Query Equivalent |
|-------|-------------|----------------------|
| **Extract** | Pull data from various sources | **Get Data** - connect to files, databases, web, APIs |
| **Transform** | Clean, reshape, and prepare data | **Query Editor** - remove columns, filter, split, merge, pivot |
| **Load** | Load the cleaned data into your workbook | **Close & Load** - output to worksheet table or Data Model |

### Why Power Query?

- **Repeatable:** Set up once, refresh with one click when source data changes
- **No formulas needed:** Transformations are recorded as steps, not cell formulas
- **Handles large data:** Can process millions of rows efficiently
- **Multiple sources:** Combine data from CSV, databases, web pages, folders, and more
- **Non-destructive:** Original data is never modified; transformations create a new view

### Where to Find Power Query

| Excel Version | Location |
|---------------|----------|
| Excel 365 / 2021 | **Data** tab → **Get & Transform Data** group |
| Excel 2016+ | **Data** tab → **Get & Transform Data** group |
| Excel 2010-2013 | Download the free **Power Query add-in** from Microsoft |

---

## 2. Getting Data from Various Sources

### 2.1 From CSV / Text Files

1. **Data** → **Get Data** → **From File** → **From Text/CSV**
2. Browse to your `.csv` or `.txt` file
3. Excel previews the data and auto-detects:
   - Delimiter (comma, tab, semicolon, etc.)
   - Data types
4. Click **Transform Data** (opens Query Editor) or **Load** (loads directly)

**Common scenarios:**
- Monthly sales exports from an ERP system
- Data dumps from reporting tools
- Log files

### 2.2 From Excel Workbooks

1. **Data** → **Get Data** → **From File** → **From Workbook**
2. Select the `.xlsx` file
3. Navigator shows all sheets and named ranges
4. Select which table(s) to import

### 2.3 From the Web

1. **Data** → **Get Data** → **From Other Sources** → **From Web**
2. Enter a URL
3. Power Query detects **HTML tables** on the page
4. Select the table you want from the Navigator

**Example:** Import a Wikipedia table of countries by GDP.

> ⚠️ **Note:** Web scraping depends on the page structure. Tables in standard HTML `<table>` tags work best.

### 2.4 From Databases

1. **Data** → **Get Data** → **From Database**
2. Choose your database type:
   - **SQL Server** - enter server name, optional database
   - **Access** - select `.accdb` file
   - **SQL Server Analysis Services** - for OLAP cubes
   - **Oracle, MySQL, PostgreSQL** - via connectors
3. Navigator shows available tables and views
4. Select and load

### 2.5 From a Folder (Combining Files)

This is one of Power Query's most powerful features:

1. **Data** → **Get Data** → **From File** → **From Folder**
2. Select a folder containing multiple files (e.g., 12 monthly CSV files)
3. Power Query shows all files in the folder
4. Click **Combine & Transform** to:
   - Automatically stack all files into one table
   - Apply consistent column mapping
   - Handle files with the same structure but different data

**Use case:** Combine 12 monthly sales CSV files into a single annual dataset.

---

## 3. The Query Editor Interface

When you click **Transform Data**, the **Power Query Editor** opens:

### Key Areas of the Interface

```
┌─────────────────────────────────────────────────────────┐
│  Ribbon (Home, Transform, Add Column, View)             │
├──────────┬──────────────────────┬───────────────────────┤
│          │                      │                       │
│ Queries  │   Preview Pane       │  Query Settings       │
│ (left    │   (data grid)        │  (Applied Steps)      │
│  panel)  │                      │                       │
│          │                      │                       │
├──────────┴──────────────────────┴───────────────────────┤
│  Formula Bar (M language expression for current step)   │
└─────────────────────────────────────────────────────────┘
```

| Area | Purpose |
|------|---------|
| **Queries Pane** (left) | Lists all queries in the workbook |
| **Preview Pane** (center) | Shows a preview of the data at the current step |
| **Query Settings** (right) | **Applied Steps** - lists every transformation step in order |
| **Formula Bar** | Shows the M language formula for the selected step |
| **Ribbon Tabs** | Home (common tasks), Transform (reshape), Add Column (new columns), View (display options) |

### The Applied Steps Panel

Every action you take in Power Query is recorded as a **step**. You can:
- **Click any step** to see the data state at that point
- **Delete a step** to undo that transformation
- **Reorder steps** by dragging (be careful - order matters!)
- **Rename steps** for clarity (right-click → Rename)

**Example Applied Steps:**
```
1. Source              → Connected to SalesData.csv
2. Promoted Headers    → First row became column headers
3. Changed Type        → Columns assigned data types
4. Filtered Rows       → Kept only rows where Region = "North"
5. Removed Columns     → Deleted unnecessary columns
6. Renamed Columns     → Renamed "amt" to "Amount"
```

---

## 4. Basic Transformations

### 4.1 Remove Columns

**Method 1:** Select columns → **Home** → **Remove Columns**
**Method 2:** Right-click column header → **Remove**
**Method 3:** Select columns to KEEP → **Home** → **Remove Other Columns**

> 💡 **Best Practice:** Use "Remove Other Columns" when you want to keep only specific columns - it's more resilient to source data changes.

### 4.2 Change Data Types

Click the **data type icon** (left of column header) to change:
- **ABC** → Text
- **123** → Whole Number
- **1.2** → Decimal Number
- **📅** → Date
- **🕐** → Time
- **TRUE/FALSE** → Boolean

Or: **Transform** → **Data Type** dropdown

**Auto-detection:** Power Query tries to detect types when you first load data, but always verify!

### 4.3 Filter Rows

Click the **filter dropdown** on any column header:

**For text columns:**
- Text Filters → Equals, Contains, Begins With, Ends With, etc.
- Select/deselect specific values from the list

**For number columns:**
- Number Filters → Greater Than, Less Than, Between, Top 10, etc.

**For date columns:**
- Date Filters → Before, After, Between, This Month, Last Year, etc.

**M Language example:**
```m
= Table.SelectRows(#"Previous Step", each [Region] = "North" and [Amount] > 1000)
```

### 4.4 Sort Data

- Click column header dropdown → **Sort Ascending** or **Sort Descending**
- **Transform** → **Sort** for multi-level sorting

### 4.5 Rename Columns

- **Double-click** the column header → type new name
- Right-click → **Rename**
- Select column → **Transform** → **Rename Column**

### 4.6 Replace Values

Right-click a column → **Replace Values**

**Example:** Replace "N/A" with null:
- Value to Find: `N/A`
- Replace With: *(leave empty)*

**Bulk replacements:** Use **Transform** → **Replace Values** for find-and-replace operations across the entire column.

---

## 5. Split, Merge, and Extract Columns

### 5.1 Split Columns

**Transform** → **Split Column**

| Option | Use Case |
|--------|----------|
| **By Delimiter** | Split "John-Smith" by "-" → "John" and "Smith" |
| **By Number of Characters** | Split at a fixed position (e.g., first 3 chars are a code) |
| **By Positions** | Split at specific character positions |
| **By Lowercase to Uppercase** | Split "excelPower" → "excel" and "Power" |
| **By Digit to Non-Digit** | Split "ABC123DEF" → "ABC", "123", "DEF" |

**Example - Split Full Name:**

| Before | After (Split by space) |
|--------|----------------------|
| John Smith | John | Smith |
| Mary Jane Watson | Mary | Jane | Watson |

**Delimiter options:** Comma, Semicolon, Space, Tab, or Custom

**Split into:** Rows (vertical) or Columns (horizontal)

### 5.2 Merge Columns

Select two or more columns → **Transform** → **Merge Columns**

- Choose a **separator** (space, comma, custom, or none)
- Enter a **new column name**

**Example:**

| FirstName | LastName | → Merged (space separator): |
|-----------|----------|---------------------------|
| John | Smith | John Smith |

### 5.3 Extract

**Transform** → **Extract** (or **Add Column** → **Extract**)

| Option | Description | Example |
|--------|-------------|---------|
| **Text Length** | Count characters | "Hello" → 5 |
| **First Characters** | Extract N chars from start | "ABCDEF" (3) → "ABC" |
| **Last Characters** | Extract N chars from end | "ABCDEF" (2) → "EF" |
| **Text Range** | Extract from position, N chars | "ABCDEF" (pos 2, len 3) → "BCD" |
| **Text Before Delimiter** | Like TEXTBEFORE function | "A-B-C" (before "-") → "A" |
| **Text After Delimiter** | Like TEXTAFTER function | "A-B-C" (after last "-") → "C" |
| **Text Between Delimiters** | Extract between two delimiters | "[A][B][C]" (between "[", "]") → "A", "B", "C" |

---

## 6. Pivot and Unpivot Columns

### 6.1 Unpivot Columns (Wide → Long)

**Use case:** Convert a cross-tab/table layout into a flat list.

**Before (pivoted):**

| Product | Q1 | Q2 | Q3 | Q4 |
|---------|-----|-----|-----|-----|
| Widget A | 100 | 120 | 110 | 130 |
| Widget B | 200 | 210 | 190 | 220 |

**Select Q1-Q4 columns** → **Transform** → **Unpivot Columns**

**After (unpivoted):**

| Product | Attribute | Value |
|---------|-----------|-------|
| Widget A | Q1 | 100 |
| Widget A | Q2 | 120 |
| Widget A | Q3 | 110 |
| Widget A | Q4 | 130 |
| Widget B | Q1 | 200 |
| Widget B | Q2 | 210 |
| Widget B | Q3 | 190 |
| Widget B | Q4 | 220 |

> 💡 Rename "Attribute" to "Quarter" and "Value" to "Sales" for clarity.

**Variants:**
- **Unpivot Columns** - unpivot only selected columns
- **Unpivot Other Columns** - unpivot everything except selected columns

### 6.2 Pivot Columns (Long → Wide)

**Use case:** Reverse of unpivot - convert a flat list into a cross-tab.

Select the column whose values should become headers → **Transform** → **Pivot Column**

**Settings:**
- **Values Column:** Which column provides the values
- **Aggregate Value Function:** Don't Aggregate, Count, Sum, Min, Max, Average, Median

**Example:** Pivot the unpivoted data back into quarterly columns.

---

## 7. Merge Queries

### What is a Merge?

A **Merge** combines two queries side-by-side (horizontally), similar to a **SQL JOIN**. You match rows from one table to another based on a common column (key).

### How to Merge

1. **Home** → **Merge Queries** (or **Merge Queries as New** for a new query)
2. Select the **first table** (top) and the **second table** (bottom)
3. Click the **matching column(s)** in each table
4. Choose the **Join Kind**
5. Click **OK**

### Join Types Explained

| Join Kind | SQL Equivalent | What It Returns |
|-----------|---------------|-----------------|
| **Left Outer** | LEFT JOIN | All rows from Table 1 + matching rows from Table 2 |
| **Right Outer** | RIGHT JOIN | All rows from Table 2 + matching rows from Table 1 |
| **Inner** | INNER JOIN | Only rows that match in BOTH tables |
| **Full Outer** | FULL OUTER JOIN | All rows from BOTH tables, matched where possible |
| **Left Anti** | NOT IN | Rows in Table 1 that have NO match in Table 2 |
| **Right Anti** | NOT IN | Rows in Table 2 that have NO match in Table 1 |

### Example: Merge Orders with Customer Details

**Table 1 - Orders:**

| OrderID | CustomerID | Amount |
|---------|------------|--------|
| 1001 | C01 | $500 |
| 1002 | C02 | $300 |
| 1003 | C05 | $450 |
| 1004 | C01 | $200 |

**Table 2 - Customers:**

| CustomerID | Name | City |
|------------|------|------|
| C01 | Alice | New York |
| C02 | Bob | Chicago |
| C03 | Carol | Boston |
| C04 | David | Denver |

**Left Outer Merge (Orders + Customer info):**

| OrderID | CustomerID | Amount | Name | City |
|---------|------------|--------|------|------|
| 1001 | C01 | $500 | Alice | New York |
| 1002 | C02 | $300 | Bob | Chicago |
| 1003 | C05 | $450 | null | null |
| 1004 | C01 | $200 | Alice | New York |

> Notice: Customer C05 has no match → null values. C03 and C04 (no orders) are excluded.

**Inner Merge:** Would exclude Order 1003 (no matching customer).

**Left Anti Merge:** Would show only Customer C05 from Orders (those without a customer record).

### Expanding Merged Columns

After merging, you get a **Table** column. Click the **expand icon** (↔) in the column header to:
- Select which columns to bring in
- Choose to keep the original column name as prefix

---

## 8. Append Queries

### What is an Append?

An **Append** stacks queries on top of each other (vertically), similar to a **SQL UNION**. It combines rows from multiple tables into one.

### How to Append

1. **Home** → **Append Queries**
2. Choose:
   - **Two tables** - append one table to another
   - **Three or more tables** - append multiple tables
3. Select the tables to append
4. Click **OK**

### Rules for Appending

| Scenario | Behavior |
|----------|----------|
| Same columns | Rows stack perfectly |
| Extra columns in one table | Columns are included; missing values become `null` |
| Different column names | Treated as separate columns (won't merge) |

> 💡 **Tip:** Rename columns to match before appending if the structure differs.

### Example: Combine Monthly Reports

**January:**

| Date | Product | Sales |
|------|---------|-------|
| 2024-01-05 | Widget A | $1,000 |
| 2024-01-12 | Widget B | $800 |

**February:**

| Date | Product | Sales |
|------|---------|-------|
| 2024-02-03 | Widget A | $1,200 |
| 2024-02-18 | Widget C | $600 |

**Appended Result:**

| Date | Product | Sales |
|------|---------|-------|
| 2024-01-05 | Widget A | $1,000 |
| 2024-01-12 | Widget B | $800 |
| 2024-02-03 | Widget A | $1,200 |
| 2024-02-18 | Widget C | $600 |

---

## 9. Custom Columns with M Language

### What is M Language?

Every Power Query transformation is written in **M language** (also called Power Query Formula Language). You can see it in the **Formula Bar** and the **Advanced Editor**.

### Adding a Custom Column

**Add Column** → **Custom Column**

**Example 1 - Simple Calculation:**

```m
= [Price] * [Quantity]
```

Column name: `Revenue`

**Example 2 - Conditional Logic:**

```m
if [Amount] > 1000 then "High"
else if [Amount] > 500 then "Medium"
else "Low"
```

Column name: `Priority`

**Example 3 - Text Manipulation:**

```m
Text.Upper(Text.Start([FirstName], 1)) & Text.Lower(Text.End([FirstName], Text.Length([FirstName]) - 1))
```

Capitalizes the first letter of a name.

### Common M Functions

| Category | Function | Example |
|----------|----------|---------|
| **Text** | `Text.Length` | `Text.Length([Name])` → character count |
| **Text** | `Text.StartsWith` | `Text.StartsWith([Code], "US")` → true/false |
| **Text** | `Text.Contains` | `Text.Contains([Email], "@gmail")` → true/false |
| **Text** | `Text.Replace` | `Text.Replace([Name], "Inc", "LLC")` |
| **Text** | `Text.Combine` | `Text.Combine({[First], [Last]}, " ")` |
| **Number** | `Number.Round` | `Number.Round([Amount], 2)` |
| **Number** | `Number.Abs` | `Number.Abs([Variance])` |
| **Date** | `Date.Year` | `Date.Year([OrderDate])` → 2024 |
| **Date** | `Date.Month` | `Date.Month([OrderDate])` → 6 |
| **Date** | `Date.DayOfWeekName` | `Date.DayOfWeekName([OrderDate])` → "Monday" |
| **Date** | `Date.AddDays` | `Date.AddDays([ShipDate], 7)` |
| **Logic** | `if ... then ... else` | Conditional expressions |
| **Logic** | `try ... otherwise` | Error handling |

### The Advanced Editor

**Home** → **Advanced Editor** shows the complete M code for a query:

```m
let
    Source = Csv.Document(File.Contents("C:\Sales.csv")),
    PromotedHeaders = Table.PromoteHeaders(Source),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Date", type date}, {"Amount", type number}
    }),
    FilteredRows = Table.SelectRows(ChangedTypes, each [Amount] > 0)
in
    FilteredRows
```

**Structure:**
- `let` block: Define each step as a variable
- `in` block: The final result (what gets loaded)

---

## 10. Parameters

### What are Parameters?

Parameters make your queries **flexible and reusable**. Instead of hardcoding values (like a file path, date range, or filter value), you use a parameter that can be changed from a single place.

### Creating a Parameter

1. **Home** → **Manage Parameters** → **New Parameter**
2. Set:
   - **Name:** e.g., `FilterYear`
   - **Type:** Text, Number, Date, Boolean, etc.
   - **Current Value:** e.g., `2024`
   - **Suggested Values:** List of values, query, or any value

### Using Parameters in Queries

**Example - Filter by Year:**

```m
= Table.SelectRows(Source, each Date.Year([OrderDate]) = FilterYear)
```

**Example - Dynamic File Path:**

```m
= Csv.Document(File.Contents("C:\Data\" & DataFile & ".csv"))
```

Where `DataFile` is a text parameter (e.g., "Q1_2024_Sales").

### Using Parameters with Dropdowns

1. Create a parameter with **Suggested Values** → **List of values**
2. The parameter shows as a dropdown in the Queries pane
3. Users can change the value without editing the query

---

## 11. Refreshing and Updating Queries

### Manual Refresh

- **Data** tab → **Refresh All** (or press **Ctrl+Alt+F5**)
- Right-click a specific query → **Refresh**

### Auto-Refresh on File Open

1. **Data** → **Queries & Connections**
2. Right-click a query → **Properties**
3. Check **"Refresh data when opening the file"**

### Scheduled Refresh (Power BI Service only)

For cloud-hosted workbooks via Power BI:
- Configure **Scheduled Refresh** in the Power BI service
- Set frequency (daily, hourly, etc.)
- Requires a **data gateway** for on-premises sources

### Refresh Considerations

| Issue | Solution |
|-------|----------|
| Source file moved | Update the file path in the Source step |
| Column names changed | Update the relevant transformation steps |
| New columns added | "Remove Other Columns" step may exclude them - update |
| Slow refresh | Consider **query folding** (pushes transformations to the source) |
| Credentials expired | **Data** → **Queries & Connections** → enter credentials |

### Data Load Settings

Right-click a query → **Properties**:
- **Connection Name:** Display name
- **Load to:** Table (worksheet), Connection Only (no output), or **Data Model** (for Power Pivot)
- **Include in Refresh:** Yes/No

---

# Part B: Power Pivot

---

## 12. What is Power Pivot?

### Power Pivot vs Regular PivotTable

| Feature | Regular PivotTable | Power Pivot |
|---------|-------------------|-------------|
| Data source | Single table/range | Multiple tables via Data Model |
| Relationships | Not supported | Yes - define table relationships |
| Row limit | ~1,048,576 rows (worksheet limit) | Millions of rows (in-memory engine) |
| Calculations | Standard formulas in cells | DAX (Data Analysis Expressions) |
| Calculated fields | Limited | Rich DAX measures and calculated columns |
| KPIs | Not available | Built-in KPI support |
| Hierarchies | Manual grouping | Named hierarchies (drill-down) |
| Memory | Worksheet memory | VertiPaq column-store engine (compressed) |

### Enabling Power Pivot

1. **File** → **Options** → **Add-ins**
2. Manage: **COM Add-ins** → **Go...**
3. Check **Microsoft Power Pivot for Excel**
4. Click **OK**

A new **Power Pivot** tab appears on the ribbon.

> ⚠️ **Note:** Power Pivot is available in Excel Professional Plus and standalone Excel licenses. It is NOT included in Excel Standard or Excel for Mac.

---

## 13. Data Model Concepts

### What is the Data Model?

The **Data Model** is an in-memory database inside your Excel workbook. It:
- Stores data in a **compressed columnar format** (VertiPaq engine)
- Supports **relationships** between tables
- Enables **DAX** calculations
- Can hold **millions of rows** without worksheet limits

### Adding Data to the Data Model

**Method 1: From Power Query**
- When loading a query, choose **"Only Create Connection"** and check **"Add this data to the Data Model"**

**Method 2: From a Table**
- Click any cell in an Excel table → **Power Pivot** tab → **Add to Data Model**

**Method 3: From External Data**
- **Data** → **Get Data** → connect to a source → load to **Data Model**

### Power Pivot Window

**Power Pivot** tab → **Manage** opens the Power Pivot window:

```
┌─────────────────────────────────────────────────┐
│  Ribbon (Home, Design, Advanced)                │
├──────────┬──────────────────────────────────────┤
│          │                                      │
│  Table   │   Data View (rows & columns)         │
│  Tabs    │   or                                 │
│ (bottom) │   Diagram View (relationship map)    │
│          │                                      │
├──────────┴──────────────────────────────────────┤
│  Calculation Area (measures for active table)   │
└─────────────────────────────────────────────────┘
```

---

## 14. Relationships Between Tables

### Why Relationships?

Instead of one massive flat table, you use a **star schema**:

```
         ┌──────────┐
         │ DimDate  │
         │ (DateKey)│
         └────┬─────┘
              │
┌──────────┐  │  ┌──────────────┐
│DimProduct├──┼──┤ FactSales    │
│(ProductID)│  │ │(DateKey,     │
└──────────┘  │  │ ProductID,   │
              │  │ CustomerID)  │
┌──────────┐  │  └──────────────┘
│DimCust.  ├──┘
│(CustID) │
└──────────┘
```

- **Fact table** (center): Contains measurable data - sales amounts, quantities
- **Dimension tables** (surrounding): Contain descriptive data - dates, products, customers

### Creating Relationships

**Method 1: Diagram View**
1. Open Power Pivot → **Home** → **Diagram View**
2. **Drag** a field from one table to the matching field in another table
3. The relationship line appears

**Method 2: Ribbon**
1. **Home** → **Create Relationship**
2. Select Table 1, Column, Table 2, Related Column

### Relationship Settings

| Setting | Description |
|---------|-------------|
| **Cardinality** | Many-to-One (most common), One-to-One, Many-to-Many |
| **Cross Filter Direction** | Single (filter flows one way) or Both (bidirectional) |
| **Active** | Only one active relationship between two tables (inactive ones used via DAX USERELATIONSHIP) |

### Example: Sales and Products

**FactSales table:**
| ProductID | DateKey | Amount |
|-----------|---------|--------|
| P01 | 20240101 | 500 |
| P02 | 20240101 | 300 |
| P01 | 20240102 | 450 |

**DimProduct table:**
| ProductID | ProductName | Category |
|-----------|-------------|----------|
| P01 | Widget A | Electronics |
| P02 | Widget B | Hardware |
| P03 | Widget C | Electronics |

**Relationship:** FactSales[ProductID] → DimProduct[ProductID] (Many-to-One)

Now you can create PivotTables that show Sales by Category - even though "Category" lives in DimProduct, not FactSales.

---

## 15. DAX Basics

### What is DAX?

**DAX (Data Analysis Expressions)** is the formula language for Power Pivot. It looks similar to Excel formulas but has key differences:

| Feature | Excel Formula | DAX |
|---------|--------------|-----|
| Operates on | Individual cells | Columns and tables |
| Context | Cell context | Row context + Filter context |
| Aggregation | SUM(A1:A10) | SUM(Table[Column]) |
| Time intelligence | Limited | Rich built-in functions |

### Core DAX Concepts

#### Filter Context
When you place fields on a PivotTable's rows/columns/filters, they create a **filter context** that determines which rows a measure sees.

#### Row Context
Exists when a formula is being evaluated **row by row** (calculated columns, iterators like SUMX).

### Essential DAX Functions

#### SUM - Simple Aggregation

```dax
Total Sales = SUM(FactSales[Amount])
```

#### SUMX - Iterator (Row-by-Row Calculation)

```dax
Revenue = SUMX(FactSales, FactSales[Price] * FactSales[Quantity])
```

SUMX iterates through each row, calculates `Price * Quantity`, then sums the results.

#### COUNTX - Count with Expression

```dax
Order Count = COUNTX(FactSales, FactSales[OrderID])
```

#### RELATED - Pull from Related Table

```dax
Category = RELATED(DimProduct[Category])
```

Used in **calculated columns** to bring in a value from a related table.

#### CALCULATE - The Most Important DAX Function

```dax
Electronics Sales = CALCULATE(
    SUM(FactSales[Amount]),
    DimProduct[Category] = "Electronics"
)
```

CALCULATE modifies the **filter context**:
1. Starts with the existing filter context
2. Applies additional filters (or overrides existing ones)
3. Evaluates the expression in the modified context

**Key uses:**
- **Override filters:** `CALCULATE(SUM(...), ALL(DimProduct))` - ignores product filters
- **Add filters:** `CALCULATE(SUM(...), DimDate[Year] = 2024)` - adds year filter
- **Time intelligence:** `CALCULATE(SUM(...), DATEADD(DimDate[Date], -1, YEAR))` - prior year

#### FILTER - Return Filtered Table

```dax
High Value Orders = CALCULATE(
    COUNTROWS(FactSales),
    FILTER(FactSales, FactSales[Amount] > 1000)
)
```

FILTER returns a table that is a subset of the original, used inside CALCULATE.

#### RELATEDTABLE - Count/Sum Related Rows

```dax
Product Sales Count = COUNTROWS(RELATEDTABLE(FactSales))
```

Counts the sales rows related to each product.

#### DIVIDE - Safe Division

```dax
Avg Order Value = DIVIDE(SUM(FactSales[Amount]), COUNTROWS(FactSales), 0)
```

DIVIDE handles division by zero gracefully (returns 0 instead of error).

### DAX Examples Table

| DAX Measure | Description |
|-------------|-------------|
| `Total = SUM(Sales[Amount])` | Total sales amount |
| `Avg = AVERAGE(Sales[Amount])` | Average order value |
| `Orders = COUNTROWS(Sales)` | Number of orders |
| `Distinct Products = DISTINCTCOUNT(Sales[ProductID])` | Unique products sold |
| `YoY Growth = DIVIDE([Total] - [Prior Year], [Prior Year])` | Year-over-year growth % |
| `Prior Year = CALCULATE([Total], DATEADD(DimDate[Date], -1, YEAR))` | Sales from prior year |
| `YTD Sales = TOTALYTD([Total], DimDate[Date])` | Year-to-date sales |
| `MTD Sales = TOTALMTD([Total], DimDate[Date])` | Month-to-date sales |

---

## 16. Calculated Columns vs Measures

### Calculated Columns

**Created in:** Power Pivot data view or as a column in the Data Model
**Calculated:** When data is loaded/refreshed (stored in the model)
**Context:** Row context (evaluates per row)

**Example - Profit calculated column:**

```dax
= [Revenue] - [Cost]
```

**When to use:**
- You need the value for **row-by-row** filtering, slicing, or sorting
- The result is used in **relationships** or **groupings**
- The calculation doesn't change based on the PivotTable layout

### Measures (Calculated Fields)

**Created in:** Power Pivot → Calculation Area, or PivotTable → New Measure
**Calculated:** At query time (when the PivotTable is rendered)
**Context:** Filter context (evaluates based on slicers, rows, columns)

**Example - Profit margin measure:**

```dax
Profit Margin = DIVIDE(
    SUM(FactSales[Revenue]) - SUM(FactSales[Cost]),
    SUM(FactSales[Revenue]),
    0
)
```

**When to use:**
- The result depends on **user selections** (slicers, filters)
- You need **aggregations** (sum, average, count, etc.)
- The calculation should **respond dynamically** to the PivotTable layout

### Comparison

| Aspect | Calculated Column | Measure |
|--------|------------------|---------|
| Stored in model | ✅ Yes (uses memory) | ❌ No (calculated on demand) |
| Row context | ✅ Yes | ❌ No |
| Filter context | ❌ No | ✅ Yes |
| Use in PivotTable rows/values | Rows, filters | Values area |
| Performance impact | Increases file size | No file size impact |
| Best for | Static per-row values | Dynamic aggregations |

> 💡 **Rule of Thumb:** Prefer **measures** unless you specifically need a calculated column.

---

## 17. Key Performance Indicators (KPIs)

### What are KPIs?

KPIs in Power Pivot compare an **actual value** (base measure) against a **target value** to show performance status (e.g., on track, needs attention, off track).

### Creating a KPI

1. Create a **base measure** (e.g., `Total Sales`)
2. Create a **target measure** (e.g., `Sales Target`)
3. In Power Pivot → **Home** → **KPIs** → **New KPI**
4. Configure:
   - **KPI base measure:** Total Sales
   - **Target:** Sales Target (or Absolute Value)
   - **Thresholds:** Define ranges for Red/Yellow/Green

### Example KPI Setup

**Base Measure:**
```dax
Total Sales = SUM(FactSales[Amount])
```

**Target Measure:**
```dax
Sales Target = SUM(FactSales[TargetAmount])
```

**KPI Thresholds:**
| Status | Threshold |
|--------|-----------|
| 🔴 Red (Bad) | Actual < 80% of Target |
| 🟡 Yellow (Caution) | 80% ≤ Actual < 100% of Target |
| 🟢 Green (Good) | Actual ≥ 100% of Target |

### Using KPIs in PivotTables

When you add a KPI to a PivotTable, you can display:
- **Value:** The actual measure result
- **Status:** Icon (traffic light, gauge, etc.)
- **Goal:** The target value

---

## 18. Hierarchies

### What are Hierarchies?

Hierarchies define **drill-down paths** in your data. For example:
- **Date Hierarchy:** Year → Quarter → Month → Day
- **Geography Hierarchy:** Country → State → City
- **Product Hierarchy:** Category → Subcategory → Product

### Creating a Hierarchy

**In Diagram View:**
1. Right-click a field → **Create Hierarchy**
2. Name it (e.g., "Date Hierarchy")
3. Drag related fields into the hierarchy in order:
   - Year (Level 1)
   - Quarter (Level 2)
   - Month (Level 3)
   - Day (Level 4)

### Using Hierarchies in PivotTables

Once created, the hierarchy appears as a single expandable field in the PivotTable field list. Users can:
- Click **drill-down arrows** (⊞) to go from Year → Quarter → Month
- Click **drill-up arrows** (⊟) to go back up
- Use **Expand to Entire Hierarchy** for multi-level analysis

### Auto Date Hierarchies

Excel can automatically create date hierarchies:
1. **Power Pivot** → **Diagram View**
2. Select a date column
3. **Design** → **Mark as Date Table** → specify the date column

Excel auto-generates: Year, Quarter, Month, Day hierarchies.

---

# Part C: Integration & Practice

---

## 19. Connecting Power Query to Power Pivot

### The Ideal Workflow

```
┌──────────────┐     ┌──────────────┐     ┌──────────────┐
│  Data        │     │  Power       │     │  Power       │
│  Sources     │────▶│  Query       │────▶│  Pivot       │
│  (CSV, DB,   │     │  (Clean &    │     │  (Data Model │
│   Web, etc.) │     │   Transform) │     │   + DAX)     │
└──────────────┘     └──────────────┘     └──────┬───────┘
                                                  │
                                                  ▼
                                          ┌──────────────┐
                                          │  PivotTables │
                                          │  & Charts    │
                                          │  (Dashboard) │
                                          └──────────────┘
```

### Step-by-Step Integration

**Step 1: Load data with Power Query**
1. **Data** → **Get Data** → connect to your sources
2. Transform and clean in the Query Editor
3. In **Close & Load** dropdown → **Close & Load To...**
4. Select **"Only Create Connection"**
5. Check **"Add this data to the Data Model"**
6. Click **Load**

**Step 2: Define relationships in Power Pivot**
1. **Power Pivot** tab → **Manage**
2. Switch to **Diagram View**
3. Drag fields to create relationships

**Step 3: Create DAX measures**
1. In the **Calculation Area** below the data
2. Write DAX measures for your analysis

**Step 4: Build PivotTables from the Data Model**
1. **Insert** → **PivotTable**
2. Select **"Use this workbook's Data Model"**
3. Build your PivotTable using fields from multiple tables

### Benefits of This Workflow

- **Single source of truth:** Power Query handles all data prep
- **Refreshable:** One click updates all data, transformations, and PivotTables
- **Scalable:** Handle millions of rows without worksheet limits
- **Maintainable:** Transformation steps are visible and editable
- **Shareable:** Data Model connections persist when sharing the workbook

---

## 20. Practice Exercises

### Exercise 1: Power Query Basics

**Scenario:** You receive a messy CSV file with employee data.

**Tasks:**
1. Import the CSV file using Power Query
2. Promote the first row to headers
3. Remove unnecessary columns (keep only: ID, Name, Department, Salary, HireDate)
4. Change data types (Salary = decimal, HireDate = date)
5. Filter to keep only active employees
6. Split the "Name" column into "FirstName" and "LastName"
7. Add a custom column for "Years of Service" (current year minus hire year)
8. Load the cleaned data to a worksheet

### Exercise 2: Merge and Append

**Scenario:** Combine data from multiple sources.

**Tasks:**
1. Import two tables: **Orders** (OrderID, CustomerID, Amount) and **Customers** (CustomerID, Name, City)
2. **Merge** Orders with Customers using a Left Outer join on CustomerID
3. Expand the merged columns to bring in Name and City
4. Import three monthly CSV files (January, February, March)
5. **Append** them into a single "Q1 Sales" query
6. Group by Product and calculate total sales per product

### Exercise 3: Unpivot and Reshape

**Scenario:** A report exports data in cross-tab format.

**Tasks:**
1. Import a table with months as columns (Jan, Feb, Mar...) and products as rows
2. **Unpivot** the month columns to create a flat list
3. Rename the "Attribute" column to "Month" and "Value" column to "Sales"
4. Split a combined "ProductCode-ProductName" column into two columns
5. Pivot the data back to cross-tab format for verification

### Exercise 4: M Language and Parameters

**Tasks:**
1. Create a parameter called `MinAmount` (type: Decimal Number, value: 100)
2. Write a custom column formula that categorizes amounts: "Low" (<MinAmount), "Medium" (MinAmount-500), "High" (>500)
3. Open the Advanced Editor and read the M code for your query
4. Add a custom step using M: `Table.SelectRows(#"Previous Step", each [Amount] > MinAmount)`
5. Create a parameter for a file path and use it in the Source step

### Exercise 5: Data Model and Relationships

**Scenario:** Build a star schema for sales analysis.

**Tasks:**
1. Create/load these tables into the Data Model:
   - **FactSales** (OrderID, DateKey, ProductID, CustomerID, Amount, Quantity)
   - **DimDate** (DateKey, Date, Year, Quarter, Month, DayOfWeek)
   - **DimProduct** (ProductID, ProductName, Category, SubCategory)
   - **DimCustomer** (CustomerID, CustomerName, City, State)
2. Create relationships between FactSales and each dimension table
3. Build a PivotTable showing **Total Sales by Category and Year**
4. Add a slicer for City
5. Create a hierarchy: Year → Quarter → Month

### Exercise 6: DAX Measures

**Tasks:**
1. Create a measure: `Total Sales = SUM(FactSales[Amount])`
2. Create a measure: `Order Count = COUNTROWS(FactSales)`
3. Create a measure: `Avg Order Value = DIVIDE([Total Sales], [Order Count], 0)`
4. Create a measure for **Prior Year Sales** using CALCULATE and DATEADD
5. Create a measure for **YoY Growth %** = (Current - Prior) / Prior
6. Create a **KPI** comparing Total Sales to a Sales Target
7. Create a measure using SUMX: `Revenue = SUMX(FactSales, FactSales[Amount] * FactSales[Quantity])`

### Exercise 7: End-to-End Project

**Scenario:** Build a complete Sales Analytics Dashboard.

**Tasks:**
1. **Power Query:** Import sales CSV files from a folder, clean and transform
2. **Power Query:** Merge with a product lookup table and a customer table
3. **Data Model:** Load all tables, create relationships
4. **DAX:** Create measures for Total Revenue, Profit, Profit Margin, YoY Growth
5. **PivotTable:** Build a summary with Category × Year, with drill-down hierarchies
6. **KPI:** Add a KPI showing performance against targets
7. **Dashboard:** Create PivotCharts connected to your PivotTable
8. **Refresh:** Add new data to the source folder, refresh all queries, verify updates

---

## 21. Sources and Further Reading

### Microsoft Official Documentation
- [Power Query Overview - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [Connect to Data Sources - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [Merge Queries - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [Power Pivot Overview - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [DAX Reference - Microsoft Support](https://learn.microsoft.com/en-us/dax/)
- [CALCULATE Function - Microsoft DAX Reference](https://learn.microsoft.com/en-us/dax/calculate-function-dax)
- [Data Model in Excel - Microsoft Support](https://support.microsoft.com/en-us/excel)

### Exceljet
- [Power Query Overview - Exceljet](https://exceljet.net/)
- [DAX Functions Reference - Exceljet](https://exceljet.net/)

### Community Resources
- [Power Query Tutorial - MyOnlineTrainingHub](https://www.myonlinetraininghub.com/power-query-overview)
- [DAX Guide - Curbal](https://www.curbal.com/)
- [Power Query Tips - Chandoo.org](https://chandoo.org/?s=power+query)
- [Power Pivot and DAX - Excel Off The Grid](https://exceloffthegrid.com/)
- [M Language Reference - Microsoft](https://learn.microsoft.com/en-us/powerquery-m/)

### YouTube Video References
- [Power Query Full Tutorial - Leila Gharani](https://www.youtube.com/watch?v=ohG6zAslhXM)
- [Power Query for Beginners - ExcelIsFun](https://www.youtube.com/watch?v=fHF6ZMPiFmI)
- [Merge Queries (JOIN) Explained - Curbal](https://www.youtube.com/watch?v=YBzFgt5qWOc)
- [Append Queries - Curbal](https://www.youtube.com/watch?v=kwtmE5sMf6k)
- [M Language Basics - Curbal](https://www.youtube.com/watch?v=0yCyxjPFGAM)
- [Power Pivot Full Tutorial - Leila Gharani](https://www.youtube.com/watch?v=Q3EiNBMPxHg)
- [DAX for Beginners - Curbal](https://www.youtube.com/watch?v=TMR2MBKDOuI)
- [CALCULATE Function Deep Dive - Curbal](https://www.youtube.com/watch?v=Jb2DxhfDCb8)
- [Data Model and Relationships - MyOnlineTrainingHub](https://www.youtube.com/watch?v=Yb-Sahv0K3g)
- [Power Query Parameters - Excel Off The Grid](https://www.youtube.com/watch?v=5D3K0CbOFoM)
- [DAX SUMX vs SUM - Curbal](https://www.youtube.com/watch?v=ciHIenRJaKI)
- [Power Query Folder Connection - ExcelIsFun](https://www.youtube.com/watch?v=dHxNdVJSPCE)

### Books
- *Power Query for Power BI and Excel* - Chris Webb
- *The Definitive Guide to DAX* - Alberto Ferrari & Marco Russo
- *Power Pivot and Power BI* - Rob Collie & Avichal Singh
- *M Is for (Data) Monkey* - Ken Puls & Miguel Escobar

---

> **Previous Module:** [Module 5 - Advanced Functions](../05-Advanced-Functions/05-advanced-functions.md)
