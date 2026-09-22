# Module 4: Data Analysis

> **Level:** Beginner to Intermediate | **Estimated Time:** 4–5 hours | **Prerequisites:** Modules 1–3 (Fundamentals, Formulas & Functions, Formatting & Visualization)

---

## Table of Contents

1. [Sorting Data](#1-sorting-data)
2. [Filtering Data](#2-filtering-data)
3. [Advanced Filtering](#3-advanced-filtering)
4. [Data Validation](#4-data-validation)
5. [What-If Analysis](#5-what-if-analysis)
6. [PivotTables (Detailed)](#6-pivottables-detailed)
7. [Subtotals](#7-subtotals)
8. [Data Consolidation](#8-data-consolidation)
9. [Practice Exercises](#9-practice-exercises)
10. [Video References](#10-video-references)
11. [Sources](#11-sources)

---

## Sample Dataset

Throughout this module, we'll use the following **Sales Data** table. Create this in a new workbook on a sheet named `SalesData` to follow along:

| OrderID | Date       | Region  | SalesRep  | Product     | Category    | Units | UnitPrice | Total   |
|---------|------------|---------|-----------|-------------|-------------|-------|-----------|---------|
| 1001    | 2025-01-15 | North   | Alice     | Widget A    | Electronics | 25    | 10.00     | 250.00  |
| 1002    | 2025-01-20 | South   | Bob       | Widget B    | Hardware    | 40    | 8.50      | 340.00  |
| 1003    | 2025-02-03 | East    | Charlie   | Gadget X    | Electronics | 15    | 25.00     | 375.00  |
| 1004    | 2025-02-14 | West    | Diana     | Gadget Y    | Software    | 60    | 15.00     | 900.00  |
| 1005    | 2025-03-01 | North   | Alice     | Widget A    | Electronics | 30    | 10.00     | 300.00  |
| 1006    | 2025-03-10 | South   | Eve       | Widget C    | Hardware    | 20    | 12.00     | 240.00  |
| 1007    | 2025-03-22 | East    | Charlie   | Widget B    | Hardware    | 55    | 8.50      | 467.50  |
| 1008    | 2025-04-05 | West    | Frank     | Gadget X    | Electronics | 35    | 25.00     | 875.00  |
| 1009    | 2025-04-18 | North   | Bob       | Gadget Y    | Software    | 45    | 15.00     | 675.00  |
| 1010    | 2025-05-02 | South   | Alice     | Widget A    | Electronics | 50    | 10.00     | 500.00  |
| 1011    | 2025-05-15 | East    | Diana     | Widget C    | Hardware    | 25    | 12.00     | 300.00  |
| 1012    | 2025-06-01 | West    | Eve       | Gadget Y    | Software    | 70    | 15.00     | 1050.00 |
| 1013    | 2025-06-20 | North   | Charlie   | Gadget X    | Electronics | 20    | 25.00     | 500.00  |
| 1014    | 2025-07-04 | South   | Frank     | Widget A    | Electronics | 65    | 10.00     | 650.00  |
| 1015    | 2025-07-18 | East    | Bob       | Widget B    | Hardware    | 30    | 8.50      | 255.00  |
| 1016    | 2025-08-01 | West    | Alice     | Gadget X    | Electronics | 40    | 25.00     | 1000.00 |
| 1017    | 2025-08-15 | North   | Diana     | Widget C    | Hardware    | 35    | 12.00     | 420.00  |
| 1018    | 2025-09-03 | South   | Eve       | Gadget Y    | Software    | 80    | 15.00     | 1200.00 |
| 1019    | 2025-09-20 | East    | Frank     | Widget A    | Electronics | 45    | 10.00     | 450.00  |
| 1020    | 2025-10-05 | West    | Charlie   | Widget B    | Hardware    | 25    | 8.50      | 212.50  |

**Tip:** Type this data into Excel starting at cell `A1`. Use `=G2*H2` in the Total column (cell I2) and copy down to auto-calculate totals.

---

## 1. Sorting Data

Sorting rearranges rows based on the values in one or more columns. It does **not** change your data - it reorders it.

### 1.1 Single-Column Sort

**Quick Method:**
1. Click any cell in the column you want to sort by (e.g., click on a cell in the `Total` column)
2. Go to **Data** tab → **Sort & Filter** group
3. Click **Sort A to Z** (ascending ↑) or **Sort Z to A** (descending ↓)

**Example:** Sort by Total (highest to lowest):
- Click any cell in column I (Total)
- Data → Sort Z to A
- Result: Order 1018 ($1,200.00) appears first, Order 1020 ($212.50) appears last

> **Keyboard Shortcuts:**
> - `Alt + D, S, S` - Quick ascending sort
> - `Alt + D, S, O` - Quick descending sort
> - `Ctrl + Shift + L` - Toggle AutoFilter on/off

### 1.2 Multi-Level Sort

When two rows have the same value in the primary sort column, a secondary sort breaks the tie.

**Steps:**
1. Click any cell in your data range
2. Go to **Data** tab → **Sort** (the larger Sort button, not the A→Z icon)
3. The Sort dialog box opens:
   - **Sort by:** Select `Region` - Sort On: `Cell Values` - Order: `A to Z`
   - Click **Add Level**
   - **Then by:** Select `Total` - Sort On: `Cell Values` - Order: `Largest to Smallest`
4. Click **OK**

**Result:** All East records are grouped together (sorted by Total descending within East), then North, South, West.

### 1.3 Custom Sort Order

For data that doesn't sort alphabetically (e.g., days of the week, months, priority levels):

**Steps:**
1. Data → Sort
2. For the Order dropdown, select **Custom List...**
3. Choose an existing list (e.g., `Jan, Feb, Mar...`) or type your own (e.g., `High, Medium, Low`)
4. Click **Add** → **OK** → **OK**

**Example Custom List for Regions:**
```
North, South, East, West
```

### 1.4 Sort by Color

If cells have font colors or fill colors (e.g., conditional formatting, manual highlights):

1. Data → Sort
2. Under **Sort On**, choose `Cell Color`, `Font Color`, or `Cell Icon`
3. Under **Order**, pick the color and choose `On Top` or `On Bottom`
4. Add levels for additional colors

**Use Case:** Move all red-highlighted (flagged) rows to the top of the list.

### 1.5 Sort Left to Right (By Column)

Normally Excel sorts by rows, but you can sort columns:

1. Data → Sort → **Options...**
2. Select **Sort left to right**
3. Now sort by Row number instead of Column name

### Sort Pitfalls

| Problem | Cause | Solution |
|---------|-------|----------|
| Headers sort into the data | "My data has headers" checkbox unchecked | Re-sort and check the box |
| Formulas return wrong values after sort | Relative references shifted | Use absolute references `$A$1` or structured table references |
| Mixed numbers and text in one column | Numbers stored as text | Use `Data → Text to Columns → Finish` to convert |

---

## 2. Filtering Data

Filtering **hides** rows that don't match your criteria - the data stays intact, but you only see what's relevant.

### 2.1 Enable AutoFilter

1. Click any cell inside your data range
2. Go to **Data** tab → click **Filter** (or press `Ctrl + Shift + L`)
3. Dropdown arrows (▼) appear in each header cell

### 2.2 Filter by Selection (Quick Filter)

**Steps:**
1. Right-click a cell with the value you want to filter by (e.g., a cell containing "North")
2. Choose **Filter** → **Filter by Selected Cell's Value**
3. Only rows where Region = "North" are shown; row numbers turn blue to indicate hidden rows

### 2.3 Text Filters

Click the dropdown arrow on the `Region` column:

1. Uncheck **(Select All)** to clear all
2. Check only **North** and **South**
3. Click **OK**

**Advanced Text Filters** (click the dropdown → Text Filters):
- **Equals...** - exact match (e.g., `East`)
- **Does Not Equal...** - exclude a value
- **Begins With...** - e.g., `W` → matches "West"
- **Ends With...** - e.g., `th` → matches "North", "South"
- **Contains...** - e.g., `ou` → matches "South"
- **Does Not Contain...** - exclude rows with a substring

### 2.4 Number Filters

Click the dropdown arrow on the `Total` column:

- **Equals / Does Not Equal**
- **Greater Than...** - e.g., `500` → shows only totals above $500
- **Less Than...**
- **Between...** - e.g., `300` to `700`
- **Top 10...** - shows the top/bottom N items or top/bottom N%
- **Above Average / Below Average**
- **Custom Filter...** - combine two conditions with AND/OR

**Example - Top 5 Sales:**
1. Click dropdown on `Total`
2. Number Filters → Top 10...
3. Set: `Top` `5` `Items`
4. Click **OK**

### 2.5 Date Filters

Click the dropdown on the `Date` column:

- Date filters offer a **calendar picker** for selecting dates
- **Dynamic filters:** Today, Yesterday, Tomorrow, This Week, Last Month, Next Quarter, This Year, Year to Date...
- **Between...** - specify a date range (e.g., 2025-03-01 to 2025-06-30)
- **Custom Filter...** - combine conditions (e.g., After 2025-03-01 AND Before 2025-09-01)

### 2.6 Search Filter

When you click any column dropdown, there's a **search box** at the top:

1. Click the dropdown on `SalesRep`
2. Type `Alice` in the search box
3. Only "Alice" is shown in the checkbox list - check it → **OK**

### 2.7 Clear Filters

- **Clear one column:** Click the column dropdown → **Clear Filter from [ColumnName]**
- **Clear all filters:** Data tab → **Clear** (the funnel icon with an X)
- **Remove filter arrows entirely:** Data tab → **Filter** (toggle off)

### 2.8 Filter with Wildcards

In the search box or custom filter, use wildcards:
- `*` = any number of characters: `W*` matches "Widget A", "Widget B", "Widget C"
- `?` = exactly one character: `Gadget ?` matches "Gadget X", "Gadget Y"
- `~` = escape the wildcard: `~*` searches for a literal asterisk

---

## 3. Advanced Filtering

Advanced Filtering lets you use a **criteria range** - a separate area of the sheet where you define complex conditions with AND/OR logic.

### 3.1 Setting Up a Criteria Range

**Rules for criteria ranges:**
- The criteria range **must** have the same column headers as your data
- Place it on the same sheet (or a different sheet) away from your data
- **Same row** = AND logic (both conditions must be true)
- **Different row** = OR logic (either condition can be true)

**Example - Set up in cells K1:M4:**

| Region | Category    | Total     |
|--------|-------------|-----------|
| North  | Electronics | >400      |
| South  |             | >600      |

This means:
- Row 2 (AND): Region="North" AND Category="Electronics" AND Total>400
- Row 3 (OR): Region="South" AND Total>600

### 3.2 Apply Advanced Filter

1. Go to **Data** tab → **Advanced** (in the Sort & Filter group)
2. **Action:** Choose one:
   - **Filter the list, in-place** - hides non-matching rows (like AutoFilter)
   - **Copy to another location** - outputs matching rows to a different area
3. **List range:** Select your data including headers (e.g., `A1:I21`)
4. **Criteria range:** Select your criteria including headers (e.g., `K1:M4`)
5. If copying to another location: set **Copy to** (e.g., `K10`)
6. Check **Unique records only** to remove duplicates
7. Click **OK**

### 3.3 Criteria Range Examples

**AND conditions (same row):**

| Product | Units |
|---------|-------|
| Widget A | >30 |

→ Shows only Widget A orders with more than 30 units

**OR conditions (different rows):**

| Region |
|--------|
| North  |
| South  |

→ Shows all North OR South records

**Wildcard in criteria:**

| Product |
|---------|
| Widget* |

→ Shows all products starting with "Widget"

---

## 4. Data Validation

Data Validation controls what users can enter in a cell. It prevents bad data before it enters your spreadsheet.

### 4.1 Access Data Validation

1. Select the cell(s) where you want to restrict input
2. Go to **Data** tab → **Data Validation** (the checkmark icon with a dropdown)
3. The Data Validation dialog has three tabs: **Settings**, **Input Message**, **Error Alert**

### 4.2 Validation Types

#### Whole Number
Restricts to whole numbers only.

**Example - Units column must be 1–1000:**
1. Select the Units column (e.g., G2:G100)
2. Data → Data Validation → Settings tab
3. Allow: `Whole number`
4. Data: `between`
5. Minimum: `1`
6. Maximum: `1000`

#### Decimal
Allows decimal numbers.

**Example - Unit Price must be 0.01 to 9999.99:**
1. Select the UnitPrice column (H2:H100)
2. Allow: `Decimal`
3. Data: `between`
4. Minimum: `0.01`
5. Maximum: `9999.99`

#### List (Drop-down)
Creates a dropdown menu in the cell.

**Example - Region dropdown:**
1. Select the Region column (C2:C100)
2. Allow: `List`
3. Source: `North,South,East,West`
4. **Or** Source: `=$K$1:$K$4` (a range containing the options)
5. Check **In-cell dropdown** (should be on by default)
6. Click **OK**

**Drop-down from a Named Range:**
1. First, create a named range: select cells K1:K4 → type `RegionList` in the Name Box → Enter
2. In Data Validation, set Source: `=RegionList`

> **Tip:** Use an Excel Table (`Ctrl+T`) for dynamic dropdown lists. When you add items to the table, the dropdown automatically expands.

#### Date
Restricts to valid dates within a range.

**Example - Date must be in 2025:**
1. Allow: `Date`
2. Data: `between`
3. Start date: `2025-01-01`
4. End date: `2025-12-31`

#### Time
Restricts to times.

**Example - Business hours only:**
1. Allow: `Time`
2. Data: `between`
3. Start: `9:00:00 AM`
4. End: `5:00:00 PM`

#### Text Length
Restricts by character count.

**Example - Product code must be exactly 5 characters:**
1. Allow: `Text length`
2. Data: `equal to`
3. Length: `5`

### 4.3 Custom Formula Validation

Use a formula for the most flexible validation.

**Example - Total column must equal Units × Unit Price:**
1. Select the Total column (I2:I100)
2. Allow: `Custom`
3. Formula: `=I2=G2*H2`

**Example - No duplicate Order IDs:**
1. Select the OrderID column (A2:A100)
2. Allow: `Custom`
3. Formula: `=COUNTIF($A:$A,A2)<=1`

**Example - Date must be a weekday (Mon–Fri):**
1. Allow: `Custom`
2. Formula: `=WEEKDAY(B2,2)<=5`

### 4.4 Input Message

This is a tooltip that appears when the user selects the cell.

1. Go to the **Input Message** tab
2. Check **Show input message when cell is selected**
3. Title: `Region`
4. Message: `Select a region from the dropdown: North, South, East, or West`

### 4.5 Error Alert

Controls what happens when invalid data is entered.

1. Go to the **Error Alert** tab
2. Check **Show error alert after invalid data is entered**
3. Style: Choose one:
   - **Stop** 🛑 - Prevents the entry entirely (default)
   - **Warning** ⚠️ - Warns but allows the user to proceed
   - **Information** ℹ️ - Notifies but accepts the entry
4. Title: `Invalid Entry`
5. Message: `Please enter a value between 1 and 1000.`

### 4.6 Circle Invalid Data

After setting up validation, you can highlight cells that already contain invalid data:

1. Go to **Data** tab → **Data Validation** (dropdown) → **Circle Invalid Data**
2. Red circles appear around cells that violate the validation rules
3. To clear: **Data Validation** dropdown → **Clear Validation Circles**

---

## 5. What-If Analysis

What-If Analysis tools let you experiment with different input values to see how they affect your results.

### 5.1 Goal Seek

Goal Seek works **backwards** from a desired result. It asks: "I know what answer I want - what input value gives me that answer?"

**Scenario:** You have a formula `Total = Units × 10`. You want Total to equal $750. How many units do you need?

**Setup:**
- Cell G2 = Units (e.g., currently 25)
- Cell H2 = Unit Price = `10`
- Cell I2 = formula `=G2*H2` (currently 250)

**Steps:**
1. Click on the cell with the formula you want to change (I2)
2. Go to **Data** tab → **What-If Analysis** → **Goal Seek**
3. **Set cell:** `I2` (the formula cell)
4. **To value:** `750` (your desired result)
5. **By changing cell:** `G2` (the input cell)
6. Click **OK**

Excel iterates and finds: G2 = 75 (i.e., you need 75 units). Click **OK** to keep the result or **Cancel** to revert.

### 5.2 Scenario Manager

Scenarios let you save and compare multiple sets of input values for the same model.

**Example - Best Case, Worst Case, Expected Case for a profit model:**

Suppose you have:
- B2: Units Sold (input)
- B3: Price per Unit (input)
- B4: Cost per Unit (input)
- B5: `=B2*(B3-B4)` - Profit formula

**Create Scenarios:**
1. Go to **Data** tab → **What-If Analysis** → **Scenario Manager**
2. Click **Add...**
3. Scenario name: `Worst Case`
4. Changing cells: `B2,B3,B4` (hold Ctrl to select multiple)
5. Click **OK** → enter values: `100`, `15`, `12` → click **Add**
6. Scenario name: `Best Case` → values: `500`, `25`, `8` → click **Add**
7. Scenario name: `Expected` → values: `300`, `20`, `10` → click **OK**

**View a Scenario:** Select it from the list → click **Show**. The sheet updates to display those input values and recalculates formulas.

**Create a Summary Report:**
1. In Scenario Manager, click **Summary...**
2. Choose **Scenario summary** (table format) or **Scenario PivotTable report**
3. **Result cells:** Select B5 (your profit formula)
4. Click **OK**

Excel creates a new sheet with a comparison table showing all scenarios side by side.

### 5.3 Data Tables

Data Tables show how changing **one or two input variables** affects a formula result, creating a matrix of outcomes.

#### One-Variable Data Table

**Scenario:** How does profit change at different unit quantities?

**Setup:**
- C1: `Profit Model` (label)
- C2: `=B2*(B3-B4)` (link to your profit formula, or duplicate the formula here)
- Column D (input column): List different unit values: 100, 200, 300, 400, 500
- Row E1: leave blank or label

| D        | E          |
|----------|------------|
|          | `=C2`      |
| 100      |            |
| 200      |            |
| 300      |            |
| 400      |            |
| 500      |            |

**Steps:**
1. Select the entire table range (D1:E6)
2. **Data** tab → **What-If Analysis** → **Data Table...**
3. **Column input cell:** `B2` (the cell that the column values replace)
4. Click **OK**

Excel fills in the profit for each unit quantity.

#### Two-Variable Data Table

**Scenario:** How does profit change with BOTH different unit quantities AND different prices?

**Setup (E1:H6):**

|          | 15       | 20       | 25       |
|----------|----------|----------|----------|
|          | `=C2`    |          |          |
| 100      |          |          |          |
| 200      |          |          |          |
| 300      |          |          |          |
| 400      |          |          |          |
- Top-left corner (E2) contains the formula: `=C2`
- Row 2 (F2:H2) contains different price values
- Column E (E3:E6) contains different unit quantities

**Steps:**
1. Select the entire table (E2:H6)
2. Data → What-If Analysis → Data Table...
3. **Row input cell:** `B3` (price)
4. **Column input cell:** `B2` (units)
5. Click **OK**

Excel fills in the profit matrix.

### 5.4 Solver Add-in

Solver finds the **optimal** value for a formula cell subject to constraints. It's like Goal Seek on steroids.

**Enable Solver:**
1. File → Options → Add-ins
2. At the bottom, Manage: `Excel Add-ins` → **Go...**
3. Check **Solver Add-in** → **OK**
4. Now find it under **Data** tab → **Solver**

**Example - Maximize Profit with Constraints:**

Setup:
- B2: Units of Product A (variable)
- B3: Units of Product B (variable)
- B4: Profit formula = `=B2*10 + B3*15`
- Constraint: Total units ≤ 500 → `B2 + B3 <= 500`
- Constraint: Product A units ≥ 50 → `B2 >= 50`

**Steps:**
1. Data → Solver
2. **Set Objective:** `B4`
3. **To:** `Max` (maximize profit)
4. **By Changing Variable Cells:** `B2,B3`
5. **Subject to Constraints:** Click **Add**
   - `B2+B3 <= 500`
   - `B2 >= 50`
   - `B3 >= 0`
6. Solving Method: `GRG Nonlinear` (or `LP Simplex` for linear problems)
7. Click **Solve**

Solver finds the optimal combination (e.g., B2=50, B3=450) that maximizes profit.

---

## 6. PivotTables (Detailed)

PivotTables are the **most powerful** data analysis tool in Excel. They summarize, group, and analyze large datasets without writing formulas.

### 6.1 Creating a PivotTable

**Steps:**
1. Click any cell in your data range (e.g., the SalesData table)
2. Go to **Insert** tab → **PivotTable**
3. In the dialog:
   - **Table/Range:** Excel auto-detects your data range (e.g., `SalesData!$A$1:$I$21`)
   - **Choose where to place the PivotTable:**
     - `New Worksheet` (recommended) - creates a new sheet
     - `Existing Worksheet` - place it on the current sheet at a specific cell
4. Click **OK**

A blank PivotTable appears on a new sheet with the **PivotTable Fields** pane on the right.

### 6.2 Understanding the Field Areas

The PivotTable Fields pane has four areas:

| Area | Purpose | Example |
|------|---------|---------|
| **Filters** | Creates a report-level filter (dropdown at top of PivotTable) | `Region` → filter to show only North |
| **Columns** | Creates column headers across the top | `Category` → Electronics, Hardware, Software as columns |
| **Rows** | Creates row labels down the left side | `SalesRep` → each rep as a row |
| **Values** | The data to calculate (aggregated) | `Sum of Total` → sum of sales for each intersection |

**Example - Build your first PivotTable:**
1. Drag `SalesRep` to the **Rows** area
2. Drag `Total` to the **Values** area
3. You now see: each sales rep with their total sales

**Add a second dimension:**
4. Drag `Category` to the **Columns** area
5. Now you see: Sales reps × Product Categories with subtotals

**Add a filter:**
6. Drag `Region` to the **Filters** area
7. Use the dropdown at the top to filter by region

### 6.3 Value Field Settings

Click on the value in the Values area (e.g., "Sum of Total") → **Value Field Settings...**

#### Summarize By (Calculation Type)

| Function | What It Does | Example |
|----------|-------------|---------|
| **Sum** | Adds all values (default for numbers) | Total sales |
| **Count** | Counts cells with data | Number of orders |
| **Average** | Calculates mean | Average order value |
| **Max** | Largest value | Highest single sale |
| **Min** | Smallest value | Lowest single sale |
| **Product** | Multiplies all values | Compound growth |
| **StdDev / StdDevp** | Standard deviation | Sales volatility |
| **Var / Varp** | Variance | Data spread |

**Steps to change:**
1. Right-click any value in the PivotTable
2. **Summarize Values By** → choose `Average` (or any function)
3. Or: Value Field Settings → Summarize By tab → select function

### 6.4 Show Values As (Display Type)

This powerful feature shows values as **percentages, rankings, running totals**, and more - without extra formulas.

**Access:** Right-click a value → **Show Values As** → choose:

| Display Type | What It Shows | Example |
|-------------|---------------|---------|
| **% of Grand Total** | Each cell as % of the grand total | Alice's Widget A = 3.1% of all sales |
| **% of Column Total** | % within each column | Alice's % of Electronics column |
| **% of Row Total** | % within each row | Alice's % of her own total |
| **% of Parent Row Total** | % of parent row item | In a grouped row, % of the group |
| **% of Parent Column Total** | % of parent column item | % of parent column group |
| **Running Total In** | Cumulative sum across rows/columns | Year-to-date running total |
| **Rank Smallest to Largest** | Rank from lowest to highest | Rank reps by sales |
| **Rank Largest to Smallest** | Rank from highest to lowest | Top salesperson |
| **Difference From** | Difference vs. a base item | vs. previous month, vs. a specific region |
| **% Difference From** | % change vs. a base item | Month-over-month % change |
| **Index** | Relative importance of each value | Weighted comparison |

**Example - Show each rep's sales as % of Grand Total:**
1. Add `SalesRep` to Rows, `Total` to Values
2. Right-click any value cell → **Show Values As** → `% of Grand Total`
3. Alice now shows `13.2%` (her share of all sales)

**Example - Running Total by Date:**
1. Add `Date` (grouped by month) to Rows, `Total` to Values
2. Right-click → **Show Values As** → **Running Total In** → select the Date field
3. Each month now shows cumulative sales from January onward

### 6.5 Grouping

Grouping collapses detailed data into meaningful categories.

#### Group Dates

1. Place a date field in the Rows area
2. Right-click any date in the PivotTable
3. Select **Group...**
4. In the Grouping dialog, select the levels:
   - `Months`, `Quarters`, `Years` - can select multiple
5. Click **OK**

**Result:** Instead of 20 individual dates, you see:
```
2025
  Q1
    Jan          590.00
    Feb          1,275.00
    Mar          1,007.50
  Q2
    Apr          1,550.00
    ...
```

#### Group Numbers

If you have numeric data (e.g., order totals), you can create ranges:

1. Place `Total` in Rows
2. Right-click → **Group...**
3. Set: Starting at `0`, Ending at `1500`, By `250`
4. Click **OK**

**Result:** Groups like `0-250`, `250-500`, `500-750`, etc.

#### Ungroup

Right-click a grouped item → **Ungroup** (or select and press `Alt + Shift + Left Arrow`)

### 6.6 Calculated Fields

A **Calculated Field** creates a new field using a formula based on existing fields.

**Example - Create a "Commission" field (5% of Total):**
1. Click anywhere in the PivotTable
2. Go to **PivotTable Analyze** tab (or **Analyze** tab) → **Fields, Items & Sets** → **Calculated Field...**
3. Name: `Commission`
4. Formula: `=Total * 0.05`
5. Click **Add** → **OK**

A new "Commission" column appears in your PivotTable, calculated for each row/column intersection.

### 6.7 Calculated Items

A **Calculated Item** creates a new item within an existing field.

**Example - Create a "Coastal" item combining East and West:**
1. Click a cell in the `Region` field in the PivotTable
2. PivotTable Analyze → Fields, Items & Sets → **Calculated Item...**
3. Name: `Coastal`
4. Formula: `=East + West`
5. Click **Add** → **OK**

> **Note:** Calculated Items can cause unexpected results with subtotals. Use with caution.

### 6.8 PivotTable Design Tab

When you click inside a PivotTable, the **PivotTable Design** tab (or **Design** tab) appears:

#### Report Layout
- **Compact Form** (default): All row fields stacked in one column
- **Outline Form**: Each row field in its own column, subtotals at top
- **Tabular Form**: Each row field in its own column, subtotals at bottom

#### Subtotals
- **Do Not Show Subtotals**
- **Show All Subtotals at Top of Group**
- **Show All Subtotals at Bottom of Group**

#### Grand Totals
- **Off for Rows and Columns**
- **On for Rows and Columns** (default)
- **On for Rows Only**
- **On for Columns Only**

#### Report Filters
- Show in **Rows** (stacked vertically) or **Columns** (side by side)

#### Banded Rows / Columns
- Alternate shading for readability

#### PivotTable Styles
- Choose from 85+ pre-built styles in the Styles gallery
- Options: Banded rows, banded columns, first/last column headers

### 6.9 Refreshing PivotTables

PivotTables **do not automatically update** when source data changes.

**Manual Refresh:**
- Right-click the PivotTable → **Refresh**
- Or: **PivotTable Analyze** tab → **Refresh** (or press `Alt + F5`)
- Refresh all PivotTables in the workbook: `Ctrl + Alt + F5`

**Auto-Refresh on File Open:**
1. Right-click PivotTable → **PivotTable Options...**
2. Go to the **Data** tab
3. Check **Refresh data when opening the file**
4. Click **OK**

**Dynamic Source Data (Best Practice):**
Convert your source data to an **Excel Table** (`Ctrl+T`) before creating the PivotTable. When you add rows to the table, the PivotTable automatically includes them after refresh.

### 6.10 GETPIVOTDATA Function

When you reference a PivotTable cell in a formula, Excel uses `GETPIVOTDATA`:

```
=GETPIVOTDATA("Total",$A$3,"SalesRep","Alice","Category","Electronics")
```

This extracts Alice's Electronics total from the PivotTable.

**Arguments:**
- `"Total"` - the data field name
- `$A$3` - reference to any cell in the PivotTable
- `"SalesRep","Alice"` - field/item pair (filter)
- `"Category","Electronics"` - another field/item pair

**Toggle GETPIVOTDATA:**
If you prefer regular cell references instead:
- **PivotTable Analyze** → **Options** dropdown → uncheck **Generate GetPivotData**

### 6.11 PivotCharts

A PivotChart is a chart linked to a PivotTable - it updates automatically when the PivotTable changes.

**Create a PivotChart:**
1. Click inside your PivotTable
2. **PivotTable Analyze** tab → **PivotChart**
3. Choose a chart type (Column, Bar, Line, Pie, etc.)
4. Click **OK**

**PivotChart Features:**
- **Field buttons** on the chart let you filter directly on the chart
- Right-click field buttons → **Hide All Field Buttons** for a cleaner look
- The chart filters when you filter the PivotTable (and vice versa)

**Best Chart Types for PivotChart:**
| Data Type | Recommended Chart |
|-----------|-------------------|
| Sales by region | Clustered Column |
| Trend over time | Line |
| Category breakdown | Pie or Donut |
| Comparison + trend | Combo (Column + Line) |
| Ranking | Bar (horizontal) |

---

## 7. Subtotals

Subtotals automatically insert summary rows (SUM, AVERAGE, COUNT, etc.) at group boundaries in a sorted list.

### 7.1 Prerequisites

**⚠️ Important:** You MUST sort your data by the grouping column BEFORE applying subtotals. Excel inserts subtotal rows wherever the grouping column value changes.

### 7.2 Automatic Subtotals

**Example - Total sales by Region:**

1. Sort your data by `Region` (Data → Sort A to Z on Region column)
2. Click any cell in the data
3. Go to **Data** tab → **Subtotal** (in the Outline group)
4. In the Subtotal dialog:
   - **At each change in:** `Region`
   - **Use function:** `Sum`
   - **Add subtotal to:** check `Total` (and `Units` if desired)
   - Check **Summary below data** (default)
   - Check **Replace current subtotals** (if re-applying)
   - Check **Page break between groups** (optional - for printing)
5. Click **OK**

**Result:**
```
Alice    North    Widget A    25    10.00    250.00
Alice    North    Widget A    30    10.00    300.00
Bob      North    Gadget Y    45    15.00    675.00
Charlie  North    Gadget X    20    25.00    500.00
Diana    North    Widget C    35    12.00    420.00
                                        North Total   2145.00
...
```

### 7.3 Other Subtotal Functions

Besides `Sum`, you can use:
- **Count** - number of items per group
- **Average** - average per group
- **Max** - maximum per group
- **Min** - minimum per group
- **Product** - product of all values
- **StdDev / StdDevp** - standard deviation
- **Var / Varp** - variance

### 7.4 Nesting Subtotals (Multiple Levels)

You can add subtotals at multiple levels. For example, subtotal by Region, then by Category within each Region:

1. Sort by `Region` first, then by `Category` (multi-level sort)
2. Apply subtotals: At each change in `Region`, Sum of `Total` → OK
3. Apply subtotals again: At each change in `Category`, Sum of `Total`
4. **⚠️ UNCHECK** "Replace current subtotals" - this adds the second level on top of the first

### 7.5 Group and Ungroup (Outline)

After applying subtotals, Excel creates an **outline** with level buttons in the top-left corner:

- **Level 1:** Grand total only (most collapsed)
- **Level 2:** Subtotals + grand total
- **Level 3:** All data + subtotals + grand total (most expanded)

**Click the level numbers** (1, 2, 3) at the top-left to collapse/expand.

**Manual Grouping (without Subtotals):**
1. Select the rows you want to group
2. **Data** tab → **Group** (or `Alt + Shift + Right Arrow`)
3. To ungroup: select grouped rows → **Ungroup** (`Alt + Shift + Left Arrow`)

### 7.6 Remove Subtotals

1. Data → **Subtotal**
2. Click **Remove All** at the bottom-left of the dialog
3. All subtotal rows and the outline are removed

---

## 8. Data Consolidation

Consolidation combines data from multiple ranges, worksheets, or workbooks into a single summary.

### 8.1 Consolidate by Category

Use when data is organized differently across sources but has **matching labels** (e.g., same product names in different regional reports).

**Example - Consolidate quarterly sales:**

Imagine three sheets: `Q1_Sales`, `Q2_Sales`, `Q3_Sales`, each with:

| Product     | Sales |
|-------------|-------|
| Widget A    | 500   |
| Widget B    | 350   |
| Gadget X    | 750   |

**Steps:**
1. Go to a new summary sheet
2. Click where you want the consolidated data to start (e.g., A1)
3. **Data** tab → **Consolidate** (in the Data Tools group)
4. In the Consolidate dialog:
   - **Function:** `Sum` (or Average, Count, etc.)
   - **Reference:** Click the range picker, go to `Q1_Sales!$A$1:$B$4`, click **Add**
   - Repeat for Q2 and Q3
   - **Use labels in:** check `Top row` and `Left column`
5. Click **OK**

**Result:** Excel creates a merged summary:
| Product     | Sales |
|-------------|-------|
| Widget A    | 1650  |
| Widget B    | 1100  |
| Gadget X    | 2400  |

### 8.2 Consolidate by Position

Use when all source ranges have **identical layout** (same rows, same columns, same order).

**Steps:**
1. Same as above, but **do NOT check** "Top row" or "Left column"
2. Excel simply sums (or averages, etc.) the values in the same cell positions across all ranges

### 8.3 Creating Links to Source Data

In the Consolidate dialog, check **Create links to source data**.

- The consolidated result will be grouped/linked to the source data
- When source data changes, the consolidation updates automatically
- Creates an outline similar to Subtotals that you can expand/collapse

> **Note:** When linking is enabled, you cannot edit the consolidated cells directly - they are formulas.

### 8.4 Consolidation Tips

- **Different workbooks:** You can consolidate data from files that are not open - browse to the file path
- **Named ranges:** Use named ranges for cleaner references
- **Multiple functions:** You can run consolidation multiple times with different functions (Sum, Average, Count) and place them side by side
- **3D References:** An alternative to Consolidate - use formulas like `=SUM(Q1_Sales:Q3_Sales!B2)` to sum the same cell across multiple sheets

---

## 9. Practice Exercises

### Exercise 1: Sorting & Filtering (Beginner)

Using the sample SalesData from the beginning of this module:

1. Sort the data by `Region` (A→Z), then by `Total` (largest to smallest)
2. Filter to show only `Electronics` category with more than 50 units
3. Clear filters, then use AutoFilter to find the **Top 3** sales by Total
4. Sort by color: Highlight all orders over $800 in yellow, then sort yellow cells to the top

### Exercise 2: Advanced Filtering (Intermediate)

Using the SalesData:

1. Set up a criteria range that shows:
   - East region AND Electronics category
   - OR West region AND total > $800
2. Apply Advanced Filter and copy results to cell K15
3. Use a wildcard filter to show all "Gadget" products (begins with "Gadget")

### Exercise 3: Data Validation (Beginner)

1. Add a dropdown for the `Region` column using a list: North, South, East, West
2. Add validation to `Units` column: must be a whole number between 1 and 1000
3. Add an input message: "Enter number of units (1-1000)"
4. Add an error alert (Stop style): "Units must be between 1 and 1000!"
5. Create a formula-based validation: Order Date must not be in the future (`<=TODAY()`)

### Exercise 4: What-If Analysis (Intermediate)

Create a simple pricing model:
- A1: "Units", A2: 100 (input)
- B1: "Price", B2: 25.00 (input)
- C1: "Cost/Unit", C2: 15.00 (input)
- D1: "Profit", D2: `=A2*(B2-C2)` (formula)

1. Use **Goal Seek**: What price do you need to achieve $2,000 profit with 100 units?
2. Create **Scenarios**: Best Case (200 units, $30 price, $10 cost), Worst Case (50 units, $20 price, $18 cost), Base Case (100 units, $25 price, $15 cost)
3. Create a **Scenario Summary** report
4. Create a **1-Variable Data Table** showing profit at 50, 100, 150, 200, 250, 300 units
5. Create a **2-Variable Data Table** with units (rows: 50–300) and price (columns: $20–$35)

### Exercise 5: PivotTables (Intermediate–Advanced)

Using the SalesData:

1. Create a PivotTable showing **Total Sales by SalesRep**
2. Add `Category` to Columns - now see sales per rep per category
3. Change the value display to **% of Column Total**
4. Group the dates by **Quarters**
5. Create a **Calculated Field** called "AvgPrice" = Total / Units
6. Add a **Region** filter and show only North and South
7. Change the layout to **Tabular Form** with subtotals at bottom
8. Create a **PivotChart** (Clustered Column) showing sales by rep by category
9. Refresh the PivotTable after adding two new rows to the source data

### Exercise 6: Subtotals & Consolidation (Intermediate)

**Part A - Subtotals:**
1. Sort SalesData by `Region`, then by `Category`
2. Apply **nested subtotals**: Sum of Total at each Region change, then Sum of Total at each Category change
3. Collapse to Level 2 to see only subtotals
4. Remove all subtotals

**Part B - Consolidation:**
1. Create three sheets: `Store_A`, `Store_B`, `Store_C` each with the same products but different sales numbers
2. Use Consolidate (Sum, by category) to create a combined total on a `Summary` sheet
3. Try consolidating with "Create links to source data" checked
4. Edit a value in Store_A and verify the Summary updates

---

## 10. Video References

### Sorting & Filtering
- Excel Sorting and Filtering - Leila Gharani
  https://www.youtube.com/watch?v=ylI8uHb_gmg
- Advanced Filter in Excel - Leila Gharani
  https://www.youtube.com/watch?v=HK1UOB2bPBI

### Data Validation
- Excel Data Validation Drop-down List - Leila Gharani
  https://www.youtube.com/watch?v=7moKWkCKbzE
- Dynamic Drop-down List with Data Validation - ExcelJet
  https://www.youtube.com/watch?v=5sD2CBLSMkU

### What-If Analysis
- Goal Seek, Scenario Manager, Data Tables - Leila Gharani
  https://www.youtube.com/watch?v=SfGkEGvq4JE
- Excel Solver Explained - Leila Gharani
  https://www.youtube.com/watch?v=kOO_PQ7VSzQ

### PivotTables
- PivotTable Tutorial - Leila Gharani
  https://www.youtube.com/watch?v=qu-Acm06sVY
- PivotTable Master Class - Chandoo
  https://www.youtube.com/watch?v=S2PFMBons0A
- PivotTable Grouping (Dates, Numbers) - Leila Gharani
  https://www.youtube.com/watch?v=bWpGPYctav4
- PivotTable Show Values As - Leila Gharani
  https://www.youtube.com/watch?v=Cz0VBjApM0M
- Calculated Fields in PivotTables - MyOnlineTrainingHub
  https://www.youtube.com/watch?v=nOkTIL8K3GM

### Subtotals & Consolidation
- Excel Subtotals - MyOnlineTrainingHub
  https://www.youtube.com/watch?v=3CJzmYUMgE0
- Consolidate Data from Multiple Sheets - Leila Gharani
  https://www.youtube.com/watch?v=HK1UOB2bPBI

---

## 11. Sources

### Microsoft Support
- Sort data in a range or table
  https://support.microsoft.com/en-us/excel
- Filter data in a range or table
  https://support.microsoft.com/en-us/excel
- Advanced Filter
  https://support.microsoft.com/en-us/excel
- Apply data validation to cells
  https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249
- Introduction to What-If Analysis
  https://support.microsoft.com/en-us/excel
- Goal Seek
  https://support.microsoft.com/en-us/excel
- Create a PivotTable
  https://support.microsoft.com/en-us/excel
- Group or Ungroup data in a PivotTable
  https://support.microsoft.com/en-us/excel
- Insert subtotals in a list of data
  https://support.microsoft.com/en-us/excel
- Consolidate data from multiple ranges
  https://support.microsoft.com/en-us/excel

### ExcelJet.net
- Excel PivotTable - Detailed Guide
  https://exceljet.net/
- Excel Sorting Guide
  https://exceljet.net/
- Excel Filtering Guide
  https://exceljet.net/
- Data Validation Guide
  https://exceljet.net/
- Goal Seek
  https://exceljet.net/

### Chandoo.org
- Pivot Tables - Comprehensive Guide
  https://chandoo.org/?s=excelpivot-tables/
- Sorting and Filtering Tips
  https://chandoo.org/?s=sorting+filtering
- Data Validation Tricks
  https://chandoo.org/?s=data+validation

### Contextures.com
- Excel Pivot Table Tutorial
  https://www.contextures.com/
- Excel Data Validation Examples
  https://www.contextures.com/xlDataVal01.html
- Excel Subtotals
  https://www.contextures.com/
- Advanced Filter Examples
  https://www.contextures.com/xladvfilter01.html

### Corporate Finance Institute (CFI)
- PivotTable Guide
  https://corporatefinanceinstitute.com/resources/excel/
- What-If Analysis
  https://corporatefinanceinstitute.com/resources/excel/what-if-analysis/
- Data Tables in Excel
  https://corporatefinanceinstitute.com/resources/excel/data-tables/

---

> **Next Module:** [Module 5 - Advanced Functions](../05-Advanced-Functions/05-advanced-functions.md) covers XLOOKUP, INDEX/MATCH, dynamic arrays, LAMBDA, and more.
