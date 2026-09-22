# Module 5: Advanced Functions in Microsoft Excel

> **Target Audience:** Intermediate to Advanced Excel users  
> **Prerequisites:** Modules 1-4 (Basic formulas, cell references, named ranges, IF/SUMIFS/COUNTIFS)  
> **Excel Version:** Microsoft 365 / Excel 2021 (some functions require Excel 365)  
> **Last Updated:** September 2026

---

## Table of Contents

1. [XLOOKUP Function](#1-xlookup-function)
2. [XMATCH Function](#2-xmatch-function)
3. [INDEX/MATCH Combinations](#3-indexmatch-combinations)
4. [OFFSET Function and Dynamic Ranges](#4-offset-function-and-dynamic-ranges)
5. [INDIRECT Function and Dynamic References](#5-indirect-function-and-dynamic-references)
6. [CHOOSE Function](#6-choose-function)
7. [Dynamic Arrays (Excel 365/2021)](#7-dynamic-arrays)
8. [LET Function for Formula Readability](#8-let-function)
9. [LAMBDA Function (Custom Functions)](#9-lambda-function)
10. [Lambda Helper Functions](#10-lambda-helper-functions)
11. [New Text Functions](#11-new-text-functions)
12. [VSTACK and HSTACK](#12-vstack-and-hstack)
13. [TOCOL and TOROW](#13-tocol-and-torow)
14. [WRAPCOLS and WRAPROWS](#14-wrapcols-and-wraprows)
15. [Practice Exercises](#15-practice-exercises)
16. [Sources and Further Reading](#16-sources-and-further-reading)

---

## 1. XLOOKUP Function

### What is XLOOKUP?

XLOOKUP is Excel's modern, all-in-one lookup function introduced in Microsoft 365. It searches a range or array for a value and returns the corresponding value from another range. It replaces VLOOKUP, HLOOKUP, and the legacy LOOKUP function with a cleaner, more powerful syntax.

### Syntax

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

| Argument | Required? | Description |
|----------|-----------|-------------|
| `lookup_value` | ✅ Yes | The value to search for |
| `lookup_array` | ✅ Yes | The range or array to search within |
| `return_array` | ✅ Yes | The range or array to return values from |
| `if_not_found` | ❌ Optional | Value to return if no match is found (default: #N/A error) |
| `match_mode` | ❌ Optional | 0 = exact (default), -1 = exact or next smaller, 1 = exact or next larger, 2 = wildcard, 3 = regex |
| `search_mode` | ❌ Optional | 1 = first-to-last (default), -1 = last-to-first, 2 = binary ascending, -2 = binary descending |

### Example 1: Basic Exact Match

**Sample Data - Employee Table:**

| A (Emp ID) | B (Name) | C (Department) | D (Salary) |
|------------|----------|----------------|------------|
| E001 | Alice Johnson | Engineering | $95,000 |
| E002 | Bob Smith | Marketing | $72,000 |
| E003 | Carol Davis | Engineering | $88,000 |
| E004 | David Lee | Sales | $67,000 |
| E005 | Eve Wilson | Marketing | $78,000 |

**Formula:** Look up the salary for employee "E003"

```excel
=XLOOKUP("E003", A2:A6, D2:D6)
```

**Result:** `$88,000`

### Example 2: Custom "Not Found" Message

```excel
=XLOOKUP("E099", A2:A6, D2:D6, "Employee Not Found")
```

**Result:** `Employee Not Found` (instead of #N/A)

### Example 3: Approximate Match - Discount Tiers

**Sample Data - Discount Table:**

| A (Min Qty) | B (Discount %) |
|-------------|----------------|
| 1 | 0% |
| 10 | 5% |
| 50 | 10% |
| 100 | 15% |
| 500 | 20% |

```excel
=XLOOKUP(75, A2:A6, B2:B6, , -1)
```

- `match_mode = -1` means "exact match or next smaller value"
- 75 falls between 50 and 100, so it returns the discount for 50 → **10%**

### Example 4: Return Multiple Values (Entire Row)

XLOOKUP can return an entire row of data at once:

```excel
=XLOOKUP("E003", A2:A6, B2:D6)
```

**Result:** Spills across 3 cells: `Carol Davis | Engineering | $88,000`

### Example 5: Nested XLOOKUP (Two-Way Lookup)

Look up a value by matching both a row header AND a column header:

**Sample Data - Sales by Region and Quarter:**

| | E (Q1) | F (Q2) | G (Q3) | H (Q4) |
|---|--------|--------|--------|--------|
| **North** | 12000 | 15000 | 13000 | 18000 |
| **South** | 9000 | 11000 | 10000 | 14000 |
| **East** | 14000 | 16000 | 15000 | 19000 |
| **West** | 8000 | 10000 | 9500 | 12000 |

```excel
=XLOOKUP("South", A2:A5, XLOOKUP("Q3", E1:H1, E2:H5))
```

**How it works:**
1. Inner XLOOKUP: Finds "Q3" in the header row → returns column G (the Q3 data for all regions)
2. Outer XLOOKUP: Finds "South" in that column → returns **10,000**

### Example 6: Reverse Search (Find Last Match)

```excel
=XLOOKUP("Engineering", C2:C6, D2:D6, , 0, -1)
```

- `search_mode = -1` searches from **last to first**
- Returns the salary of the **last** employee in Engineering

### Example 7: Wildcard Match

```excel
=XLOOKUP("*Wilson*", B2:B6, D2:D6, , 2)
```

- `match_mode = 2` enables wildcard matching with `*` and `?`
- Returns the salary of any name containing "Wilson"

### Advantages Over VLOOKUP

| Feature | VLOOKUP | XLOOKUP |
|---------|---------|---------|
| Look left | ❌ No | ✅ Yes |
| Return multiple columns | ❌ Single column | ✅ Entire rows |
| Default match | Approximate | Exact |
| Custom #N/A message | ❌ Requires IFERROR | ✅ Built-in |
| Reverse search | ❌ No | ✅ Yes |
| Column insert safe | ❌ Breaks on insert | ✅ References are ranges |

---

## 2. XMATCH Function

### What is XMATCH?

XMATCH is the modern replacement for the MATCH function. It returns the **relative position** of an item in an array. Unlike MATCH, it supports reverse search and binary search.

### Syntax

```
=XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
```

| Argument | Required? | Description |
|----------|-----------|-------------|
| `lookup_value` | ✅ Yes | The value to find |
| `lookup_array` | ✅ Yes | The range or array to search |
| `match_mode` | ❌ Optional | 0 = exact (default), -1 = next smaller, 1 = next larger, 2 = wildcard |
| `search_mode` | ❌ Optional | 1 = first-to-last (default), -1 = last-to-first, 2 = binary ascending, -2 = binary descending |

### Example 1: Find Position of a Value

```excel
=XMATCH("Carol Davis", B2:B6)
```

**Result:** `3` (Carol Davis is the 3rd name in the list)

### Example 2: Find Last Occurrence

```excel
=XMATCH("Engineering", C2:C6, 0, -1)
```

Returns the position of the **last** Engineering employee.

### Example 3: Approximate Match with XMATCH

```excel
=XMATCH(75, {1, 10, 50, 100, 500}, -1)
```

**Result:** `3` (75 falls closest to 50 from below, which is the 3rd item)

### Example 4: XMATCH + INDEX (Alternative to XLOOKUP)

```excel
=INDEX(D2:D6, XMATCH("E003", A2:A6))
```

This combination works like XLOOKUP but uses INDEX/MATCH syntax.

---

## 3. INDEX/MATCH Combinations

### Why INDEX/MATCH Still Matters

Even with XLOOKUP, INDEX/MATCH is valuable because:
- It works in **all Excel versions** (including 2010, 2013, 2016)
- It's the foundation many professionals learned
- It's more flexible for certain multi-criteria scenarios

### Basic Syntax

```
=INDEX(return_range, MATCH(lookup_value, lookup_range, match_type))
```

### Example 1: Single Criteria Lookup

Using the employee table from Section 1:

```excel
=INDEX(D2:D6, MATCH("E003", A2:A6, 0))
```

**Result:** `$88,000`

- `MATCH("E003", A2:A6, 0)` → finds "E003" is in position 3
- `INDEX(D2:D6, 3)` → returns the 3rd value in column D

### Example 2: Left Lookup (Something VLOOKUP Can't Do)

```excel
=INDEX(A2:A6, MATCH("Carol Davis", B2:B6, 0))
```

**Result:** `E003` (looks LEFT from the Name column to return the Emp ID)

### Example 3: Two-Way Lookup (Row + Column)

```excel
=INDEX(E2:H5, MATCH("South", A2:A5, 0), MATCH("Q3", E1:H1, 0))
```

- Row match: "South" is in row 3 of the data range
- Column match: "Q3" is in column 3 of the data range
- Returns the value at intersection → **10,000**

### Example 4: Multiple Criteria with INDEX/MATCH

To look up a value based on **two criteria** (e.g., Name = "Alice Johnson" AND Department = "Engineering"):

```excel
=INDEX(D2:D6, MATCH(1, (B2:B6="Alice Johnson") * (C2:C6="Engineering"), 0))
```

> ⚠️ **Important:** This is an **array formula**. In Excel 2019 and earlier, press **Ctrl+Shift+Enter** instead of just Enter. In Excel 365/2021, just press Enter.

**How it works:**
- `(B2:B6="Alice Johnson")` returns `{TRUE;FALSE;FALSE;FALSE;FALSE}` → `{1;0;0;0;0}`
- `(C2:C6="Engineering")` returns `{TRUE;FALSE;TRUE;FALSE;FALSE}` → `{1;0;1;0;0}`
- Multiplying: `{1;0;0;0;0}` × `{1;0;1;0;0}` = `{1;0;0;0;0}`
- MATCH finds the first `1` → position 1
- INDEX returns the 1st salary → **$95,000**

### Example 5: Return All Matches (with SMALL/IF - Legacy Approach)

```excel
=INDEX(D2:D6, SMALL(IF(C2:C6="Engineering", ROW(C2:C6)-ROW(C2)+1), ROW(1:1)))
```

> ⚠️ Array formula - Ctrl+Shift+Enter in older Excel versions.

This returns the 1st Engineering salary. Copy down to get the 2nd, 3rd, etc.

> 💡 **Modern Alternative:** Use `=FILTER(D2:D6, C2:C6="Engineering")` in Excel 365!

---

## 4. OFFSET Function and Dynamic Ranges

### What is OFFSET?

OFFSET returns a reference to a range that is a specified number of rows and columns from a starting cell or range. It's commonly used to create **dynamic named ranges** that automatically expand as data is added.

### Syntax

```
=OFFSET(reference, rows, cols, [height], [width])
```

| Argument | Description |
|----------|-------------|
| `reference` | The starting point (cell or range) |
| `rows` | Number of rows to move from the starting point (can be negative) |
| `cols` | Number of columns to move from the starting point (can be negative) |
| `height` | [Optional] Height of the returned range in rows |
| `width` | [Optional] Width of the returned range in columns |

### Example 1: Basic OFFSET

```excel
=OFFSET(A1, 2, 3)
```

Starting from A1, move **2 rows down** and **3 columns right** → returns the value in **D3**.

### Example 2: Return a Range

```excel
=SUM(OFFSET(A1, 0, 0, 5, 1))
```

Returns the sum of A1:A5 - a range that is 5 rows tall and 1 column wide starting from A1.

### Example 3: Dynamic Named Range

Create a named range that automatically grows as you add data:

1. Go to **Formulas → Name Manager → New**
2. Name: `SalesData`
3. Refers to: `=OFFSET(Sheet1!$A$1, 0, 0, COUNTA(Sheet1!$A:$A), 1)`

**How it works:**
- `COUNTA($A:$A)` counts non-empty cells in column A
- OFFSET starts at A1 and creates a range that many rows tall
- As you add data, the range automatically expands

### Example 4: Rolling Average (Last N Values)

```excel
=AVERAGE(OFFSET(A1, COUNTA(A:A)-7, 0, 7, 1))
```

Returns the average of the **last 7 values** in column A.

### ⚠️ OFFSET Limitations

- **Volatile function**: Recalculates every time the worksheet recalculates, even if its inputs haven't changed. This can slow down large workbooks.
- **Not compatible with structured table references.**
- **Modern alternative:** Use `INDEX` to create non-volatile dynamic ranges:
  ```
  =A1:INDEX(A:A, COUNTA(A:A))
  ```

---

## 5. INDIRECT Function and Dynamic References

### What is INDIRECT?

INDIRECT converts a **text string** into a valid cell reference. This allows you to build cell references dynamically - for example, pulling data from different sheets based on a dropdown selection.

### Syntax

```
=INDIRECT(ref_text, [a1])
```

| Argument | Description |
|----------|-------------|
| `ref_text` | A text string that represents a cell reference (e.g., "A1", "Sheet2!B5") |
| `a1` | [Optional] TRUE = A1-style reference (default), FALSE = R1C1-style |

### Example 1: Basic Dynamic Reference

If cell A1 contains the text `"B5"`:

```excel
=INDIRECT(A1)
```

**Result:** Returns the value in cell **B5**.

### Example 2: Dynamic Sheet Reference

Create a dropdown in cell A1 with sheet names (e.g., "January", "February", "March"). Then:

```excel
=INDIRECT("'" & A1 & "'!B10")
```

This pulls the value from cell B10 on whichever sheet is selected in A1.

**Explanation:**
- If A1 = "January", the formula builds: `='January'!B10`
- If A1 = "February", it becomes: `='February'!B10`

### Example 3: Dynamic Range with INDIRECT

```excel
=SUM(INDIRECT("D2:D" & A1))
```

If A1 = 100, this sums D2:D100. Change A1 to 200, and it sums D2:D200.

### Example 4: Dependent Dropdown Lists

Use INDIRECT for cascading/dependent dropdown validation:

1. Name your lists: `Fruits` = {"Apple","Banana","Cherry"}, `Vegetables` = {"Carrot","Broccoli","Spinach"}
2. Cell A1: Data Validation list → "Fruits", "Vegetables"
3. Cell B1: Data Validation list → `=INDIRECT(A1)`

Now B1's dropdown changes based on A1's selection!

### ⚠️ INDIRECT Limitations

- **Volatile function** (recalculates on every change)
- **Cannot reference closed workbooks**
- **Text-based references are fragile** - typos in sheet names cause #REF! errors

---

## 6. CHOOSE Function

### What is CHOOSE?

CHOOSE returns a value from a list of values based on a **position number**. It's useful for creating simple lookup tables or selecting from predefined options.

### Syntax

```
=CHOOSE(index_num, value1, [value2], [value3], ...)
```

### Example 1: Day of the Week Name

```excel
=CHOOSE(WEEKDAY(A1), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
```

If A1 is a date, this returns the day name.

### Example 2: Month Name from Number

```excel
=CHOOSE(MONTH(A1), "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
```

### Example 3: Conditional Category Labels

```excel
=CHOOSE((A1>=90)*1 + (A1>=80)*2 + (A1>=70)*3 + (A1>=60)*4 + (A1<60)*5, "A", "B", "C", "D", "F")
```

> 💡 Note: For more complex scenarios, `IFS()` or nested `IF()` may be clearer.

### Example 4: CHOOSE with INDEX for Dynamic Column Selection

```excel
=SUM(CHOOSE({1,2,3}, B2:B10, D2:D10, F2:F10))
```

This sums columns B, D, and F together - useful when column selection needs to be dynamic.

---

## 7. Dynamic Arrays

### What Are Dynamic Arrays?

Dynamic Arrays are a **revolutionary feature** in Excel 365 and Excel 2021. Instead of returning a single value, a formula can now return an **array of values** that automatically "spills" into neighboring cells.

### The Spill Operator `#`

When a formula spills results into multiple cells, you can reference the entire spill range using the `#` operator:

```excel
=A2#    → References the entire spill range starting at A2
```

**Example:** If `=FILTER(A2:C10, B2:B10="Engineering")` spills into A2:C4, then `A2#` refers to that entire range.

### Common Spill Errors

| Error | Meaning |
|-------|---------|
| `#SPILL!` | Spill range is blocked by existing data in neighboring cells |
| `#CALC!` | No results match the criteria (e.g., FILTER returns nothing) |

---

### 7.1 FILTER Function

**Purpose:** Extract rows from a range that meet one or more conditions.

```
=FILTER(array, include, [if_empty])
```

#### Example 1: Single Condition

```excel
=FILTER(A2:D6, C2:C6="Engineering", "No results")
```

Returns all rows where Department = "Engineering".

**Result (spills into multiple cells):**

| A | B | C | D |
|---|---|---|---|
| E001 | Alice Johnson | Engineering | $95,000 |
| E003 | Carol Davis | Engineering | $88,000 |

#### Example 2: Multiple Conditions (AND)

Use `*` for AND logic:

```excel
=FILTER(A2:D6, (C2:C6="Engineering") * (D2:D6>90000))
```

Returns Engineering employees with salary > $90,000.

#### Example 3: Multiple Conditions (OR)

Use `+` for OR logic:

```excel
=FILTER(A2:D6, (C2:C6="Engineering") + (C2:C6="Sales"))
```

Returns employees in Engineering OR Sales.

#### Example 4: Filter and Sort

Combine FILTER with SORT:

```excel
=SORT(FILTER(A2:D6, C2:C6="Engineering"), 4, -1)
```

Filters for Engineering employees, then sorts by salary (column 4) in **descending** order (-1).

---

### 7.2 SORT and SORTBY Functions

#### SORT

```
=SORT(array, [sort_index], [sort_order], [by_col])
```

| Argument | Default | Description |
|----------|---------|-------------|
| `array` | Required | Range to sort |
| `sort_index` | 1 | Column (or row) number to sort by |
| `sort_order` | 1 (ascending) | 1 = ascending, -1 = descending |
| `by_col` | FALSE | FALSE = sort by rows, TRUE = sort by columns |

**Example - Sort employees by salary descending:**

```excel
=SORT(A2:D6, 4, -1)
```

**Result:**

| E001 | Alice Johnson | Engineering | $95,000 |
|------|--------------|-------------|---------|
| E003 | Carol Davis | Engineering | $88,000 |
| E005 | Eve Wilson | Marketing | $78,000 |
| E002 | Bob Smith | Marketing | $72,000 |
| E004 | David Lee | Sales | $67,000 |

#### SORTBY

Sort by a column that is **not in the output range**:

```
=SORTBY(array, by_array1, [sort_order1], [by_array2], [sort_order2], ...)
```

**Example - Sort names alphabetically but return salary column:**

```excel
=SORTBY(D2:D6, B2:B6, 1)
```

Returns salaries sorted by employee name (A→Z).

**Multi-level sort - Department ascending, then Salary descending:**

```excel
=SORTBY(A2:D6, C2:C6, 1, D2:D6, -1)
```

---

### 7.3 UNIQUE Function

```
=UNIQUE(array, [by_col], [exactly_once])
```

| Argument | Description |
|----------|-------------|
| `array` | Range from which to extract unique values |
| `by_col` | FALSE (default) = unique rows, TRUE = unique columns |
| `exactly_once` | FALSE (default) = all unique values, TRUE = values that appear exactly once |

**Example 1 - Unique departments:**

```excel
=UNIQUE(C2:C6)
```

**Result:** `Engineering | Marketing | Sales`

**Example 2 - Unique rows (full records):**

```excel
=UNIQUE(A2:D6)
```

**Example 3 - Combine with SORT:**

```excel
=SORT(UNIQUE(C2:C6))
```

Returns unique departments sorted alphabetically.

---

### 7.4 SEQUENCE Function

Generates an array of sequential numbers.

```
=SEQUENCE(rows, [columns], [start], [step])
```

**Example 1 - Numbers 1 to 10:**

```excel
=SEQUENCE(10)
```

**Result:** `{1; 2; 3; 4; 5; 6; 7; 8; 9; 10}` (spills down)

**Example 2 - 4×5 multiplication table:**

```excel
=SEQUENCE(4, 5, 1, 1)
```

Creates a 4-row, 5-column grid of numbers.

**Example 3 - Generate dates for 30 days:**

```excel
=TODAY() + SEQUENCE(30) - 1
```

**Example 4 - Odd numbers from 1 to 19:**

```excel
=SEQUENCE(10, 1, 1, 2)
```

**Result:** `{1; 3; 5; 7; 9; 11; 13; 15; 17; 19}`

---

### 7.5 RANDARRAY Function

Generates an array of random numbers.

```
=RANDARRAY([rows], [columns], [min], [max], [whole_number])
```

**Example 1 - 5 random decimals between 0 and 1:**

```excel
=RANDARRAY(5)
```

**Example 2 - 10 random integers between 1 and 100:**

```excel
=RANDARRAY(10, 1, 1, 100, TRUE)
```

**Example 3 - Random sample (shuffle a list):**

```excel
=SORTBY(A2:A100, RANDARRAY(COUNTA(A2:A100)))
```

Randomly shuffles the list in A2:A100.

---

## 8. LET Function

### What is LET?

The LET function lets you **define named variables inside a formula**. This makes complex formulas:
- **Easier to read** - meaningful variable names instead of repeated expressions
- **Faster to calculate** - expensive sub-calculations run only once
- **Easier to maintain** - change a value in one place instead of multiple

### Syntax

```
=LET(name1, value1, [name2, value2, ...], calculation)
```

- You can define up to **126 name/value pairs**
- The **last argument** must be the final calculation
- Variables are scoped to the formula only (can't reference them from other cells)

### Example 1: Simple LET

Without LET:
```excel
=IF(SUM(A1:A10) > 1000, SUM(A1:A10) * 0.1, SUM(A1:A10) * 0.05)
```
(SUM is calculated 3 times!)

With LET:
```excel
=LET(total, SUM(A1:A10), IF(total > 1000, total * 0.1, total * 0.05))
```
(SUM is calculated once!)

### Example 2: Complex Formula Made Readable

**Calculate the tax on income using progressive brackets:**

```excel
=LET(
  income, A1,
  bracket1, MIN(income, 10000) * 0.10,
  bracket2, MAX(MIN(income, 40000) - 10000, 0) * 0.20,
  bracket3, MAX(income - 40000, 0) * 0.30,
  bracket1 + bracket2 + bracket3
)
```

With line breaks (Alt+Enter in the formula bar), this reads almost like code!

### Example 3: Filtering and Sorting with LET

```excel
=LET(
  data, A2:D100,
  dept, "Engineering",
  filtered, FILTER(data, CHOOSECOLS(data, 3) = dept),
  SORT(filtered, 4, -1)
)
```

---

## 9. LAMBDA Function

### What is LAMBDA?

LAMBDA lets you create **custom reusable functions** without VBA or macros. You define a function with parameters and a formula, then use it like any built-in function.

### Syntax

```
=LAMBDA(parameter1, [parameter2, ...], calculation)
```

### Example 1: Simple LAMBDA

```excel
=LAMBDA(x, y, x + y)(3, 5)
```

**Result:** `8`

The `(3, 5)` at the end **invokes** the LAMBDA with those values.

### Example 2: Named LAMBDA (Reusable)

1. Go to **Formulas → Name Manager → New**
2. Name: `CelsiusToFahrenheit`
3. Refers to: `=LAMBDA(celsius, celsius * 9/5 + 32)`

Now use it anywhere in your workbook:

```excel
=CelsiusToFahrenheit(100)
```

**Result:** `212`

### Example 3: LAMBDA with Text Processing

Create a named LAMBDA called `Initials`:

```
=LAMBDA(name, TEXTJOIN(".", TRUE, LEFT(TEXTSPLIT(name, " "), 1) & "."))
```

Usage: `=Initials("John Robert Smith")` → `J.R.S.`

### Example 4: Percentage of Total (Named LAMBDA)

```
=LAMBDA(value, total, IF(total = 0, 0, value / total))
```

Usage: `=PctOfTotal(A1, SUM(A:A))`

---

## 10. Lambda Helper Functions

These functions are designed to work with LAMBDA for array processing.

### 10.1 MAP

Applies a LAMBDA to each element of an array and returns an array of results.

```
=MAP(array1, [array2, ...], LAMBDA(parameter, calculation))
```

**Example - Convert temperatures from Celsius to Fahrenheit:**

```excel
=MAP(A2:A10, LAMBDA(c, c * 9/5 + 32))
```

**Example - Concatenate first and last name:**

```excel
=MAP(A2:A10, B2:B10, LAMBDA(first, last, first & " " & last))
```

### 10.2 REDUCE

Reduces an array to a **single value** by applying a LAMBDA cumulatively.

```
=REDUCE(initial_value, array, LAMBDA(accumulator, value, calculation))
```

**Example - Sum of all values:**

```excel
=REDUCE(0, A2:A10, LAMBDA(acc, val, acc + val))
```

(This is equivalent to `=SUM(A2:A10)` but demonstrates the concept.)

**Example - Count values greater than 100:**

```excel
=REDUCE(0, A2:A10, LAMBDA(acc, val, acc + IF(val > 100, 1, 0)))
```

### 10.3 SCAN

Like REDUCE, but returns **every intermediate result** (running total).

```
=SCAN(initial_value, array, LAMBDA(accumulator, value, calculation))
```

**Example - Running total:**

```excel
=SCAN(0, A2:A10, LAMBDA(acc, val, acc + val))
```

**Result:** Spills a cumulative sum into each row.

### 10.4 BYROW

Applies a LAMBDA to each **row** of an array and returns one result per row.

```
=BYROW(array, LAMBDA(row, calculation))
```

**Example - Sum of each row:**

```excel
=BYROW(A2:D10, LAMBDA(row, SUM(row)))
```

**Example - Max value per row:**

```excel
=BYROW(A2:D10, LAMBDA(row, MAX(row)))
```

### 10.5 BYCOL

Applies a LAMBDA to each **column** of an array and returns one result per column.

```
=BYCOL(array, LAMBDA(col, calculation))
```

**Example - Average of each column:**

```excel
=BYCOL(A2:D10, LAMBDA(col, AVERAGE(col)))
```

---

## 11. New Text Functions

### 11.1 TEXTBEFORE

Extracts text **before** a specified delimiter.

```
=TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
```

**Examples:**

```excel
=TEXTBEFORE("John-Smith-Jones", "-")       → "John"
=TEXTBEFORE("John-Smith-Jones", "-", 2)     → "John-Smith"
=TEXTBEFORE("John-Smith-Jones", "-", -1)    → "John-Smith" (from end)
```

### 11.2 TEXTAFTER

Extracts text **after** a specified delimiter.

```
=TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
```

**Examples:**

```excel
=TEXTAFTER("John-Smith-Jones", "-")       → "Smith-Jones"
=TEXTAFTER("John-Smith-Jones", "-", -1)   → "Jones" (after last dash)
=TEXTAFTER("report.pdf", ".")             → "pdf"
```

### 11.3 TEXTSPLIT

Splits text into rows and/or columns based on delimiters.

```
=TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])
```

**Example 1 - Split into columns:**

```excel
=TEXTSPLIT("Apple,Banana,Cherry", ",")
```

**Result:** `Apple | Banana | Cherry` (spills across 3 cells)

**Example 2 - Split into rows:**

```excel
=TEXTSPLIT("Apple;Banana;Cherry", , ";")
```

**Result:** Spills down - Apple, Banana, Cherry in separate rows.

**Example 3 - Split a CSV-like block:**

```excel
=TEXTSPLIT("A1,B1,C1;A2,B2,C2", ",", ";")
```

**Result:** A 2×3 table:

| A1 | B1 | C1 |
|----|----|----|
| A2 | B2 | C2 |

---

## 12. VSTACK and HSTACK

### VSTACK - Stack Arrays Vertically

```
=VSTACK(array1, [array2], ...)
```

**Example:** Combine two tables stacked on top of each other:

```excel
=VSTACK(A1:C3, E1:G4)
```

**Result:** A combined range with rows from both tables.

**Use case:** Consolidate data from multiple sheets/regions into one view.

### HSTACK - Stack Arrays Horizontally

```
=HSTACK(array1, [array2], ...)
```

**Example:** Combine names and scores side by side:

```excel
=HSTACK(A2:A10, D2:D10)
```

**Result:** A range with names in column 1 and scores in column 2.

---

## 13. TOCOL and TOROW

### TOCOL - Convert Array to Single Column

```
=TOCOL(array, [ignore], [scan_by_col])
```

| `ignore` | Behavior |
|----------|----------|
| 0 | Keep all values |
| 1 | Ignore blanks |
| 2 | Ignore errors |
| 3 | Ignore blanks and errors |

**Example:**

```excel
=TOCOL(A1:C3)
```

Converts a 3×3 range into a single column of 9 values.

### TOROW - Convert Array to Single Row

```
=TOROW(array, [ignore], [scan_by_col])
```

**Example:**

```excel
=TOROW(A1:C3)
```

Converts a 3×3 range into a single row of 9 values.

---

## 14. WRAPCOLS and WRAPROWS

### WRAPCOLS - Wrap Values into Columns

```
=WRAPCOLS(vector, wrap_count, [pad_with])
```

**Example:**

```excel
=WRAPCOLS(SEQUENCE(10), 3, "N/A")
```

Wraps numbers 1-10 into a table with 3 rows per column:

| 1 | 4 | 7 | 10 |
|---|---|---|----|
| 2 | 5 | 8 | N/A |
| 3 | 6 | 9 | N/A |

### WRAPROWS - Wrap Values into Rows

```
=WRAPROWS(vector, wrap_count, [pad_with])
```

**Example:**

```excel
=WRAPROWS(SEQUENCE(10), 4, "")
```

Wraps numbers 1-10 into a table with 4 columns per row:

| 1 | 2 | 3 | 4 |
|---|---|---|---|
| 5 | 6 | 7 | 8 |
| 9 | 10 | | |

---

## 15. Practice Exercises

### Exercise 1: XLOOKUP Mastery

**Setup:** Create a product catalog with columns: ProductID, ProductName, Category, Price, Stock

**Tasks:**
1. Use XLOOKUP to find the price of product "P1042"
2. Use XLOOKUP with a custom "Out of Stock" message when ProductID isn't found
3. Use nested XLOOKUP to find the stock level by matching both ProductName and Category
4. Use XLOOKUP with `search_mode = -1` to find the last product in the "Electronics" category

### Exercise 2: Dynamic Array Dashboard

**Setup:** Create a sales data table with: Date, Salesperson, Region, Product, Amount

**Tasks:**
1. Use FILTER to extract all sales for "North" region
2. Use SORT + FILTER to show top 5 sales by amount for the "South" region
3. Use UNIQUE to list all unique salespeople, then SORT alphabetically
4. Use SEQUENCE to generate a month number list and combine with TEXT to create month names
5. Create a formula that filters, sorts, and extracts unique values all at once using LET

### Exercise 3: INDEX/MATCH Multi-Criteria

**Setup:** Employee table with: EmpID, FirstName, LastName, Department, HireDate, Salary

**Tasks:**
1. Find salary using INDEX/MATCH with EmpID as the lookup value
2. Find EmpID using INDEX/MATCH (demonstrating left-lookup)
3. Use INDEX/MATCH with two criteria (FirstName AND Department)
4. Create a two-way lookup using INDEX with two MATCH functions

### Exercise 4: LET and LAMBDA

**Tasks:**
1. Create a LAMBDA that calculates Body Mass Index (BMI = weight / height²)
2. Name it `BMI` and use it: `=BMI(70, 1.75)`
3. Write a LET formula that calculates the average, max, and count of a range, then returns a formatted string
4. Create a LAMBDA called `Grades` that converts a numeric score to letter grades (A/B/C/D/F)

### Exercise 5: Text Processing

**Tasks:**
1. Use TEXTSPLIT to parse "Last, First | Department | Salary" into separate columns
2. Use TEXTBEFORE and TEXTAFTER to extract the domain from email addresses
3. Use VSTACK to combine two separate regional sales tables into one
4. Use TOCOL and WRAPROWS to reshape a vertical list into a 4-column table

### Exercise 6: Lambda Helpers

**Setup:** Sales table with amounts by product and month

**Tasks:**
1. Use BYROW to calculate the total for each product (row sum)
2. Use BYCOL to find the maximum sale in each month (column max)
3. Use SCAN to create a running total of monthly sales
4. Use MAP to apply a 10% markup to all prices
5. Use REDUCE to count how many sales exceed $10,000

---

## 16. Sources and Further Reading

### Microsoft Official Documentation
- [XLOOKUP Function - Microsoft Support](https://support.microsoft.com/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929)
- [XMATCH Function - Microsoft Support](https://support.microsoft.com/office/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312)
- [FILTER Function - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [SORT Function - Microsoft Support](https://support.microsoft.com/office/sort-function-22f63bd0-ccc8-492f-953d-c20e8e44b86c)
- [UNIQUE Function - Microsoft Support](https://support.microsoft.com/office/unique-function-c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e)
- [SEQUENCE Function - Microsoft Support](https://support.microsoft.com/office/sequence-function-57467a98-57e0-4817-9f14-2eb78519ca90)
- [LET Function - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [LAMBDA Function - Microsoft Support](https://support.microsoft.com/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67)
- [TEXTBEFORE, TEXTAFTER, TEXTSPLIT - Microsoft Support](https://support.microsoft.com/en-us/excel)

### Exceljet (Recommended for Examples and Videos)
- [XLOOKUP Function - Exceljet](https://exceljet.net/functions/xlookup-function)
- [FILTER Function - Exceljet](https://exceljet.net/functions/filter-function)
- [LET Function - Exceljet](https://exceljet.net/functions/let-function)
- [Dynamic Array Functions - Exceljet](https://exceljet.net/)

### Community Resources
- [XLOOKUP vs INDEX/MATCH - Chandoo.org](https://chandoo.org/?s=xlookup)
- [Dynamic Arrays Guide - MyOnlineTrainingHub](https://www.myonlinetraininghub.com/excel-dynamic-array-functions)
- [LAMBDA Function Tutorial - Excel Off The Grid](https://exceloffthegrid.com/)

### YouTube Video References
- [XLOOKUP - ExcelIsFun (Mike Girvin)](https://www.youtube.com/watch?v=GPSY2lE7PSA)
- [Dynamic Arrays - Leila Gharani](https://www.youtube.com/watch?v=0EiOaCMq_FA)
- [FILTER Function Deep Dive - ExcelIsFun](https://www.youtube.com/watch?v=MQ3WqK5eJzI)
- [LET and LAMBDA - Leila Gharani](https://www.youtube.com/watch?v=JTXixVvfSZU)
- [INDEX/MATCH vs XLOOKUP - MyOnlineTrainingHub](https://www.youtube.com/watch?v=PN8nUASSEm4)
- [TEXTBEFORE, TEXTAFTER, TEXTSPLIT - Excel Off The Grid](https://www.youtube.com/watch?v=xl_8DLK9oXw)
- [Dynamic Arrays Crash Course - Chandoo](https://www.youtube.com/watch?v=ZP1HjV4M4Y4)
- [LAMBDA Functions Explained - Curbal](https://www.youtube.com/watch?v=miMjEFMNE0Y)

---

> **Next Module:** [Module 6 - Power Query & Power Pivot](../06-Power-Query-Pivot/06-power-query-power-pivot.md)
