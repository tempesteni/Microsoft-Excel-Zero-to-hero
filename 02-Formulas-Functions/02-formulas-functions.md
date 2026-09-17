# Module 2: Microsoft Excel Formulas & Functions

> **Level:** Beginner to Intermediate | **Estimated Time:** 3–4 hours | **Prerequisites:** [Module 1: Excel Fundamentals](../01-Fundamentals/01-excel-fundamentals.md)

---

## Table of Contents

1. [Formula Basics](#1-formula-basics)
2. [Statistical Functions](#2-statistical-functions-sum-average-count-min-max)
3. [Conditional Functions](#3-conditional-functions-sumif-countif-averageif)
4. [Logical Functions](#4-logical-functions-if-and-or-not)
5. [Advanced Logical Functions](#5-advanced-logical-functions-ifs-switch-nested-if)
6. [Text Functions](#6-text-functions)
7. [Date Functions](#7-date-functions)
8. [Lookup Functions](#8-lookup-functions-vlookup-hlookup)
9. [Error Handling](#9-error-handling)
10. [Named Ranges](#10-named-ranges)
11. [Formula Auditing Tools](#11-formula-auditing-tools)
12. [Practice Exercises](#12-practice-exercises)
13. [Quick Reference Cheat Sheet](#13-quick-reference-cheat-sheet)
14. [Video References](#14-video-references)
15. [Sources](#15-sources)

---

## 1. Formula Basics

### 1.1 What is a Formula?

A **formula** is an expression that calculates the value of a cell. Every formula in Excel:
- **Starts with an equals sign** `=`
- Can contain **numbers**, **cell references**, **operators**, and **functions**
- Displays its **result** in the cell and the **formula** in the Formula Bar

### 1.2 Formula Syntax

```
= value1 OPERATOR value2
```

**Examples:**
```
=10+5          → 15
=A1+B1         → Sum of values in A1 and B1
=SUM(A1:A10)   → Sum of values in A1 through A10
```

### 1.3 Arithmetic Operators

| Operator | Meaning | Example | Result |
|----------|---------|---------|--------|
| `+` | Addition | `=5+3` | 8 |
| `-` | Subtraction | `=10-4` | 6 |
| `*` | Multiplication | `=6*7` | 42 |
| `/` | Division | `=20/4` | 5 |
| `^` | Exponentiation | `=2^3` | 8 |
| `%` | Percentage | `=50%` | 0.5 |

### 1.4 Order of Operations (PEMDAS)

Excel follows the standard mathematical order of operations:

1. **P**arentheses `()`
2. **E**xponents `^`
3. **M**ultiplication `*` and **D**ivision `/` (left to right)
4. **A**ddition `+` and **S**ubtraction `-` (left to right)

**Example:**
```
=2+3*4
```
- Multiplication first: `3*4 = 12`
- Then addition: `2+12 = 14`
- Result: **14**

With parentheses:
```
=(2+3)*4
```
- Parentheses first: `2+3 = 5`
- Then multiplication: `5*4 = 20`
- Result: **20**

**Complex Example:**
```
=2^3+4*5-6/3
```
Step by step:
1. Exponents: `2^3 = 8` → `8+4*5-6/3`
2. Multiplication/Division (left to right): `4*5 = 20`, `6/3 = 2` → `8+20-2`
3. Addition/Subtraction (left to right): `8+20 = 28`, `28-2 = 26`
4. **Result: 26**

---

## 2. Statistical Functions: SUM, AVERAGE, COUNT, MIN, MAX

### Sample Data for This Section

Create this table in your Excel sheet:

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **Employee** | **Department** | **Salary** | **Sales** |
| **2** | Alice | Marketing | 55000 | 12000 |
| **3** | Bob | Sales | 48000 | 18500 |
| **4** | Carol | Marketing | 62000 | 15000 |
| **5** | David | IT | 75000 | 9000 |
| **6** | Eve | Sales | 51000 | 22000 |
| **7** | Frank | IT | 70000 | 11000 |

### 2.1 SUM

Adds all numbers in a range.

**Syntax:** `=SUM(number1, [number2], ...)`

**Examples:**
```
=SUM(C2:C7)          → Sum of all salaries: 361000
=SUM(C2,C4,C6)       → Sum of specific cells: 168000
=SUM(C2:C7,D2:D7)    → Sum of two ranges: 448500
=SUM(10,20,30)       → Sum of literal values: 60
```

### 2.2 AVERAGE

Returns the arithmetic mean of a range.

**Syntax:** `=AVERAGE(number1, [number2], ...)`

**Examples:**
```
=AVERAGE(C2:C7)      → Average salary: 60166.67
=AVERAGE(D2:D3)      → Average of first two sales: 15250
```

### 2.3 COUNT

Counts the number of cells that contain **numbers**.

**Syntax:** `=COUNT(value1, [value2], ...)`

**Examples:**
```
=COUNT(C2:C7)        → 6 (all cells contain numbers)
=COUNT(A2:A7)        → 0 (all cells contain text, not numbers)
=COUNT(C2:C7,D2:D7)  → 12
```

### 2.4 COUNTA

Counts the number of **non-empty** cells (text, numbers, errors, etc.).

**Syntax:** `=COUNTA(value1, [value2], ...)`

**Examples:**
```
=COUNTA(A2:A7)       → 6 (all cells have text)
=COUNTA(C2:C7)       → 6
```

### 2.5 MIN and MAX

**MIN** returns the smallest value; **MAX** returns the largest value.

**Syntax:** `=MIN(number1, [number2], ...)` / `=MAX(number1, [number2], ...)`

**Examples:**
```
=MIN(C2:C7)          → 48000 (Bob's salary)
=MAX(C2:C7)          → 75000 (David's salary)
=MIN(D2:D7)          → 9000 (David's sales)
=MAX(D2:D7)          → 22000 (Eve's sales)
```

---

## 3. Conditional Functions: SUMIF, COUNTIF, AVERAGEIF

These functions perform calculations **only on cells that meet a specified condition**.

### 3.1 SUMIF

Sums values in a range that meet a single criterion.

**Syntax:** `=SUMIF(range, criteria, [sum_range])`

- `range` — the range to evaluate (check against the criteria)
- `criteria` — the condition (e.g., `"Marketing"`, `">50000"`, `"A*"`)
- `sum_range` — the range to sum (optional; if omitted, sums the `range`)

**Examples using the sample data above:**
```
=SUMIF(B2:B7,"Marketing",C2:C7)   → 117000 (Alice 55000 + Carol 62000)
=SUMIF(B2:B7,"Sales",C2:C7)       → 99000 (Bob 48000 + Eve 51000)
=SUMIF(B2:B7,"IT",D2:D7)          → 20000 (David 9000 + Frank 11000)
=SUMIF(C2:C7,">60000")            → 205000 (sums salaries over 60000: 62000+75000+70000)
```

### 3.2 SUMIFS

Sums values that meet **multiple** criteria.

**Syntax:** `=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Examples:**
```
=SUMIFS(C2:C7,B2:B7,"Marketing",C2:C7,">55000")  → 62000 (only Carol meets both conditions)
=SUMIFS(D2:D7,B2:B7,"Sales",D2:D7,">15000")      → 405000 (Bob 18500 + Eve 22000)
```

### 3.3 COUNTIF

Counts cells that meet a single criterion.

**Syntax:** `=COUNTIF(range, criteria)`

**Examples:**
```
=COUNTIF(B2:B7,"Marketing")       → 2
=COUNTIF(B2:B7,"Sales")           → 2
=COUNTIF(C2:C7,">60000")          → 3
=COUNTIF(A2:A7,"A*")              → 1 (starts with "A" — Alice)
```

### 3.4 COUNTIFS

Counts cells that meet **multiple** criteria.

**Syntax:** `=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Examples:**
```
=COUNTIFS(B2:B7,"Marketing",C2:C7,">55000")  → 1 (only Carol)
=COUNTIFS(B2:B7,"Sales",D2:D7,">15000")      → 2 (Bob and Eve)
```

### 3.5 AVERAGEIF

Averages values that meet a single criterion.

**Syntax:** `=AVERAGEIF(range, criteria, [average_range])`

**Examples:**
```
=AVERAGEIF(B2:B7,"Marketing",C2:C7)    → 58500 ((55000+62000)/2)
=AVERAGEIF(B2:B7,"Sales",C2:C7)        → 49500 ((48000+51000)/2)
=AVERAGEIF(C2:C7,">60000")             → 68333.33 ((62000+75000+70000)/3)
```

### 3.6 AVERAGEIFS

Averages values that meet **multiple** criteria.

**Syntax:** `=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)`

**Examples:**
```
=AVERAGEIFS(D2:D7,B2:B7,"Sales",D2:D7,">15000")  → 20250 ((18500+22000)/2)
```

### 3.7 Criteria Operators Reference

| Operator | Meaning | Example |
|----------|---------|---------|
| `">100"` | Greater than 100 | `COUNTIF(C:C,">100")` |
| `"<50"` | Less than 50 | `COUNTIF(C:C,"<50")` |
| `">=100"` | Greater than or equal to 100 | `COUNTIF(C:C,">=100")` |
| `"<>0"` | Not equal to 0 | `COUNTIF(C:C,"<>0")` |
| `"A*"` | Starts with "A" (wildcard) | `COUNTIF(A:A,"A*")` |
| `"*son"` | Ends with "son" | `COUNTIF(A:A,"*son")` |
| `"*mi*"` | Contains "mi" | `COUNTIF(A:A,"*mi*")` |
| `"???"` | Exactly 3 characters | `COUNTIF(A:A,"???")` |

---

## 4. Logical Functions: IF, AND, OR, NOT

### 4.1 IF Function

Returns one value if a condition is TRUE, and another if it is FALSE.

**Syntax:** `=IF(logical_test, value_if_true, value_if_false)`

**Examples using sample data:**
```
=IF(C2>60000,"High","Low")          → "Low" (Alice's 55000 is not > 60000)
=IF(C5>60000,"High","Low")          → "High" (David's 75000 > 60000)
=IF(D2>15000,"Bonus","No Bonus")    → "No Bonus" (Alice's 12000 is not > 15000)
```

**Practical Examples:**
```
=IF(A2>=90,"A",IF(A2>=80,"B",IF(A2>=70,"C","F")))   → Nested IF for letter grades
=IF(B2="Sales",C2*0.1,C2*0.05)                       → Different commission rates by department
```

### 4.2 AND Function

Returns TRUE if **ALL** conditions are true.

**Syntax:** `=AND(condition1, [condition2], ...)`

**Examples:**
```
=AND(C2>50000,D2>10000)     → TRUE (Alice: 55000>50000 AND 12000>10000)
=AND(C2>70000,D2>15000)     → FALSE (55000 is not > 70000)
```

### 4.3 OR Function

Returns TRUE if **ANY** condition is true.

**Syntax:** `=OR(condition1, [condition2], ...)`

**Examples:**
```
=OR(C2>70000,D2>20000)      → FALSE (neither condition is true for Alice)
=OR(C5>70000,D5>20000)      → TRUE (David: 75000>70000)
```

### 4.4 NOT Function

Reverses the logical value — TRUE becomes FALSE and vice versa.

**Syntax:** `=NOT(logical)`

**Examples:**
```
=NOT(C2>70000)              → TRUE (55000 is NOT > 70000)
=NOT(B2="Sales")            → TRUE (Alice is NOT in Sales)
```

### 4.5 Combining Logical Functions with IF

**Real-world Example — Performance Bonus:**
```
=IF(AND(C2>50000,D2>10000),"Bonus: $2000","No Bonus")
```
- Alice (Salary 55000, Sales 12000): **Bonus: $2000** ✅
- Bob (Salary 48000, Sales 18500): **No Bonus** ❌ (salary ≤ 50000)

**Another Example — Eligibility Check:**
```
=IF(OR(B2="Sales",B2="Marketing"),"Revenue Team","Support Team")
```
- Alice (Marketing): **Revenue Team**
- David (IT): **Support Team**

---

## 5. Advanced Logical Functions: IFS, SWITCH, Nested IF

### 5.1 Nested IF

Nesting multiple IF functions for multi-condition logic.

**Example — Letter Grades:**
```
=IF(C2>=90000,"A",IF(C2>=80000,"B",IF(C2>=70000,"C",IF(C2>=60000,"D","F"))))
```

**Readability Tip:** For more than 3 conditions, use IFS instead.

### 5.2 IFS Function (Excel 2019+ / Microsoft 365)

Tests multiple conditions and returns the value for the first TRUE condition.

**Syntax:** `=IFS(condition1, value1, condition2, value2, ..., TRUE, default_value)`

**Example — Salary Band:**
```
=IFS(C2>=70000,"Senior",C2>=55000,"Mid-Level",C2>=40000,"Junior",TRUE,"Entry-Level")
```
- Alice (55000): **Mid-Level**
- David (75000): **Senior**
- Bob (48000): **Junior**

**Note:** Always include a `TRUE, "default"` catch-all at the end to avoid #N/A errors.

### 5.3 SWITCH Function (Excel 2019+ / Microsoft 365)

Compares a value against a list of values and returns the first match.

**Syntax:** `=SWITCH(expression, value1, result1, value2, result2, ..., default)`

**Example — Department Codes:**
```
=SWITCH(B2,"Marketing","MKT","Sales","SLS","IT","IT","FIN")
```
- Alice (Marketing): **MKT**
- Bob (Sales): **SLS**
- David (IT): **IT**

**Comparison:**
| Function | Best For | Max Conditions |
|----------|----------|---------------|
| IF | Simple yes/no | 1 |
| Nested IF | 2-3 conditions | 64 (practical limit ~7) |
| IFS | Many conditions | 127 |
| SWITCH | Matching exact values | 126 |

---

## 6. Text Functions

### Sample Data for Text Functions

| | A | B |
|---|---|---|
| **1** | **Full Name** | **Email** |
| **2** | John Smith | john.smith@company.com |
| **3** | jane doe | JANE.DOE@CORP.NET |
| **4** | Bob Johnson | bob.j@startup.io |
| **5** | Alice Marie Brown | alice.b@bigco.com |

### 6.1 LEFT, RIGHT, MID

Extract characters from a text string.

**Syntax:**
- `=LEFT(text, num_chars)` — Extract from the **left**
- `=RIGHT(text, num_chars)` — Extract from the **right**
- `=MID(text, start_num, num_chars)` — Extract from the **middle**

**Examples:**
```
=LEFT(A2,4)           → "John" (first 4 characters)
=RIGHT(B2,4)          → ".com" (last 4 characters)
=MID(A2,6,5)          → "Smith" (5 characters starting at position 6)
=LEFT(A2,FIND(" ",A2)-1)  → "John" (everything before the first space)
```

### 6.2 LEN

Returns the number of characters in a text string.

**Syntax:** `=LEN(text)`

**Examples:**
```
=LEN(A2)              → 10 ("John Smith" = 10 characters)
=LEN(A4)              → 12 ("Bob Johnson" = 12 characters)
```

### 6.3 TRIM

Removes extra spaces from text (leading, trailing, and multiple spaces between words).

**Syntax:** `=TRIM(text)`

**Example:**
```
=TRIM("  Hello   World  ")  → "Hello World"
```

### 6.4 UPPER, LOWER, PROPER

Change text case.

**Syntax:**
- `=UPPER(text)` — ALL CAPS
- `=LOWER(text)` — all lowercase
- `=PROPER(text)` — Title Case

**Examples:**
```
=UPPER(A2)            → "JOHN SMITH"
=LOWER(A3)            → "jane doe"
=PROPER(A3)           → "Jane Doe"
```

### 6.5 CONCATENATE / CONCAT

Joins two or more text strings together.

**Syntax:**
- `=CONCATENATE(text1, text2, ...)` — Legacy function
- `=CONCAT(text1, text2, ...)` — Modern replacement (Excel 2016+)
- `=text1 & text2 & ...` — Ampersand operator (most common)

**Examples:**
```
=CONCATENATE(A2," - ",B2)     → "John Smith - john.smith@company.com"
=CONCAT(A2," - ",B2)          → "John Smith - john.smith@company.com"
=A2 & " - " & B2               → "John Smith - john.smith@company.com"
="First: " & LEFT(A2,4)       → "First: John"
```

### 6.6 TEXTJOIN (Excel 2019+ / Microsoft 365)

Joins text with a specified delimiter.

**Syntax:** `=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`

**Examples:**
```
=TEXTJOIN(", ",TRUE,A2,A3,A4,A5)  → "John Smith, jane doe, Bob Johnson, Alice Marie Brown"
=TEXTJOIN(" | ",FALSE,A2,B2)      → "John Smith | john.smith@company.com"
```

### 6.7 FIND and SEARCH

Find the position of text within a string.

**Syntax:**
- `=FIND(find_text, within_text, [start_num])` — Case-sensitive
- `=SEARCH(find_text, within_text, [start_num])` — Case-insensitive, supports wildcards

**Examples:**
```
=FIND(" ",A2)                    → 5 (position of space in "John Smith")
=SEARCH("@",B2)                  → 11 (position of @ in email)
=LEFT(B2,SEARCH("@",B2)-1)      → "john.smith" (extract username from email)
=MID(B2,SEARCH("@",B2)+1,100)   → "company.com" (extract domain from email)
```

### 6.8 SUBSTITUTE and REPLACE

Replace text within a string.

**Syntax:**
- `=SUBSTITUTE(text, old_text, new_text, [instance_num])` — Replace by content
- `=REPLACE(old_text, start_num, num_chars, new_text)` — Replace by position

**Examples:**
```
=SUBSTITUTE(A2," ",".")         → "John.Smith" (replace space with dot)
=SUBSTITUTE(B2,".com",".org")   → "john.smith@company.org"
```

### 6.9 VALUE and TEXT

Convert between text and numbers.

**Syntax:**
- `=VALUE(text)` — Converts text to number
- `=TEXT(value, format_text)` — Converts number to text with a format

**Examples:**
```
=VALUE("123")                   → 123 (number)
=TEXT(TODAY(),"MM/DD/YYYY")     → "09/10/2026" (formatted date as text)
=TEXT(1234.5,"$#,##0.00")       → "$1,234.50" (formatted currency as text)
```

---

## 7. Date Functions

### 7.1 TODAY and NOW

**Syntax:**
- `=TODAY()` — Returns today's date (updates when the sheet recalculates)
- `=NOW()` — Returns today's date AND current time

**Examples:**
```
=TODAY()              → 9/10/2026 (or current date)
=NOW()                → 9/10/2026 2:30 PM (or current date and time)
```

**Note:** These are volatile functions — they recalculate every time the worksheet changes.

### 7.2 DATE Function

Creates a date from individual year, month, and day values.

**Syntax:** `=DATE(year, month, day)`

**Examples:**
```
=DATE(2026,9,10)      → September 10, 2026
=DATE(2026,12,25)     → December 25, 2026
```

### 7.3 YEAR, MONTH, DAY

Extract components from a date.

**Syntax:**
- `=YEAR(date)` — Returns the year (e.g., 2026)
- `=MONTH(date)` — Returns the month (1-12)
- `=DAY(date)` — Returns the day (1-31)

**Examples:**
```
=YEAR(DATE(2026,9,10))   → 2026
=MONTH(DATE(2026,9,10))  → 9
=DAY(DATE(2026,9,10))    → 10
```

### 7.4 DATEDIF

Calculates the difference between two dates in years, months, or days.

**Syntax:** `=DATEDIF(start_date, end_date, unit)`

| Unit | Returns | Description |
|------|---------|-------------|
| `"Y"` | Years | Complete years between dates |
| `"M"` | Months | Complete months between dates |
| `"D"` | Days | Total days between dates |
| `"YM"` | Months | Months ignoring years |
| `"YD"` | Days | Days ignoring years |
| `"MD"` | Days | Days ignoring months and years |

**Examples:**
```
=DATEDIF(DATE(2020,3,15),TODAY(),"Y")    → 6 (years since March 15, 2020)
=DATEDIF(DATE(2020,3,15),TODAY(),"M")    → 78 (months)
=DATEDIF(DATE(2020,3,15),TODAY(),"D")    → 2371 (days, approximate)
```

### 7.5 EOMONTH

Returns the last day of the month, a specified number of months before or after a date.

**Syntax:** `=EOMONTH(start_date, months)`

**Examples:**
```
=EOMONTH(TODAY(),0)       → September 30, 2026 (end of current month)
=EOMONTH(TODAY(),1)       → October 31, 2026 (end of next month)
=EOMONTH(TODAY(),-1)      → August 31, 2026 (end of previous month)
=EOMONTH(DATE(2026,1,15),0)  → January 31, 2026
```

### 7.6 WEEKDAY

Returns the day of the week as a number (1=Sunday by default, or customizable).

**Syntax:** `=WEEKDAY(date, [return_type])`

| return_type | Week starts on |
|-------------|----------------|
| 1 (default) | Sunday (1) through Saturday (7) |
| 2 | Monday (1) through Sunday (7) |
| 3 | Monday (0) through Sunday (6) |

**Examples:**
```
=WEEKDAY(DATE(2026,9,10))     → 5 (Thursday, when 1=Sunday)
=WEEKDAY(DATE(2026,9,10),2)   → 4 (Thursday, when 1=Monday)
```

### 7.7 NETWORKDAYS

Counts the number of working days between two dates (excludes weekends, optionally holidays).

**Syntax:** `=NETWORKDAYS(start_date, end_date, [holidays])`

**Example:**
```
=NETWORKDAYS(DATE(2026,1,1),DATE(2026,12,31))    → 261 (approximate working days in 2026)
```

---

## 8. Lookup Functions: VLOOKUP, HLOOKUP

### Sample Data for Lookup Section

**Employee Table (A1:D7):**

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **Emp ID** | **Name** | **Department** | **Salary** |
| **2** | 101 | Alice | Marketing | 55000 |
| **3** | 102 | Bob | Sales | 48000 |
| **4** | 103 | Carol | Marketing | 62000 |
| **5** | 104 | David | IT | 75000 |
| **6** | 105 | Eve | Sales | 51000 |
| **7** | 106 | Frank | IT | 70000 |

### 8.1 VLOOKUP (Vertical Lookup)

Searches for a value in the **first column** of a range and returns a value from a specified column in the same row.

**Syntax:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

- `lookup_value` — The value to search for
- `table_array` — The range containing the data (first column must contain the lookup value)
- `col_index_num` — The column number to return from (1 = first column)
- `range_lookup` — `FALSE` for exact match (most common), `TRUE` for approximate match

**Examples:**
```
=VLOOKUP(103,A2:D7,2,FALSE)      → "Carol" (look up Emp ID 103, return column 2)
=VLOOKUP(103,A2:D7,3,FALSE)      → "Marketing" (return column 3)
=VLOOKUP(103,A2:D7,4,FALSE)      → 62000 (return column 4)
=VLOOKUP(105,A2:D7,2,FALSE)      → "Eve"
```

**Practical Example:** If cell F2 contains an employee ID:
```
=VLOOKUP(F2,A2:D7,2,FALSE)       → Returns the employee name for the ID in F2
```

### 8.2 HLOOKUP (Horizontal Lookup)

Same concept as VLOOKUP, but searches across **rows** instead of columns.

**Syntax:** `=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`

**Use case:** When your data is organized horizontally with headers in the first row.

### 8.3 VLOOKUP Limitations

| Limitation | Description |
|-----------|-------------|
| **Left lookup only** | Can only search the leftmost column of the range |
| **No left return** | Can only return values to the RIGHT of the lookup column |
| **Column index is hard-coded** | If you insert/delete columns, the index number may be wrong |
| **Case-insensitive** | Cannot distinguish between "Apple" and "APPLE" |
| **Approximate match default** | If you forget `FALSE`, it may return wrong results |

### 8.4 Modern Alternatives (Microsoft 365 / Excel 2021+)

**XLOOKUP** overcomes all VLOOKUP limitations:
```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

**Example:**
```
=XLOOKUP(103,A2:A7,B2:B7)        → "Carol"
=XLOOKUP(103,A2:A7,D2:D7,"Not Found")  → 62000 (with error handling)
```

**Advantages over VLOOKUP:**
- Can look left OR right
- Returns a range, not a column number
- Built-in error handling
- Default is exact match
- Supports reverse search

---

## 9. Error Handling

### 9.1 Common Excel Errors

| Error | Cause | Example |
|-------|-------|---------|
| `#VALUE!` | Wrong data type in formula | `="text"+5` |
| `#REF!` | Invalid cell reference (deleted cell) | `=A1` when column A is deleted |
| `#NAME?` | Unrecognized function name or text | `=SUN(A1:A10)` (typo) |
| `#DIV/0!` | Division by zero | `=10/0` |
| `#N/A` | Value not found (lookup) | VLOOKUP can't find the value |
| `#NUM!` | Invalid numeric value | `=SQRT(-1)` |
| `#NULL!` | Incorrect range intersection | `=SUM(A1 B1)` (missing colon/comma) |
| `#SPILL!` | Dynamic array can't spill | Array output blocked by data |

### 9.2 IFERROR

Catches **any** error and returns a specified value instead.

**Syntax:** `=IFERROR(value, value_if_error)`

**Examples:**
```
=IFERROR(VLOOKUP(999,A2:D7,2,FALSE),"Not Found")     → "Not Found" (999 doesn't exist)
=IFERROR(A1/B1,"Cannot divide")                       → "Cannot divide" if B1 is 0
=IFERROR(1/0,"Error")                                 → "Error"
```

### 9.3 IFNA

Catches only `#N/A` errors (useful when you want other errors to still show).

**Syntax:** `=IFNA(value, value_if_na)`

**Examples:**
```
=IFNA(VLOOKUP(999,A2:D7,2,FALSE),"Not Found")         → "Not Found"
=IFNA(VLOOKUP(103,A2:D7,2,FALSE),"Not Found")         → "Carol" (no error)
```

### 9.4 ISERROR and ISNA

Return TRUE/FALSE — useful inside IF statements.

**Syntax:**
- `=ISERROR(value)` — Returns TRUE if ANY error
- `=ISNA(value)` — Returns TRUE only if #N/A

**Example:**
```
=IF(ISERROR(VLOOKUP(999,A2:D7,2,FALSE)),"Check ID","OK")  → "Check ID"
```

### 9.5 Error Handling Best Practice

```
=IFERROR(your_formula, "fallback value")
```

Always wrap lookup formulas and complex calculations in IFERROR to prevent #N/A, #REF!, and other errors from breaking your spreadsheet.

---

## 10. Named Ranges

### 10.1 What is a Named Range?

A **named range** is a descriptive name assigned to a cell or range of cells. Instead of writing `A2:D7`, you write `EmployeeData`.

### 10.2 Creating a Named Range

**Method 1: Name Box**
1. Select the range (e.g., A2:D7)
2. Click the **Name Box** (left of the Formula Bar)
3. Type a name (e.g., `EmployeeData`)
4. Press **Enter**

**Method 2: Define Name (Formulas Tab)**
1. Select the range
2. Go to **Formulas** tab → **Define Name**
3. Enter a name, scope (Workbook or specific sheet), and optional comment
4. Click **OK**

**Method 3: Create from Selection**
1. Select a range that includes headers
2. Go to **Formulas** tab → **Create from Selection**
3. Check where the names are (e.g., "Top row", "Left column")
4. Click **OK** — Excel creates names from the headers

### 10.3 Naming Rules

- Must begin with a letter or underscore
- Cannot contain spaces (use underscores: `Employee_Data`)
- Cannot be a cell reference (e.g., avoid names like `A1` or `R1C1`)
- Maximum 255 characters
- Not case-sensitive (`data` and `DATA` are the same)

### 10.4 Using Named Ranges in Formulas

**Before (cell references):**
```
=SUM(C2:C7)
=VLOOKUP(103,A2:D7,2,FALSE)
```

**After (named ranges):**
```
=SUM(Salary)
=VLOOKUP(103,EmployeeData,2,FALSE)
```

**Benefits:**
- Formulas are easier to read and understand
- Easier to maintain — update the range once, all formulas update
- Named ranges appear in the Name Box for quick navigation

### 10.5 Managing Named Ranges

- **View all:** Formulas tab → **Name Manager** (Ctrl+F3)
- **Edit:** Open Name Manager → select name → Edit
- **Delete:** Open Name Manager → select name → Delete
- **Scope:** Names can be scoped to the entire workbook or a specific sheet

---

## 11. Formula Auditing Tools

### 11.1 Trace Precedents and Dependents

Found on the **Formulas** tab → **Formula Auditing** group.

- **Trace Precedents:** Shows arrows pointing TO the selected cell from cells it references
  - Select a cell with a formula → click "Trace Precedents"
  - Blue arrows show which cells feed into the formula

- **Trace Dependents:** Shows arrows pointing FROM the selected cell to cells that reference it
  - Select a cell → click "Trace Dependents"
  - Shows which formulas use this cell's value

- **Remove Arrows:** Click "Remove Arrows" to clear the display

### 11.2 Show Formulas

Toggle between showing formulas and their results:
- **Formulas** tab → **Show Formulas** button
- Or press **Ctrl+`** (backtick, the key above Tab)

This is extremely useful for auditing — all cells show their formulas instead of results.

### 11.3 Evaluate Formula

**Formulas** tab → **Evaluate Formula**

This steps through a formula one calculation at a time, showing you exactly how Excel evaluates it. Essential for debugging complex nested formulas.

### 11.4 Error Checking

**Formulas** tab → **Error Checking**

Excel identifies cells with errors and offers suggestions to fix them. It will walk through each error in the worksheet.

### 11.5 Watch Window

**Formulas** tab → **Watch Window**

Add cells to a floating watch window that stays visible even when you scroll to other parts of the sheet. Useful for monitoring key values while working on distant cells.

---

## 12. Practice Exercises

### Exercise Set A: Statistical Functions

**Data: Monthly Sales Report**

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **Month** | **Region** | **Product** | **Revenue** |
| **2** | Jan | North | Widget | 15000 |
| **3** | Jan | South | Widget | 12000 |
| **4** | Feb | North | Gadget | 18000 |
| **5** | Feb | South | Gadget | 9500 |
| **6** | Mar | North | Widget | 22000 |
| **7** | Mar | South | Widget | 14000 |
| **8** | Apr | North | Gadget | 16500 |
| **9** | Apr | South | Gadget | 11000 |

**Tasks:**
1. Calculate total revenue: `=SUM(D2:D9)` → **118000**
2. Calculate average revenue: `=AVERAGE(D2:D9)`
3. Find the highest revenue: `=MAX(D2:D9)`
4. Find the lowest revenue: `=MIN(D2:D9)`
5. Count how many months had revenue > 15000: `=COUNTIF(D2:D9,">15000")`
6. Sum revenue for North region only: `=SUMIF(B2:B9,"North",D2:D9)`
7. Sum revenue for Widget products in North: `=SUMIFS(D2:D9,B2:B9,"North",C2:C9,"Widget")`
8. Average revenue for Gadget products: `=AVERAGEIF(C2:C9,"Gadget",D2:D9)`

### Exercise Set B: Logical Functions

Using the same data:

1. Create a "Performance" column in E1: In E2, write: `=IF(D2>15000,"High","Low")` — copy down to E9
2. In F1, header "Bonus": In F2, write: `=IF(AND(D2>15000,B2="North"),D2*0.1,0)` — copy down
3. Sum all bonuses: `=SUM(F2:F9)`

### Exercise Set C: Text Functions

**Data:**

| | A |
|---|---|
| **1** | **Email** |
| **2** | john.smith@company.com |
| **3** | jane.doe@corp.net |
| **4** | bob.j@startup.io |

**Tasks:**
1. Extract username (before @): `=LEFT(A2,FIND("@",A2)-1)`
2. Extract domain (after @): `=MID(A2,FIND("@",A2)+1,LEN(A2))`
3. Capitalize the username: `=PROPER(LEFT(A2,FIND("@",A2)-1))`

### Exercise Set D: Date Functions

1. Enter today's date: `=TODAY()`
2. Calculate your age in years: `=DATEDIF(DATE(1990,5,15),TODAY(),"Y")` (use your own birthdate)
3. Find the last day of the current month: `=EOMONTH(TODAY(),0)`
4. Count working days in the current month: `=NETWORKDAYS(EOMONTH(TODAY(),-1)+1,EOMONTH(TODAY(),0))`
5. What day of the week is your birthday this year? `=TEXT(DATE(2026,5,15),"dddd")`

### Exercise Set E: VLOOKUP

**Lookup Table (A1:C5):**

| | A | B | C |
|---|---|---|---|
| **1** | **Code** | **Product** | **Price** |
| **2** | W01 | Widget A | 25.00 |
| **3** | W02 | Widget B | 30.00 |
| **4** | G01 | Gadget A | 45.00 |
| **5** | G02 | Gadget B | 55.00 |

**Tasks:**
1. In cell E2, type `G01`. In F2: `=VLOOKUP(E2,A2:C5,2,FALSE)` → "Gadget A"
2. In G2: `=VLOOKUP(E2,A2:C5,3,FALSE)` → 45.00
3. In H2 (with error handling): `=IFERROR(VLOOKUP("X99",A2:C5,2,FALSE),"Product not found")`
4. Try changing E2 to different codes and watch F2 and G2 update

---

## 13. Quick Reference Cheat Sheet

### Most-Used Functions

| Category | Function | Syntax | Purpose |
|----------|----------|--------|---------|
| Math | `SUM` | `=SUM(range)` | Add numbers |
| Math | `AVERAGE` | `=AVERAGE(range)` | Calculate mean |
| Math | `COUNT` | `=COUNT(range)` | Count numbers |
| Math | `COUNTA` | `=COUNTA(range)` | Count non-empty cells |
| Math | `MIN` / `MAX` | `=MIN(range)` / `=MAX(range)` | Smallest / Largest |
| Conditional | `SUMIF` | `=SUMIF(range,crit,sum_range)` | Conditional sum |
| Conditional | `COUNTIF` | `=COUNTIF(range,crit)` | Conditional count |
| Conditional | `AVERAGEIF` | `=AVERAGEIF(range,crit,avg_range)` | Conditional average |
| Logical | `IF` | `=IF(test,true,false)` | Conditional value |
| Logical | `AND` / `OR` | `=AND(c1,c2)` / `=OR(c1,c2)` | Multiple conditions |
| Text | `LEFT/RIGHT/MID` | `=LEFT(text,n)` | Extract text |
| Text | `LEN` / `TRIM` | `=LEN(text)` / `=TRIM(text)` | Length / Clean spaces |
| Text | `UPPER/LOWER/PROPER` | `=UPPER(text)` | Change case |
| Text | `CONCAT` / `&` | `=A1&"-"&B1` | Join text |
| Date | `TODAY` / `NOW` | `=TODAY()` / `=NOW()` | Current date/time |
| Date | `DATEDIF` | `=DATEDIF(d1,d2,"Y")` | Date difference |
| Lookup | `VLOOKUP` | `=VLOOKUP(val,range,col,FALSE)` | Vertical lookup |
| Lookup | `XLOOKUP` | `=XLOOKUP(val,range,return)` | Modern lookup (365) |
| Error | `IFERROR` | `=IFERROR(formula,fallback)` | Catch errors |

---

## 14. Video References

> **Note:** These are recommended YouTube videos for visual learners. Open the links in your browser to watch.

- **Excel Formulas and Functions Tutorial** — by ExcelIsFun
  - https://www.youtube.com/watch?v=V2V0GcMYbmo

- **VLOOKUP Tutorial for Beginners** — by ExcelIsFun
  - https://www.youtube.com/watch?v=1Ej-5V7wMOk

- **IF Function Excel Tutorial** — by Leila Gharani
  - https://www.youtube.com/watch?v=g7h8DfOBPRQ

- **SUMIF and COUNTIF Functions Explained** — by ExcelJet
  - https://www.youtube.com/watch?v=QkGHzJhJ3kA

- **Excel Text Functions Tutorial** — by ExcelIsFun
  - https://www.youtube.com/watch?v=CyBV7FkKBOk

- **Excel Date Functions Explained** — by Leila Gharani
  - https://www.youtube.com/watch?v=ZO5jDaVfDhE

- **XLOOKUP vs VLOOKUP — Why XLOOKUP is Better** — by Leila Gharani
  - https://www.youtube.com/watch?v=Hn1JlSjvKkI

- **Named Ranges in Excel — Complete Guide** — by ExcelJet
  - https://www.youtube.com/watch?v=kOO31qFmi9A

- **Excel Error Handling: IFERROR, IFNA** — by ExcelIsFun
  - https://www.youtube.com/watch?v=6h0GN2eH2kY

---

## 15. Sources

1. Microsoft Support — "Excel Functions (by category)"
   - https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-836b0d5e8b09

2. Microsoft Support — "VLOOKUP function"
   - https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1

3. Microsoft Support — "IF function"
   - https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2

4. Excel Easy — "Formulas and Functions"
   - https://www.excel-easy.com/introduction/formulas-functions.html

5. Excel Easy — "Functions"
   - https://www.excel-easy.com/functions.html

6. ExcelJet — "Excel Functions"
   - https://exceljet.net/functions

7. ExcelJet — "Excel Formulas"
   - https://exceljet.net/formulas

8. Corporate Finance Institute — "Excel Functions"
   - https://corporatefinanceinstitute.com/resources/excel/study/excel-functions/

9. GCFGlobal — "Excel 2016: Working with Multiple Worksheets"
   - https://edu.gcfglobal.org/en/excel2016/working-with-multiple-worksheets/1/

10. Chandoo — "Excel Formulas"
    - https://chandoo.org/wp/excel-formulas/

11. TrumpExcel — "Excel Formulas"
    - https://trumpexcel.com/excel-formulas/

12. Excel Easy — "VLOOKUP"
    - https://www.excel-easy.com/functions/lookup-reference-functions.html

13. Microsoft Support — "XLOOKUP function"
    - https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d14-489e-8e64-f20dbb6e6dfb

---

*Module 2 Complete. Review [Module 1: Excel Fundamentals ←](../01-Fundamentals/01-excel-fundamentals.md) or practice with the exercises above!*
