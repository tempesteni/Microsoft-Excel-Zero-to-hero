# Excel Functions Cheat Sheet - Top 50

> **Quick Reference** | Print on A4/Letter | 2 pages

---

## MATH & TRIG FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `SUM` | `=SUM(number1, [number2], …)` | Adds all numbers in a range | `=SUM(A1:A10)` → 55 |
| `SUMIF` | `=SUMIF(range, criteria, [sum_range])` | Sums cells matching criteria | `=SUMIF(A:A,"Sales",B:B)` |
| `SUMIFS` | `=SUMIFS(sum_range, crit_range1, crit1, …)` | Sums with multiple criteria | `=SUMIFS(C:C,A:A,"East",B:B,">100")` |
| `SUMPRODUCT` | `=SUMPRODUCT(array1, [array2], …)` | Multiplies then sums arrays | `=SUMPRODUCT(A1:A5,B1:B5)` |
| `ROUND` | `=ROUND(number, num_digits)` | Rounds to specified digits | `=ROUND(3.456,2)` → 3.46 |
| `ROUNDUP` | `=ROUNDUP(number, num_digits)` | Always rounds up | `=ROUNDUP(3.1,0)` → 4 |
| `ROUNDDOWN` | `=ROUNDDOWN(number, num_digits)` | Always rounds down | `=ROUNDDOWN(3.9,0)` → 3 |
| `INT` | `=INT(number)` | Rounds down to nearest integer | `=INT(8.9)` → 8 |
| `ABS` | `=ABS(number)` | Returns absolute value | `=ABS(-5)` → 5 |
| `MOD` | `=MOD(number, divisor)` | Returns remainder after division | `=MOD(10,3)` → 1 |
| `RAND` | `=RAND()` | Random number between 0 and 1 | `=RAND()` → 0.73 |
| `RANDBETWEEN` | `=RANDBETWEEN(bottom, top)` | Random integer in range | `=RANDBETWEEN(1,100)` |
| `SQRT` | `=SQRT(number)` | Square root | `=SQRT(144)` → 12 |
| `POWER` | `=POWER(number, power)` | Raises to a power | `=POWER(2,10)` → 1024 |
| `SUBTOTAL` | `=SUBTOTAL(function_num, ref)` | Sum/avg/count ignoring filtered rows | `=SUBTOTAL(109,A1:A100)` |

---

## TEXT FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `LEN` | `=LEN(text)` | Returns number of characters | `=LEN("Hello")` → 5 |
| `LEFT` | `=LEFT(text, [num_chars])` | Extracts characters from left | `=LEFT("Excel",2)` → "Ex" |
| `RIGHT` | `=RIGHT(text, [num_chars])` | Extracts characters from right | `=RIGHT("Excel",3)` → "cel" |
| `MID` | `=MID(text, start_num, num_chars)` | Extracts from middle of text | `=MID("Excel",2,3)` → "xce" |
| `TRIM` | `=TRIM(text)` | Removes extra spaces | `=TRIM(" Hi ")` → "Hi" |
| `UPPER` | `=UPPER(text)` | Converts to uppercase | `=UPPER("hello")` → "HELLO" |
| `LOWER` | `=LOWER(text)` | Converts to lowercase | `=LOWER("HELLO")` → "hello" |
| `CONCATENATE` / `&` | `=text1 & text2` | Joins text strings | `="Hello" & " " & "World"` |
| `CONCAT` | `=CONCAT(text1, [text2], …)` | Joins text (range support) | `=CONCAT(A1:A5)` |
| `TEXTJOIN` | `=TEXTJOIN(delimiter, ignore_empty, …)` | Joins with delimiter | `=TEXTJOIN(",",TRUE,A1:A5)` |
| `SUBSTITUTE` | `=SUBSTITUTE(text, old, new, [n])` | Replaces text | `=SUBSTITUTE("A-B-C","-","/")` |
| `TEXT` | `=TEXT(value, format_text)` | Formats number as text | `=TEXT(0.5,"0%")` → "50%" |
| `FIND` | `=FIND(find_text, within, [start])` | Finds position (case-sensitive) | `=FIND("e","Excel")` → 2 |

---

## DATE & TIME FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `TODAY` | `=TODAY()` | Current date | `=TODAY()` → 9/17/2026 |
| `NOW` | `=NOW()` | Current date and time | `=NOW()` → 9/17/2026 14:30 |
| `DATE` | `=DATE(year, month, day)` | Creates a date from components | `=DATE(2026,9,17)` |
| `YEAR` | `=YEAR(serial_number)` | Extracts year from date | `=YEAR(TODAY())` → 2026 |
| `MONTH` | `=MONTH(serial_number)` | Extracts month from date | `=MONTH(TODAY())` → 9 |
| `DAY` | `=DAY(serial_number)` | Extracts day from date | `=DAY(TODAY())` → 17 |
| `DATEDIF` | `=DATEDIF(start, end, unit)` | Difference between dates | `=DATEDIF(A1,B1,"D")` → days |
| `EOMONTH` | `=EOMONTH(start_date, months)` | End of month | `=EOMONTH(TODAY(),0)` |
| `NETWORKDAYS` | `=NETWORKDAYS(start, end, [holidays])` | Working days between dates | `=NETWORKDAYS(A1,B1)` |
| `WEEKDAY` | `=WEEKDAY(serial, [return_type])` | Day of week (1=Sun) | `=WEEKDAY(TODAY())` |

---

## LOOKUP & REFERENCE FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `VLOOKUP` | `=VLOOKUP(lookup, table, col, [match])` | Vertical lookup (legacy) | `=VLOOKUP(A1,B:D,3,FALSE)` |
| `XLOOKUP` | `=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | Modern lookup (any direction) | `=XLOOKUP(A1,A:A,C:C)` |
| `INDEX` | `=INDEX(array, row_num, [col_num])` | Returns value at position | `=INDEX(B2:D10,3,2)` |
| `MATCH` | `=MATCH(lookup, array, [match_type])` | Returns position of match | `=MATCH("Sales",A:A,0)` → 5 |
| `HLOOKUP` | `=HLOOKUP(lookup, table, row, [match])` | Horizontal lookup | `=HLOOKUP("Q1",A1:D3,2,FALSE)` |
| `INDIRECT` | `=INDIRECT(ref_text, [a1])` | Creates reference from text | `=INDIRECT("A"&B1)` |

---

## LOGICAL FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `IF` | `=IF(logical_test, if_true, if_false)` | Conditional value | `=IF(A1>100,"High","Low")` |
| `IFS` | `=IFS(test1,val1, test2,val2, …)` | Multiple conditions (no nesting) | `=IFS(A1>90,"A",A1>80,"B")` |
| `AND` | `=AND(logical1, [logical2], …)` | TRUE if all conditions met | `=AND(A1>0,B1>0)` |
| `OR` | `=OR(logical1, [logical2], …)` | TRUE if any condition met | `=OR(A1="Yes",B1="Yes")` |
| `IFERROR` | `=IFERROR(value, value_if_error)` | Returns alt value on error | `=IFERROR(A1/B1,0)` |
| `IFNA` | `=IFNA(value, value_if_na)` | Returns alt value on #N/A | `=IFNA(VLOOKUP(…),"Not Found")` |
| `SWITCH` | `=SWITCH(expr, val1, result1, …)` | Match expression to values | `=SWITCH(A1,1,"Mon",2,"Tue")` |

---

## STATISTICAL FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `AVERAGE` | `=AVERAGE(number1, [number2], …)` | Arithmetic mean | `=AVERAGE(A1:A10)` |
| `AVERAGEIF` | `=AVERAGEIF(range, criteria, [avg_range])` | Average with condition | `=AVERAGEIF(A:A,"East",B:B)` |
| `COUNT` | `=COUNT(value1, [value2], …)` | Counts cells with numbers | `=COUNT(A1:A100)` |
| `COUNTA` | `=COUNTA(value1, [value2], …)` | Counts non-empty cells | `=COUNTA(A1:A100)` |
| `COUNTIF` | `=COUNTIF(range, criteria)` | Counts cells matching criteria | `=COUNTIF(A:A,"Sales")` |
| `COUNTIFS` | `=COUNTIFS(range1, crit1, …)` | Counts with multiple criteria | `=COUNTIFS(A:A,"East",B:B,">100")` |
| `MAX` | `=MAX(number1, [number2], …)` | Largest value | `=MAX(A1:A10)` |
| `MIN` | `=MIN(number1, [number2], …)` | Smallest value | `=MIN(A1:A10)` |
| `MEDIAN` | `=MEDIAN(number1, [number2], …)` | Middle value | `=MEDIAN(1,2,3,4,5)` → 3 |
| `MODE` | `=MODE(number1, [number2], …)` | Most frequent value | `=MODE(1,2,2,3)` → 2 |

---

## 📋 Pro Tips

- **`Ctrl+`` `** toggles showing formulas vs results
- Use **named ranges** (`Formulas > Define Name`) for readable formulas
- Press **`F4`** while editing to cycle absolute/relative references (`$A$1` → `A1`)
- **`XLOOKUP`** replaces `VLOOKUP`/`HLOOKUP`/`INDEX-MATCH` in Microsoft 365
- Use **`SUBTOTAL`** instead of `SUM` for filtered tables (function 109 = SUM)
- Wrap any function in **`IFERROR`** to handle errors gracefully
