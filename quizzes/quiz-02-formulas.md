# Quiz 02 — Formulas & Functions

**Module:** Formulas, SUM/IF/VLOOKUP, Text/Date Functions, Error Handling
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** Cell A1 contains `500`, B1 contains `200`. What does `=A1/B1` return?
- A) `2.5`
- B) `2`
- C) `2.50` (always two decimals)
- D) Depends on cell formatting

**2.** What does `=IFERROR(VLOOKUP("X",A:B,2,FALSE),"Not Found")` return when "X" doesn't exist in column A?
- A) `#N/A`
- B) `0`
- C) `"Not Found"`
- D) Blank cell

**3.** You want to sum all values in column B where column A equals "East". Which formula is correct?
- A) `=SUMIF(A:A,"East",B:B)`
- B) `=SUMIFS(B:B,A:A,"East")`
- C) `=SUM(IF(A:A="East",B:B))`
- D) All of the above

**4.** Cell A1 contains `"Hello World"`. What does `=MID(A1,7,5)` return?
- A) `"Hello"`
- B) `"World"`
- C) `" Worl"`
- D) `"World "`

**5.** What is the result of `=LEN(TRIM(" Excel "))`?
- A) `7`
- B) `5`
- C) `9`
- D) `8`

**6.** You have `=VLOOKUP(A1,Sheet2!A:D,3,FALSE)`. What does the `3` represent?
- A) The third row to return
- B) The third column of the lookup table to return
- C) Three criteria to match
- D) A three-column range

**7.** Which formula correctly extracts the last 3 characters from cell A1?
- A) `=RIGHT(A1,3)`
- B) `=MID(A1,LEN(A1)-2,3)`
- C) `=LEFT(A1,LEN(A1)-3)`
- D) Both A and B

**8.** `=TODAY()` returns `2025-01-15`. What does `=YEAR(TODAY())+1` return?
- A) `2025`
- B) `2026`
- C) `2026-01-15`
- D) `#VALUE!`

**9.** You see `#REF!` in a cell. What most likely happened?
- A) A formula divides by zero
- B) A referenced cell or range was deleted
- C) The lookup value wasn't found
- D) A circular reference exists

**10.** What does `=COUNTA(A1:A10)` return if A1:A10 contains 7 numbers, 2 text entries, and 1 blank?
- A) `10`
- B) `9`
- C) `7`
- D) `8`

**11.** You want to concatenate A1 and B1 with a space between them. Which works?
- A) `=A1&" "&B1`
- B) `=CONCATENATE(A1," ",B1)`
- C) `=TEXTJOIN(" ",TRUE,A1,B1)`
- D) All of the above

**12.** What does `=ROUND(3.456, 2)` return?
- A) `3.45`
- B) `3.46`
- C) `3.5`
- D) `3.4`

**13.** You want to find the position of the `@` in an email address in A1. Which formula?
- A) `=SEARCH("@",A1)`
- B) `=FIND("@",A1)`
- C) `=MATCH("@",A1,0)`
- D) Both A and B

**14.** `=DATE(2025,14,1)` returns what?
- A) `#VALUE!`
- B) `February 1, 2026`
- C) `January 14, 2025`
- D) `January 1, 2025`

**15.** You have nested IFs returning `#N/A`. Which function lets you provide a fallback value instead?
- A) `=ISNA()`
- B) `=IFERROR()`
- C) `=IFNA()`
- D) Both B and C

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

Write the formula that produces the required result.

**16.** Column A has product names and column B has prices. Write a formula that returns the price for "Widget" using VLOOKUP. The product name column is A and price is in column B.

> **Your answer:** `_________________________________`

**17.** Write a formula that counts how many cells in A1:A100 contain values greater than 50 AND less than 100.

> **Your answer:** `_________________________________`

**18.** Cell A1 contains a full name like `"John Smith"`. Write a formula to extract just the first name.

> **Your answer:** `_________________________________`

**19.** Write a formula that converts a date in A1 to the format `"DD-MMM-YYYY"` (e.g., `15-Jan-2025`).

> **Your answer:** `_________________________________`

**20.** Write a formula using SUMPRODUCT that multiplies quantities in B1:B10 by unit prices in C1:C10 and returns the total.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **D** | Display depends on cell format; the stored value is 2.5 but may show as `2`, `2.50`, etc. |
| 2 | **C** | IFERROR catches #N/A and returns the specified fallback text. |
| 3 | **D** | SUMIF, SUMIFS, and array SUM(IF(...)) all work. SUMIFS is preferred for new workbooks. |
| 4 | **B** | `MID(A1,7,5)` starts at position 7 and grabs 5 chars: "World". |
| 5 | **A** | TRIM removes leading/trailing spaces, leaving "Excel" (5 chars). |
| 6 | **B** | The third argument of VLOOKUP is the column index within the table array. |
| 7 | **D** | `RIGHT(A1,3)` is simplest; `MID` approach also works. |
| 8 | **B** | YEAR extracts 2025, adds 1 → 2026. |
| 9 | **B** | #REF! means a cell/range reference became invalid (usually deleted). |
| 10 | **B** | COUNTA counts non-blank cells: 7 numbers + 2 text = 9. |
| 11 | **D** | All three concatenation methods work. |
| 12 | **B** | ROUND(3.456, 2) rounds to 2 decimal places → 3.46. |
| 13 | **D** | Both SEARCH and FIND locate "@" — FIND is case-sensitive, SEARCH supports wildcards. |
| 14 | **B** | Excel rolls over month 14 to February of the next year. |
| 15 | **D** | IFERROR catches all errors; IFNA catches only #N/A. Both can provide fallback values. |
| 16 | `=VLOOKUP("Widget",A:B,2,FALSE)` | Exact match lookup on "Widget", returns column 2 (price). |
| 17 | `=COUNTIFS(A1:A100,">"&50,A1:A100,"<"&100)` | COUNTIFS with two criteria. |
| 18 | `=LEFT(A1,FIND(" ",A1)-1)` | Finds the space, takes everything left of it. |
| 19 | `=TEXT(A1,"DD-MMM-YYYY")` | TEXT function with custom format string. |
| 20 | `=SUMPRODUCT(B1:B10,C1:C10)` | Element-wise multiplication then sum. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 3
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 2
