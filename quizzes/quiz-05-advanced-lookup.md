# Quiz 05 — Advanced Lookup & Dynamic Arrays

**Module:** INDEX/MATCH, XLOOKUP, Dynamic Arrays, OFFSET/INDIRECT
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** Why is `INDEX/MATCH` often preferred over `VLOOKUP`?
- A) It's shorter to type
- B) It can look left (return a column to the left of the lookup column)
- C) It's always faster
- D) It doesn't need a lookup value

**2.** What does `=XLOOKUP("Widget",A:A,B:B,"Not Found",0)` return if "Widget" is not in column A?
- A) `#N/A`
- B) `0`
- C) `"Not Found"`
- D) Blank cell

**3.** You enter `=SEQUENCE(3,2)` in cell A1. What appears?
- A) A single column of 1, 2, 3
- B) A 3-row, 2-column array: `{1,2; 3,4; 5,6}`
- C) An error
- D) The number 6

**4.** What is the "spill range" in dynamic arrays?
- A) The area where a formula result overflows into adjacent cells
- B) A named range you create manually
- C) The error when a formula can't display results
- D) A PivotTable feature

**5.** `=OFFSET(A1,2,3,1,1)` returns the value in which cell?
- A) `A3`
- B) `D3`
- C) `D1`
- D) `A4`

**6.** You have `=XLOOKUP(A1, Products[Name], Products[Price])`. What's the 4th optional argument that controls "not found" behavior?
- A) `match_mode`
- B) `search_mode`
- C) `if_not_found`
- D) `return_array`

**7.** What does `=UNIQUE(A1:A100)` return?
- A) The first unique value only
- B) An array of all distinct values from the range
- C) A count of unique values
- D) TRUE/FALSE for each cell

**8.** A dynamic array formula shows `#SPILL!` error. What's the most likely cause?
- A) The formula is wrong
- B) Cells in the spill range are occupied by other data
- C) The range is too large
- D) The workbook is protected

**9.** Which formula correctly uses INDEX/MATCH to return the price of "Gadget" from a table where names are in A1:A10 and prices in B1:B10?
- A) `=INDEX(B1:B10,MATCH("Gadget",A1:A10,0))`
- B) `=INDEX(A1:A10,MATCH("Gadget",B1:B10,0))`
- C) `=MATCH(B1:B10,"Gadget",0)`
- D) `=INDEX("Gadget",A1:A10,B1:B10)`

**10.** `=SORT(A2:C10,3,-1)` does what?
- A) Sorts by column 3 in ascending order
- B) Sorts by column 3 in descending order
- C) Sorts 3 columns randomly
- D) Returns an error

**11.** What does `=INDIRECT("Sheet"&A1&"!B5")` do when A1 contains `2`?
- A) Returns the value in B5 of Sheet2
- B) Returns the text "Sheet2!B5"
- C) Returns `#REF!`
- D) Returns the value in B5 of the current sheet

**12.** You need a formula that returns the last non-empty value in column A. Which works?
- A) `=LOOKUP(2,1/(A:A<>""),A:A)`
- B) `=INDEX(A:A,COUNTA(A:A))`
- C) `=XLOOKUP("*",A:A,A:A,,2)`
- D) Both A and B (though B assumes no gaps)

**13.** What does `=FILTER(A2:C100, B2:B100>50, "None")` return?
- A) Rows where column B exceeds 50; shows "None" if no matches
- B) Column B values over 50 only
- C) An error
- D) All rows with "None" in column B

**14.** The `match_mode` parameter in XLOOKUP set to `-1` means:
- A) Exact match
- B) Exact match or next smaller item
- C) Exact match or next larger item
- D) Wildcard match

**15.** Which function can create a dynamic named range that automatically expands?
- A) `OFFSET` with `COUNTA`
- B) `INDEX` with `COUNTA`
- C) Both A and B
- D) Neither — dynamic ranges must be Excel Tables

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

**16.** Write an XLOOKUP formula that searches for the value in A1 in the range D1:D100, returns the corresponding value from E1:E100, and shows "N/A" if not found.

> **Your answer:** `_________________________________`

**17.** Write a formula using INDEX/MATCH that finds the sales amount for "Product X" in column B where product names are in column A and sales are in column C.

> **Your answer:** `_________________________________`

**18.** Write a formula that returns the top 5 values from A1:A100 using SORT and SEQUENCE (or SORTBY).

> **Your answer:** `_________________________________`

**19.** Write a formula using FILTER that returns all rows from A1:D100 where column C equals "Active" AND column D is greater than 1000.

> **Your answer:** `_________________________________`

**20.** Create a dynamic named range formula using OFFSET that covers column A from A1 to the last non-empty cell.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | INDEX/MATCH doesn't require the lookup column to be the leftmost. |
| 2 | **C** | The 4th argument ("Not Found") is the fallback when no match exists. |
| 3 | **B** | SEQUENCE(3,2) creates a 3×2 array: 1,2 / 3,4 / 5,6. |
| 4 | **A** | Spill is when a multi-result formula auto-fills adjacent cells. |
| 5 | **B** | OFFSET(A1, row_offset=2, col_offset=3) → A1+2 rows, +3 cols = D3. |
| 6 | **C** | `if_not_found` is the 4th argument for custom fallback values. |
| 7 | **B** | UNIQUE returns an array of all distinct values. |
| 8 | **B** | #SPILL! occurs when the output area has existing data blocking it. |
| 9 | **A** | INDEX returns from the price column; MATCH finds "Gadget" in the name column. |
| 10 | **B** | Third argument `-1` means descending sort. |
| 11 | **A** | INDIRECT evaluates the string as a reference to Sheet2's B5. |
| 12 | **D** | LOOKUP trick and INDEX+COUNTA both work (LOOKUP handles gaps better). |
| 13 | **A** | FILTER returns matching rows; the third argument is the "if_empty" fallback. |
| 14 | **B** | `-1` = exact match or next smaller item (useful for bracket lookups). |
| 15 | **C** | Both OFFSET+COUNTA and INDEX+COUNTA can define expanding ranges. |
| 16 | `=XLOOKUP(A1,D1:D100,E1:E100,"N/A")` | Basic XLOOKUP with custom not-found message. |
| 17 | `=INDEX(C1:C100,MATCH("Product X",A1:A100,0))` | MATCH finds the row, INDEX returns from sales column. |
| 18 | `=SORT(A1:A100,1,-1)` then reference the first 5 rows, or `=TAKE(SORT(A1:A100,1,-1),5)` | TAKE limits the sorted result to top 5. |
| 19 | `=FILTER(A1:D100,(C1:C100="Active")*(D1:D100>1000))` | Multiply conditions for AND logic in FILTER. |
| 20 | `=OFFSET(A1,0,0,COUNTA(A:A),1)` | Starts at A1, spans COUNTA(A:A) rows tall, 1 column wide. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 6
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 5
