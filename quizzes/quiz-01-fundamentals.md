# Quiz 01 — Excel Fundamentals

**Module:** Interface, Navigation, Cell References, Data Entry, File Formats
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** You need to select cells A1 through D10 quickly. What is the most efficient method?
- A) Click A1, hold Ctrl, click D10
- B) Click A1, hold Shift, click D10
- C) Type `A1:D10` in the Name Box and press Enter
- D) Both B and C work

**2.** You type `3/15` into a cell and press Enter. By default, Excel will:
- A) Display the text "3/15"
- B) Display the date March 15 of the current year
- C) Display the fraction 3/15 reduced to 1/5
- D) Show an error

**3.** Which cell reference, when copied from B2 to D5, still points to the original cell?
- A) `A1`
- B) `$A$1`
- C) `A$1`
- D) `$A1`

**4.** You press `Ctrl + ;` in a cell. What gets entered?
- A) Current time
- B) Current date
- C) The value from the cell above
- D) A formula

**5.** A column shows `#####` in every cell. What is the most likely fix?
- A) The formula has an error
- B) The column is too narrow to display the value
- C) The cells contain circular references
- D) The sheet is protected

**6.** Which file format preserves ALL Excel features including macros and Power Query connections?
- A) `.xlsx`
- B) `.csv`
- C) `.xlsm`
- D) `.xls`

**7.** You want to enter the same value into cells A1, B3, and C5 simultaneously. How?
- A) Type in A1, copy, paste to B3 and C5
- B) Select A1, hold Ctrl, click B3 and C5, type the value, press Ctrl+Enter
- C) It's not possible to enter into non-contiguous cells at once
- D) Use Fill > Series

**8.** What does pressing `Tab` after entering data in a cell do?
- A) Inserts a tab character into the cell
- B) Moves the selection one cell to the right
- C) Moves the selection one cell down
- D) Confirms the entry and stays on the same cell

**9.** Which shortcut jumps directly to the last used cell in a worksheet?
- A) `Ctrl + End`
- B) `Ctrl + Home`
- C) `Ctrl + ↓`
- D) Both A and C, but they behave differently

**10.** You enter `00125` into a cell formatted as General. What displays?
- A) `00125`
- B) `125`
- C) `0.00125`
- D) An error message

**11.** Which of the following is a mixed cell reference?
- A) `A1`
- B) `$A$1`
- C) `$A1`
- D) `Sheet1!A1`

**12.** You want to prevent accidental edits to a specific sheet while allowing edits on others. You should:
- A) Save the file as read-only
- B) Protect that specific sheet with a password
- C) Hide the sheet
- D) Move the sheet to a new workbook

**13.** In Excel's Name Box (left of the formula bar), typing `F15` and pressing Enter will:
- A) Insert the text "F15" into the current cell
- B) Move the cursor to cell F15
- C) Create a named range called F15
- D) Open the Find dialog

**14.** You paste data from a website and it appears in a single cell with line breaks. To split each line into its own row, you should use:
- A) Text to Columns with a space delimiter
- B) Flash Fill
- C) Find & Replace, replacing `Ctrl+J` with a unique delimiter, then Text to Columns
- D) It's impossible to split in-cell line breaks

**15.** Which of these is NOT a valid Excel data type when entering into a cell?
- A) `TRUE`
- B) `#N/A`
- C) `=SUM(A1:A5)`
- D) `12:00 AM`

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

Write the formula or describe the action that produces the required result.

**16.** In cell C1, write a formula that adds the values in A1 and B1 using the `+` operator, but only works correctly when copied across columns (both references should shift). What is the formula?

> **Your answer:** `_________________________________`

**17.** You have data in A1:A100. Write a formula in B1 that, when copied down, shows the running row number starting at 1.

> **Your answer:** `_________________________________`

**18.** Write the formula to display the text "High" if cell A1 is greater than 100, otherwise display "Low", using the IF function.

> **Your answer:** `_________________________________`

**19.** You need to reference cell B5 on a sheet named "Q1 Data" from another sheet. Write the complete cell reference.

> **Your answer:** `_________________________________`

**20.** What Excel feature (not a formula) lets you auto-fill a pattern like "Jan, Feb, Mar" after typing just the first entry?

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **D** | Both Shift+click and Name Box entry select a range. |
| 2 | **B** | Excel auto-converts short date entries to the current year. |
| 3 | **B** | `$A$1` is absolute — it doesn't shift when copied. |
| 4 | **B** | `Ctrl+;` inserts today's date; `Ctrl+Shift+;` inserts time. |
| 5 | **B** | `#####` means the column is too narrow — widen it. |
| 6 | **C** | `.xlsm` is the macro-enabled format that preserves VBA. |
| 7 | **B** | Ctrl+click for non-contiguous selection, then Ctrl+Enter fills all. |
| 8 | **B** | Tab moves right after data entry; Enter moves down. |
| 9 | **D** | `Ctrl+End` goes to last used cell; `Ctrl+↓` goes to last filled cell in column. |
| 10 | **B** | General format drops leading zeros. Use Text or custom format `00000`. |
| 11 | **C** | `$A1` locks the column but not the row — that's mixed. |
| 12 | **B** | Sheet protection allows granular control per sheet. |
| 13 | **B** | The Name Box navigates to any cell or named range. |
| 14 | **C** | Ctrl+J represents a line break in Find & Replace; use it to create a delimiter, then split. |
| 15 | **B** | `#N/A` is an error result, not something you type as data. |
| 16 | `=A1+B1` | With relative references, both shift when copied across columns. |
| 17 | `=ROW()-ROW($A$1)+1` or `=ROW(A1)` | ROW(A1) returns 1 and increments as copied down. |
| 18 | `=IF(A1>100,"High","Low")` | Standard IF with text outputs. |
| 19 | `='Q1 Data'!B5` | Sheet names with spaces need single quotes. |
| 20 | **AutoFill (Fill Handle)** | Drag the fill handle or use Data > Series for pattern recognition. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 2
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 1
