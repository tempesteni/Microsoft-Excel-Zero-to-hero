# The Fix-This-Workbook Challenge (Find 20 Issues)

Companion to `exercise-fix-this-workbook.xlsx` (deliberately broken Q2 report for a fictional product, YourCompany Inc).

## How the Challenge Works

1. Open `exercise-fix-this-workbook.xlsx` and inspect all 4 sheets (Results, Data, Lookups, Charts) for 5 minutes without editing.
2. List every problem: there are **20 planted issues** in 4 categories.
3. Score: 16-20 found = hero eyes, 11-15 = solid, 6-10 = revisit Modules 1-4, 0-5 = restudy then retry.
4. Then rebuild the workbook properly: clean Data, working formulas, honest charts. Rebuilding is where the learning happens.

## Issue Worksheet (write your findings before checking the key)

| # | Sheet | Category | What is wrong | Which module fixes it |
|---|-------|----------|---------------|-----------------------|
| 1 | | | | |
| ... up to 20 | | | | |

---

## ANSWER KEY (spoilers below the line)

### Formula Issues (5)

1. **Hardcoded total**: `Results!B3` holds the literal value `25` where the label promises a formula. Fix: live `=SUM` formulas (Module 2).
2. **#REF! error**: `Results!B4` is `=SUM(#REF!)`, a deleted-range leftover. Fix: rebuild the reference (Module 2 error handling).
3. **#DIV/0! error**: `Results!B5` divides by the zero in `B7`. Fix: `IFERROR` or fix the denominator (Module 2).
4. **Off-by-one SUM**: `Data!H16` is `=SUM(H2:H12)` but data runs to row 14, silently missing 2 rows. Fix: structured table references (Modules 2, 3).
5. **VLOOKUP approximate match**: `Lookups!B4` uses `VLOOKUP(...,TRUE)` against an unsorted SKU column, returning wrong prices silently. Fix: exact match `FALSE`/`0` or XLOOKUP (Module 5).

### Data Issues (5)

6. **Numbers stored as text**: `Data` column D (Qty) is text-formatted, breaking every SUM. Fix: Text to Columns or `--` coercion (Module 1).
7. **Dates stored as text**: `Data` column B holds strings like `03/05/2026`, so no date math works. Fix: real date values (Module 1).
8. **Trailing spaces in keys**: `"Northwind Traders "` and `"Harborline Credit Union "` break lookups and grouping. Fix: TRIM (Module 2 text functions).
9. **Duplicate rows**: rows 10-11 duplicate rows 4-5 exactly. Fix: Remove Duplicates (Module 4).
10. **Blank header**: `Data!F1` is empty, so the column cannot be part of a table. Fix: name every header (Module 3).

### Formatting Issues (5)

11. **Three fonts in one workbook**: Impact (Results title), Comic Sans MS (Data), Calibri everywhere else. Fix: one theme font pair (Module 3).
12. **Merged cells inside data**: `Data!A8:B8` and header `G1:H1` are merged, breaking sorting and tables. Fix: unmerge, use Center Across Selection (Modules 1, 3).
13. **Low contrast**: `Results!A10` is light gray on light yellow. Fix: contrast-safe colors (Module 3).
14. **Rainbow header row**: `Data` row 1 uses 5 bright fills with no system. Fix: one palette (Module 3).
15. **Mixed number formats**: `Data` column E mixes currency and General formats. Fix: consistent number formats (Module 3).

### Charts & Structure Issues (5)

16. **Pie chart for 12 months of trend data**: `Charts` pie of monthly revenue. Fix: line or column over time (Module 4).
17. **Truncated y-axis**: the "Revenue growth" chart starts at 78 on data ranging 42-91, exaggerating the climb. Fix: zero baseline (Module 4).
18. **Chart junk**: same chart shows legend + data labels + gridlines at once. Fix: data-ink cleanup (Module 4).
19. **Hidden column with real data**: `Data` column C ("corrections") is hidden, so nobody sees the notes that change the numbers. Fix: unhide, audit before trusting (Module 4).
20. **Circular reference**: `Data!E2` contains `=E2`. Fix: never let a cell depend on itself (Module 2).

## Grading Yourself

- 1 pt per issue spotted (sheet + what is wrong).
- Rebuild the workbook cleanly and re-run every formula: totals must match the rows.
- Log your score in `../10-Study-Plan/progress-tracker.md`.
