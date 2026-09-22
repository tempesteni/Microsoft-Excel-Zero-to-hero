# Sample 07: VBA Toolkit (Module 7)

Ready-to-use VBA modules plus a practice workbook. Real `.xlsm` files with embedded macros are binary artifacts, so this toolkit ships plain-text `.bas` modules instead: importing them takes 10 seconds and teaches you the VBA editor along the way.

## Contents

| File | Macros | What they do |
|------|--------|--------------|
| `mod_report_tools.bas` | `FormatReport`, `AddSummaryRow` | One-click report formatting + totals row |
| `mod_bulk_tools.bas` | `RemoveEmptyRows`, `ExportEachSheetToCSV`, `RenameSheetsFromColumnA` | Bulk cleaning and export jobs |
| `practice-data.xlsx` | (data only) | Messy multi-sheet workbook to practice on |

## How to Import and Run

1. Open `practice-data.xlsx` and immediately **Save As `practice-data.xlsm`** (macro-enabled). Excel will not keep macros in `.xlsx`.
2. Press **Alt+F11** to open the VBA editor.
3. **File > Import File** and pick `mod_report_tools.bas`, then repeat for `mod_bulk_tools.bas`.
4. Close the editor, press **Alt+F8**, pick a macro, click **Run**.
5. If Excel shows a security banner, click **Enable Content** (this is your own code: you can read every line first).

## Practice Path

1. Run `FormatReport` on the `Raw` sheet and read what each block changes.
2. Run `AddSummaryRow` and check the `=SUM` range it wrote.
3. Run `RemoveEmptyRows` on `Messy`, then `RenameSheetsFromColumnA`.
4. Break one macro on purpose (change `Columns(2)` to `Columns(99)`), run it, and read the error. Fixing broken code is half of Module 7.
5. Write your own 5-line macro that stamps today's date in `Z1` of every sheet.

## Safety Note

Macros are code. Never import `.bas` files you have not read, and never enable macros in workbooks from strangers. These two modules are plain text: open them in any editor first.
