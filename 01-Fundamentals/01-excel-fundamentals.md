# Module 1: Microsoft Excel Fundamentals

> **Level:** Beginner | **Estimated Time:** 2–3 hours | **Prerequisites:** A computer with Microsoft Excel installed (or access to Excel Online)

---

## Table of Contents

1. [What is Excel?](#1-what-is-excel)
2. [History and Versions](#2-history-and-versions)
3. [The Excel Interface](#3-the-excel-interface)
4. [Workbook vs Worksheet](#4-workbook-vs-worksheet)
5. [Cells, Rows, and Columns](#5-cells-rows-and-columns)
6. [Basic Navigation](#6-basic-navigation)
7. [Entering Data](#7-entering-data)
8. [Cell References](#8-cell-references)
9. [Basic Data Types and Formats](#9-basic-data-types-and-formats)
10. [Saving and File Formats](#10-saving-and-file-formats)
11. [Excel Online vs Desktop](#11-excel-online-vs-desktop)
12. [Practice Exercises](#12-practice-exercises)
13. [Video References](#13-video-references)
14. [Sources](#14-sources)

---

## 1. What is Excel?

**Microsoft Excel** is a spreadsheet application developed by Microsoft. It is used to organize, format, and calculate data with formulas. Excel is one of the most widely used software applications in the world, with hundreds of millions of users across business, education, science, and personal use.

### Key Capabilities

- **Data Organization:** Store and structure data in rows and columns
- **Calculations:** Perform mathematical, statistical, and financial calculations using formulas and functions
- **Data Analysis:** Sort, filter, pivot, and summarize large datasets
- **Visualization:** Create charts, graphs, and conditional formatting to visualize data
- **Automation:** Use macros (VBA) and Power Query to automate repetitive tasks
- **Collaboration:** Share workbooks and co-author in real time (Microsoft 365)

---

## 2. History and Versions

### Timeline of Major Excel Versions

| Year | Version | Key Feature |
|------|---------|-------------|
| 1985 | Excel 1.0 (Mac) | First release, for Apple Macintosh |
| 1987 | Excel 2.0 (Windows) | First Windows version |
| 1993 | Excel 5.0 | Introduced VBA (Visual Basic for Applications) |
| 1997 | Excel 97 | Improved VBA, toolbars, 65,536 rows |
| 2000 | Excel 2000 | HTML saving, improved PivotTables |
| 2003 | Excel 2003 | XML support, improved lists |
| 2007 | Excel 2007 | **Ribbon interface**, .xlsx format, 1,048,576 rows |
| 2010 | Excel 2010 | Sparklines, Slicers, improved PivotTables |
| 2013 | Excel 2013 | Flash Fill, Timeline filters, new chart types |
| 2016 | Excel 2016 | Power Query built-in, new chart types (Treemap, Sunburst) |
| 2019 | Excel 2019 | TEXTJOIN, IFS, SWITCH, funnel charts, 3D maps |
| 2021 | Excel 2021 | XLOOKUP, XMATCH, dynamic arrays, LET function |
| 365 | Microsoft 365 | Continuous updates, Copilot AI, LAMBDA, co-authoring |

### Current Versions

- **Microsoft 365 (Subscription):** Always up-to-date with the latest features, cloud storage (OneDrive), and AI-powered Copilot
- **Excel 2024 (One-time purchase):** Perpetual license, feature set frozen at release
- **Excel 2021 (One-time purchase):** Previous perpetual license version
- **Excel Online (Free):** Browser-based version at [office.com](https://www.office.com) — limited features but free with a Microsoft account

---

## 3. The Excel Interface

When you open Excel, you see the following interface components:

### 3.1 The Ribbon

The **Ribbon** is the toolbar at the top of the Excel window, organized into tabs. It replaced the old menu system starting with Excel 2007.

**Main Tabs:**

| Tab | Purpose |
|-----|---------|
| **Home** | Clipboard, Font, Alignment, Number, Styles, Cells, Editing |
| **Insert** | Tables, Illustrations, Charts, Sparklines, Filters, Links, Text |
| **Page Layout** | Themes, Page Setup, Scale to Fit, Sheet Options |
| **Formulas** | Function Library, Defined Names, Formula Auditing, Calculation |
| **Data** | Get External Data, Connections, Sort & Filter, Data Tools |
| **Review** | Proofing, Language, Comments, Changes (Protection) |
| **View** | Workbook Views, Show, Zoom, Window, Macros |

**Contextual Tabs** appear when you select specific objects (e.g., Chart Design tab when a chart is selected).

### 3.2 Quick Access Toolbar

Located at the very top-left of the window (above or below the Ribbon). Customize it by clicking the dropdown arrow to add frequently used commands like:
- Save
- Undo (Ctrl+Z)
- Redo (Ctrl+Y)
- Quick Print

### 3.3 Name Box

The **Name Box** is the small field to the left of the Formula Bar. It shows:
- The **cell address** of the currently selected cell (e.g., `A1`)
- Any **named range** if the selected cell belongs to one
- You can also type a cell address (e.g., `G50`) and press **Enter** to jump directly to that cell

### 3.4 Formula Bar

The **Formula Bar** displays the contents (value or formula) of the currently selected cell. When you click into the Formula Bar (or press F2), you can edit the cell's contents. It shows:
- Plain text values as-is
- Formulas starting with `=` (e.g., `=SUM(A1:A10)`)

### 3.5 Sheet Tabs

Located at the bottom-left of the window. Each **sheet tab** represents a separate worksheet within the workbook.

- Click a tab to switch to that worksheet
- Right-click a tab to rename, delete, move, copy, or change tab color
- Click the `+` button to add a new worksheet
- Use the navigation arrows `< >` to scroll through tabs when there are many

### 3.6 Status Bar

The bar at the very bottom of the Excel window. It shows:
- **Mode:** Ready, Edit, Enter (current state)
- **Caps Lock / Num Lock** indicators
- **Quick calculations:** Select a range of numbers and the Status Bar shows SUM, AVERAGE, and COUNT automatically (right-click the Status Bar to customize)
- **Zoom slider** on the right side to magnify the view

---

## 4. Workbook vs Worksheet

### Workbook

A **workbook** is the entire Excel file (`.xlsx`). Think of it as a **book**:
- It can contain **multiple worksheets** (like pages in a book)
- A new workbook opens with 1 worksheet by default (you can change this in Options → General → "Include this many sheets")
- The workbook name appears in the **title bar** at the top

### Worksheet

A **worksheet** (also called a **sheet** or **spreadsheet**) is a single page within a workbook:
- It is a grid of **1,048,576 rows** × **16,384 columns** (columns A through XFD)
- Each worksheet has its own tab at the bottom
- You can have multiple worksheets for related data (e.g., "January Sales", "February Sales", "Summary")

### Practical Example

```
Workbook: "2026 Budget.xlsx"
├── Sheet1: "Income"
├── Sheet2: "Expenses"
├── Sheet3: "Summary"
└── Sheet4: "Charts"
```

---

## 5. Cells, Rows, and Columns

### 5.1 The Grid Structure

| Element | Description | Count |
|---------|-------------|-------|
| **Cell** | The intersection of a row and a column — the basic unit of a spreadsheet | Over 17 billion per sheet |
| **Row** | A horizontal line of cells, numbered 1 to 1,048,576 | 1,048,576 |
| **Column** | A vertical line of cells, lettered A to XFD | 16,384 |

### 5.2 Cell Addressing (A1 Notation)

Every cell has a unique **address** (also called a **cell reference**) formed by its column letter + row number:

- **A1** = Column A, Row 1 (top-left cell)
- **B3** = Column B, Row 3
- **Z100** = Column Z, Row 100
- **AA1** = Column AA (the 27th column)

### 5.3 Selecting Cells

| Action | How To |
|--------|--------|
| Select a single cell | Click on it |
| Select a range | Click and drag from one cell to another (e.g., A1:C5) |
| Select an entire row | Click the row number (e.g., `5`) |
| Select an entire column | Click the column letter (e.g., `D`) |
| Select multiple non-adjacent cells | Hold **Ctrl** and click each cell |
| Select all cells | Press **Ctrl+A** or click the corner button (top-left intersection of row/column headers) |

---

## 6. Basic Navigation

### 6.1 Keyboard Shortcuts for Navigation

| Action | Windows Shortcut | Mac Shortcut |
|--------|-----------------|--------------|
| Move one cell up/down/left/right | Arrow Keys | Arrow Keys |
| Move to the beginning of the row | Home | Fn + ← |
| Move to cell A1 | Ctrl + Home | Fn + ⌘ + ← |
| Move to the last used cell | Ctrl + End | Fn + ⌘ + → |
| Move down one screen | Page Down | Fn + ↓ |
| Move up one screen | Page Up | Fn + ↑ |
| Go to a specific cell | Ctrl + G (or F5) | Fn + ⌘ + G |
| Jump to edge of data region | Ctrl + Arrow Key | ⌘ + Arrow Key |
| Go to next sheet | Ctrl + Page Down | ⌘ + Fn + ↓ |
| Go to previous sheet | Ctrl + Page Up | ⌘ + Fn + ↑ |

### 6.2 Using the Name Box to Navigate

1. Click on the **Name Box** (left of the Formula Bar)
2. Type a cell address (e.g., `G500`)
3. Press **Enter**
4. Excel jumps directly to that cell

### 6.3 Using Go To (Ctrl+G)

1. Press **Ctrl+G** (or **F5**)
2. The "Go To" dialog box appears
3. Type a cell reference or range (e.g., `A1:Z100`)
4. Click **OK**

You can also click **Special...** to go to specific types of cells (blanks, formulas, comments, etc.).

---

## 7. Entering Data

### 7.1 Entering Text

1. Click on a cell
2. Start typing
3. Press **Enter** to move down, **Tab** to move right, or **arrow keys** to move in any direction
4. Press **Escape** to cancel entry

**Tips:**
- Text aligns to the **left** by default
- If your text is longer than the column width, it will overflow into the next cell(s) — unless those cells contain data
- To start a new line **within** a cell, press **Alt+Enter**

### 7.2 Entering Numbers

1. Click on a cell
2. Type the number
3. Press **Enter**

**Tips:**
- Numbers align to the **right** by default
- Do not type currency symbols, commas, or percentage signs if you want to apply number formatting later
- To enter a fraction like ½, type `0 1/2` (zero, space, 1/2)
- To enter a number as text (e.g., a zip code starting with 0), type an apostrophe first: `'01234`

### 7.3 Entering Dates

1. Click on a cell
2. Type a date in a recognized format (e.g., `1/15/2026`, `Jan 15, 2026`, `2026-01-15`)
3. Press **Enter**

**Tips:**
- Excel stores dates as **serial numbers** (January 1, 1900 = 1)
- You can format dates in many display formats (right-click → Format Cells → Date)
- Use shortcuts: **Ctrl+;** enters today's date, **Ctrl+Shift+;** enters the current time

### 7.4 Entering Formulas

1. Click on a cell
2. Type `=` to begin a formula
3. Enter the formula (e.g., `=A1+B1`)
4. Press **Enter**

**Example:** To add the values in cells A1 and B1:
```
=A1+B1
```

See [Module 2: Formulas & Functions](../02-Formulas-Functions/02-formulas-functions.md) for detailed formula coverage.

### 7.5 AutoFill

Excel can automatically fill a series of data:

1. Enter a starting value (e.g., `1` in A1, `2` in A2)
2. Select both cells (A1:A2)
3. Drag the **fill handle** (the small square at the bottom-right corner of the selection) down
4. Excel continues the pattern: `3`, `4`, `5`, ...

AutoFill works with:
- Numbers (1, 2, 3, ... or 2, 4, 6, ...)
- Days of the week (Mon, Tue, Wed, ...)
- Months (Jan, Feb, Mar, ...)
- Custom lists

---

## 8. Cell References

Cell references are the foundation of Excel formulas. Understanding how they behave when copied is critical.

### 8.1 Relative References (Default)

A **relative reference** changes when you copy the formula to another cell.

**Example:**
```
Cell C1: =A1+B1
```
If you copy C1 down to C2, the formula automatically becomes:
```
Cell C2: =A2+B2
```

The reference "shifts" relative to the new position.

### 8.2 Absolute References ($A$1)

An **absolute reference** does NOT change when you copy the formula. Use `$` signs to lock the column and/or row.

**Example:**
```
Cell B1: =A1*$D$1
```
If you copy B1 down to B2:
```
Cell B2: =A2*$D$1
```
The `$D$1` stays fixed, while `A1` changes to `A2` (relative).

### 8.3 Mixed References ($A1 or A$1)

A **mixed reference** locks either the column OR the row, but not both.

| Reference | Column | Row | When Copied |
|-----------|--------|-----|-------------|
| `$A1` | **Locked** (always column A) | Relative | Row changes, column stays A |
| `A$1` | Relative | **Locked** (always row 1) | Column changes, row stays 1 |

**Practical Example — Multiplication Table:**

| | A | B | C |
|---|---|---|---|
| **1** | | 2 | 3 |
| **2** | 2 | | |
| **3** | 3 | | |

In cell B2, enter:
```
=$A2*B$1
```

- `$A2` — column is locked to A, row adjusts when copied down
- `B$1` — row is locked to 1, column adjusts when copied right

When copied across B2:C3, this produces a perfect multiplication table.

### 8.4 Quick Toggle: F4 Key

When editing a formula, place your cursor on a cell reference and press **F4** to cycle through reference types:
```
A1  →  $A$1  →  A$1  →  $A1  →  A1
(relative → absolute → mixed → mixed → relative)
```

---

## 9. Basic Data Types and Formats

### 9.1 Data Types in Excel

| Type | Description | Example | Alignment |
|------|-------------|---------|-----------|
| **Text** (String) | Letters, words, alphanumeric | "John Smith", "INV-001" | Left |
| **Number** | Numeric values | 42, 3.14, -100 | Right |
| **Date** | Date values (stored as numbers) | 1/15/2026 | Right |
| **Time** | Time values (stored as decimal numbers) | 2:30 PM | Right |
| **Boolean** | TRUE or FALSE | TRUE | Center |
| **Error** | Formula errors | #VALUE!, #REF! | Center |
| **Currency** | Numbers formatted as currency | $1,234.56 | Right |
| **Percentage** | Numbers formatted as percentages | 25% | Right |

### 9.2 Number Formatting

To change the format of a cell:
1. Select the cell(s)
2. Right-click → **Format Cells** (or press **Ctrl+1**)
3. Choose from categories:
   - **General** — default, no specific format
   - **Number** — decimal places, thousands separator
   - **Currency** — adds currency symbol
   - **Accounting** — aligns currency symbols in a column
   - **Date** — various date display formats
   - **Time** — various time display formats
   - **Percentage** — multiplies by 100 and adds %
   - **Text** — treats content as text
   - **Special** — Zip Code, Phone Number, SSN

**Quick Number Format Shortcuts:**

| Shortcut | Format |
|----------|--------|
| Ctrl+Shift+~ | General |
| Ctrl+Shift+! | Number (with commas, 2 decimals) |
| Ctrl+Shift+$ | Currency |
| Ctrl+Shift+% | Percentage |
| Ctrl+Shift+# | Date (dd-mmm-yy) |
| Ctrl+Shift+@ | Time (h:mm AM/PM) |

---

## 10. Saving and File Formats

### 10.1 Saving Files

| Action | Shortcut |
|--------|----------|
| Save | **Ctrl+S** |
| Save As | **F12** |
| Save a copy | File → Save a Copy |

### 10.2 File Formats

| Extension | Format | Description |
|-----------|--------|-------------|
| `.xlsx` | Excel Workbook | **Default format.** XML-based, no macros |
| `.xlsm` | Excel Macro-Enabled Workbook | Contains VBA macros |
| `.xlsb` | Excel Binary Workbook | Faster for large files, smaller file size |
| `.xls` | Excel 97-2003 Workbook | Legacy format (limited to 65,536 rows) |
| `.csv` | Comma-Separated Values | Plain text, no formatting, no formulas |
| `.pdf` | Portable Document Format | Read-only, for sharing/printing |
| `.ods` | OpenDocument Spreadsheet | LibreOffice/Google Sheets compatible |
| `.xltm` | Macro-Enabled Template | Template with macro support |
| `.xltx` | Excel Template | Template without macros |

### 10.3 AutoSave and AutoRecover

- **AutoSave** (Microsoft 365): Automatically saves to OneDrive/SharePoint as you work
- **AutoRecover**: Saves temporary copies at intervals (default: every 10 minutes). Configure in **File → Options → Save**
- If Excel crashes, AutoRecover files appear in the **Document Recovery** pane when you restart

---

## 11. Excel Online vs Desktop

| Feature | Excel Desktop | Excel Online (Free) |
|---------|--------------|---------------------|
| Cost | Requires Microsoft 365 subscription or one-time purchase | Free with Microsoft account |
| Full feature set | ✅ All features | ❌ Limited subset |
| VBA Macros | ✅ Full support | ❌ Not supported |
| Power Query | ✅ Full support | ❌ Not supported |
| PivotTables | ✅ Full support | ✅ Basic support |
| Co-authoring | ✅ (via OneDrive) | ✅ Real-time by default |
| Add-ins | ✅ Full support | ❌ Limited |
| Offline use | ✅ Yes | ❌ Requires internet |
| Advanced charts | ✅ All chart types | ❌ Limited chart types |
| Data Analysis ToolPak | ✅ Yes | ❌ No |
| Maximum rows | 1,048,576 | 1,048,576 |

**When to use Excel Online:**
- Quick edits when you don't have your computer
- Collaborative editing in real-time
- Light data entry and simple formulas
- Accessing files stored on OneDrive from any device

**When to use Excel Desktop:**
- Complex formulas and large datasets
- VBA/macro automation
- Power Query and Power Pivot
- Advanced data analysis and modeling

---

## 12. Practice Exercises

### Exercise 1: Basic Data Entry

Create a new workbook and enter the following data:

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **Product** | **Quantity** | **Price** | **Total** |
| **2** | Apples | 10 | 1.50 | |
| **3** | Bananas | 15 | 0.75 | |
| **4** | Oranges | 8 | 2.00 | |
| **5** | Grapes | 5 | 3.25 | |

**Tasks:**
1. Enter the headers in Row 1 (A1:D1)
2. Enter the product data in Rows 2-5
3. In cell D2, enter the formula `=B2*C2` and press Enter
4. Copy the formula from D2 down to D5
5. Format column D as Currency (Ctrl+1 → Currency)

### Exercise 2: Cell References Practice

Using the data from Exercise 1:

1. In cell A7, type: **Grand Total**
2. In cell D7, use a SUM formula: `=SUM(D2:D5)`
3. In cell A8, type: **Tax Rate**
4. In cell B8, enter: `0.08` (8%)
5. In cell A9, type: **Tax Amount**
6. In cell D9, enter a formula that multiplies D7 by B8 (using an absolute reference for B8)
7. In cell A10, type: **Grand Total with Tax**
8. In cell D10, add D7 + D9

### Exercise 3: Navigation Practice

1. Press **Ctrl+G** (Go To), type `Z50`, press Enter — Excel jumps to Z50
2. Press **Ctrl+Home** — returns to A1
3. Click the **Name Box**, type `A1:D5`, press Enter — selects the range
4. Press **Ctrl+End** — jumps to the last used cell
5. Press **Ctrl+Home** again — returns to A1

### Exercise 4: Formatting Practice

1. Select cells A1:D1 (headers)
2. Press **Ctrl+B** to make them bold
3. Press **Ctrl+1** → Alignment tab → check "Center" horizontal alignment
4. Select column D (click the column letter D)
5. Right-click → Format Cells → Number → 2 decimal places
6. Double-click the line between column letters D and E in the header to auto-fit column D width

### Exercise 5: File Management

1. Save the workbook as `Practice_Exercises.xlsx` (Ctrl+S)
2. Save a copy as `Practice_Exercises.csv` (File → Save As → CSV)
3. Close the workbook (Ctrl+W)
4. Reopen the .csv file — notice that all formatting and formulas are gone (only values remain)
5. Open the .xlsx file — notice everything is preserved

---

## 13. Video References

> **Note:** These are recommended YouTube videos for visual learners. Open the links in your browser to watch.

- **Excel for Beginners | Excel Tutorial** — by ExcelIsFun
  - https://www.youtube.com/watch?v=rwbho0CgEAE

- **Microsoft Excel Tutorial for Beginners - Full Course** — by freeCodeCamp.org
  - https://www.youtube.com/watch?v=Vl0H-qTcl3s

- **Excel Basics - Getting Started with Excel** — by ExcelJet
  - https://www.youtube.com/watch?v=kOO31qFmi9A

- **Excel Interface Explained (Ribbon, Tabs, Quick Access Toolbar)** — by Leila Gharani
  - https://www.youtube.com/watch?v=ywYD6vMI1Bk

- **Absolute vs Relative Cell References in Excel** — by ExcelIsFun
  - https://www.youtube.com/watch?v=dB9s1hGEKkY

- **Excel Keyboard Shortcuts You MUST Know** — by Leila Gharani
  - https://www.youtube.com/watch?v=m5qjSW_M3gQ

---

## 14. Sources

1. Microsoft Support — "Excel Quick Start Guide"
   - https://support.microsoft.com/en-us/office/excel-quick-start-guide-d40e5082-c128-4598-85e0-1d290461c580

2. Microsoft Support — "What's new in Excel"
   - https://support.microsoft.com/en-us/office/what-s-new-in-excel-d73d9d74-9e7c-4e5a-9db3-5f06f6f3b3e1

3. Excel Easy — "Introduction to Excel"
   - https://www.excel-easy.com/introduction.html

4. Excel Easy — "Basics of Excel"
   - https://www.excel-easy.com/basics.html

5. ExcelJet — "Excel Shortcuts"
   - https://exceljet.net/shortcuts

6. GCFGlobal — "Excel 2016: Getting Started with Excel"
   - https://edu.gcfglobal.org/en/excel2016/getting-started-with-excel/1/

7. Corporate Finance Institute — "Excel Shortcuts Overview"
   - https://corporatefinanceinstitute.com/resources/excel/study/excel-shortcuts/

8. TrumpExcel — "Excel Basics"
   - https://trumpexcel.com/learn-excel/

9. Chandoo — "Excel Basics"
   - https://chandoo.org/wp/excel-basics/

10. Excel Easy — "Format Cells"
    - https://www.excel-easy.com/basics/format-cells.html

---

*Module 1 Complete. Proceed to [Module 2: Formulas & Functions →](../02-Formulas-Functions/02-formulas-functions.md)*
