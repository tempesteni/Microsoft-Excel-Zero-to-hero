# Quiz 03 — Data Visualization & Formatting

**Module:** Conditional Formatting, Charts, Tables, Sparklines
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** You apply a conditional formatting rule with the formula `=B2>100` to range B2:B50. What happens?
- A) Only B2 is highlighted if it's over 100
- B) Each cell in B2:B50 is highlighted independently if its value exceeds 100
- C) All cells highlight because B2 is checked against 100
- D) An error — formulas can't be used in conditional formatting

**2.** You convert a data range to an Excel Table (Ctrl+T). Which of these is NOT an automatic benefit?
- A) Auto-filling formulas in new columns
- B) Banded rows for readability
- C) Automatic PivotTable creation
- D) Auto-expanding range references

**3.** Which chart type best shows the part-to-whole relationship of 5 budget categories?
- A) Line chart
- B) Bar chart
- C) Pie chart
- D) Scatter chart

**4.** You want to highlight the top 10% of values in a column. Which conditional formatting preset should you use?
- A) Color Scales
- B) Icon Sets
- C) Top/Bottom Rules > Top 10%
- D) Data Bars

**5.** What happens when you add a new row below an Excel Table?
- A) Nothing — you need to manually extend the table
- B) The table auto-expands to include the new row
- C) A dialog asks if you want to extend
- D) The new row stays outside the table

**6.** You create a line chart but the x-axis labels show numbers instead of your month names. What's the fix?
- A) The months are in the wrong row — move them above the data
- B) Right-click the chart > Select Data > Edit Horizontal Axis Labels
- C) Change the chart type to a bar chart
- D) Format cells as Text

**7.** A sparkline in cell E2 shows data from A2:D2. You insert a new column at C. What happens to the sparkline?
- A) It now shows A2:E2
- B) It breaks and shows an error
- C) It still shows A2:D2 (now A2:B2 + D2:E2)
- D) It shows A2:C2 only

**8.** You want conditional formatting that shows a gradient from red (low) to green (high) across a range. Use:
- A) Data Bars
- B) Color Scales
- C) Icon Sets
- D) A custom formula with RGB

**9.** Which is TRUE about structured references in Excel Tables?
- A) `=SUM(Table1[Sales])` sums the Sales column
- B) `=Table1[@Revenue]` refers to the Revenue cell in the current row
- C) `=Table1[[#Totals],[Profit]]` refers to the Profit total row
- D) All of the above

**10.** You have a combo chart with bars on the primary axis and a line on the secondary axis. When is this useful?
- A) When you want two chart types to look fancy
- B) When comparing data with very different scales (e.g., revenue in millions vs. growth rate %)
- C) When you have more than 255 data points
- D) When you want to animate the chart

**11.** What does "Clear Rules > Clear Rules from Entire Sheet" do?
- A) Deletes all data from the sheet
- B) Removes all conditional formatting rules from the sheet
- C) Clears the cell contents of formatted cells
- D) Removes all chart formatting

**12.** You want a conditional formatting rule that highlights an entire row when column C contains "Overdue". What formula should you use?
- A) `=C2="Overdue"`
- B) `=$C2="Overdue"`
- C) `=C$2="Overdue"`
- D) `=$C$2="Overdue"`

**13.** Which sparkline type shows trends across a row of monthly data?
- A) Win/Loss
- B) Line
- C) Column
- D) Both B and C

**14.** You want to show data labels on a pie chart that display both the category name and percentage. How?
- A) Right-click data labels > Format Data Labels > check Category Name and Percentage
- B) Manually type each label
- C) This is not possible in Excel
- D) Use a text box linked to cells

**15.** An Excel Table has a Total Row enabled. What function does it default to for a column containing dates?
- A) SUM
- B) COUNT
- C) AVERAGE
- D) MAX

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

Write the formula or describe the action needed.

**16.** You want conditional formatting that highlights cells in A1:A100 that contain the text "URGENT" (case-insensitive). Write the formula.

> **Your answer:** `_________________________________`

**17.** Write a conditional formatting formula that highlights every other row in A1:F100 (creating a zebra-stripe effect).

> **Your answer:** `_________________________________`

**18.** You have a Table named `SalesData` with columns `[Region]`, `[Product]`, and `[Amount]`. Write a formula that returns the total Amount where Region is "North".

> **Your answer:** `_________________________________`

**19.** Describe the steps to create a PivotChart from an existing Excel Table named `Orders`.

> **Your answer:** `_________________________________`

**20.** Write the structured reference formula to calculate the average of the `Price` column in a Table named `Inventory`.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | Relative reference `B2` adjusts per row — each cell is evaluated independently. |
| 2 | **C** | Tables don't auto-create PivotTables; you must insert one manually. |
| 3 | **C** | Pie charts are designed for part-to-whole with a small number of categories. |
| 4 | **C** | Top/Bottom Rules have a "Top 10%" preset. |
| 5 | **B** | Tables auto-expand when you type in the row directly below. |
| 6 | **B** | Select Data > Edit Horizontal Axis Labels lets you specify the label range. |
| 7 | **A** | Sparklines auto-adjust their data range when columns are inserted. |
| 8 | **B** | Color Scales apply gradient fills based on cell values. |
| 9 | **D** | All three are valid structured reference syntaxes. |
| 10 | **B** | Combo charts with dual axes handle data with vastly different scales. |
| 11 | **B** | "Clear Rules" only removes conditional formatting — not data or charts. |
| 12 | **B** | `$C2` locks the column (C) but lets the row change — highlights the entire row. |
| 13 | **D** | Both Line and Column sparklines show trends; Win/Loss shows positive/negative only. |
| 14 | **A** | Format Data Labels has checkboxes for Category Name, Percentage, Value, etc. |
| 15 | **D** | Date columns default to MAX in the Total Row. |
| 16 | `=SEARCH("URGENT",A1)>0` or `=ISNUMBER(SEARCH("URGENT",A1))` | SEARCH is case-insensitive; ISNUMBER converts the result to TRUE/FALSE. |
| 17 | `=MOD(ROW(),2)=0` | Even-numbered rows get the format; use `=1` for odd rows. |
| 18 | `=SUMIF(SalesData[Region],"North",SalesData[Amount])` | SUMIF with structured references. |
| 19 | Click anywhere in the `Orders` Table → Insert tab → PivotChart → choose location → configure fields. | — |
| 20 | `=AVERAGE(Inventory[Price])` | Structured reference for the Price column average. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 4
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 3
