# Quiz 08 — Real-World Projects & Dashboards

**Module:** Real-World Projects, Dashboards, Financial Models
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** You're building a dashboard that needs to show KPIs from a large dataset refreshed weekly. The best approach is:
- A) Copy-paste data each week
- B) Use Power Query to load and transform data, with PivotTables and charts that auto-refresh
- C) Manually re-create charts each week
- D) Use only VBA macros

**2.** In a financial model, which practice best prevents accidental formula changes?
- A) Use bright colors to highlight input cells
- B) Color-code cells: blue font for inputs, black for formulas, and protect formula cells
- C) Print the formula bar
- D) Hide all formulas

**3.** Your dashboard has 6 charts and 3 PivotTables. Users complain it's slow. What's the first optimization to try?
- A) Buy a faster computer
- B) Reduce data volume, use data model instead of cell formulas, and limit volatile functions
- C) Delete half the charts
- D) Convert everything to CSV

**4.** When building a revenue forecast model, which Excel feature lets users toggle between optimistic, pessimistic, and base scenarios?
- A) Data Validation dropdown driving scenario inputs
- B) Manual editing of each cell
- C) Separate workbooks for each scenario
- D) Conditional formatting only

**5.** A dashboard should include all EXCEPT:
- A) Clear section headers and labels
- B) Interactive filters (slicers, dropdowns)
- C) Raw data tables with thousands of rows
- D) KPI summary cards

**6.** You're building a P&L statement in Excel. Revenue is in row 10, COGS in row 11. The gross profit formula should be:
- A) `=B10-B11`
- B) `=SUM(B10:B11)`
- C) `=B10/B11`
- D) `=B10*B11`

**7.** What does a "sensitivity analysis" in a financial model show?
- A) How output changes as one or more input assumptions change
- B) Whether the model has errors
- C) The model's file size
- D) How fast the model recalculates

**8.** You want a dashboard chart to update based on a user's dropdown selection. The technique involves:
- A) VBA only
- B) Using INDEX/MATCH or XLOOKUP to pull selected data into a helper range, then charting that range
- C) Creating a separate chart for each option
- D) Using Word Art

**9.** In project management dashboards, a Gantt chart in Excel is typically built using:
- A) A scatter plot with error bars
- B) A stacked bar chart with invisible base bars
- C) A pie chart
- D) Conditional formatting only

**10.** What is the purpose of an "assumptions sheet" in a financial model?
- A) It lists all input variables and assumptions in one place for easy review
- B) It assumes the model is correct
- C) It contains only hardcoded values
- D) It's an optional decorative element

**11.** You need to track budget vs. actual across 12 months. Which visualization works best?
- A) 12 separate pie charts
- B) A clustered bar chart with Budget and Actual as series, months on the x-axis
- C) A single number showing the total difference
- D) A line chart of actuals only

**12.** Dynamic array formulas can make dashboards more powerful because:
- A) They auto-spill results without copying formulas
- B) They integrate with SORT, FILTER, and UNIQUE for live data views
- C) They reduce formula maintenance
- D) All of the above

**13.** For a professional financial model, which layout is recommended?
- A) All inputs, calculations, and outputs on one sheet
- B) Separate sheets for inputs, calculations, and outputs (clear flow)
- C) Random placement wherever there's space
- D) All formulas in one column

**14.** You're building an inventory dashboard. To show stock levels that need reorder, you use:
- A) Conditional formatting (red when stock < reorder point)
- B) A separate "reorder" sheet with manual checks
- C) A macro that deletes items below threshold
- D) Nothing — just print the full list

**15.** What is "model auditing" in Excel?
- A) Checking the model for errors, tracing precedents/dependents, and validating logic
- B) Auditing the file size
- C) Checking if the model runs on Mac
- D) A VBA feature only

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

**16.** Write a formula for a KPI card that calculates the percentage change from last month's revenue (in A1) to this month's (in B1).

> **Your answer:** `_________________________________`

**17.** You have a dropdown in cell D1 that selects a department name. Write a formula that returns the total expenses for the selected department from a Table named `Expenses` with columns `[Department]` and `[Amount]`.

> **Your answer:** `_________________________________`

**18.** Write a formula that calculates the Compound Annual Growth Rate (CAGR) given Beginning Value in A1, Ending Value in B1, and Number of Years in C1.

> **Your answer:** `_________________________________`

**19.** Write a formula that flags rows in a project tracker where the End Date (column D) is before today AND Status (column E) is not "Complete".

> **Your answer:** `_________________________________`

**20.** Describe the structure of a 3-statement financial model (what are the 3 statements and how they link).

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | Power Query + PivotTables + auto-refresh is the standard dashboard data pipeline. |
| 2 | **B** | Color-coding (blue=input, black=formula) + sheet protection is financial modeling best practice. |
| 3 | **B** | Reduce data, use data model, avoid volatile functions (NOW, INDIRECT, OFFSET in excess). |
| 4 | **A** | A Data Validation dropdown lets users select scenarios; formulas reference the selection. |
| 5 | **C** | Dashboards should show summaries — raw data belongs in a separate data sheet. |
| 6 | **A** | Gross Profit = Revenue − COGS. |
| 7 | **A** | Sensitivity analysis shows how outputs respond to changes in input assumptions. |
| 8 | **B** | INDEX/MATCH or XLOOKUP dynamically pull data for the selected item into a chart-ready range. |
| 9 | **B** | Gantt charts use stacked bars: invisible base + visible duration bar per task. |
| 10 | **A** | The assumptions sheet centralizes all inputs for transparency and easy updates. |
| 11 | **B** | Clustered bar charts are ideal for comparing two series across categories (months). |
| 12 | **D** | Dynamic arrays auto-fill, reduce maintenance, and power interactive views. |
| 13 | **B** | Separation of inputs, calculations, and outputs is the standard modeling convention. |
| 14 | **A** | Conditional formatting provides instant visual alerts for low stock. |
| 15 | **A** | Model auditing traces formula logic, finds errors, and validates assumptions. |
| 16 | `=(B1-A1)/A1` or `=(B1-A1)/ABS(A1)` | Percentage change: (New − Old) / Old. Use ABS if values can be negative. |
| 17 | `=SUMIF(Expenses[Department],D1,Expenses[Amount])` | SUMIF with the dropdown value as criteria. |
| 18 | `=(B1/A1)^(1/C1)-1` | CAGR formula: (End/Begin)^(1/years) − 1. |
| 19 | `=AND(D2<TODAY(),E2<>"Complete")` | Use in a helper column or conditional formatting. Returns TRUE for overdue incomplete tasks. |
| 20 | The 3 statements are: **Income Statement** (revenue, expenses, net income), **Balance Sheet** (assets, liabilities, equity), and **Cash Flow Statement** (operating, investing, financing). Net Income flows from the IS to the BS (retained earnings) and to the CFS (starting point for operating cash flow). The BS balances: Assets = Liabilities + Equity. Cash from CFS links back to the BS cash balance. | — |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Excel Proficiency Certificate Ready!
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 8

---

🎓 **Congratulations on completing all 8 module quizzes!**
If you scored 90%+ on all quizzes, you have demonstrated strong Excel proficiency across fundamentals through advanced projects.
