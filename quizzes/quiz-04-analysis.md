# Quiz 04 — Data Analysis & PivotTables

**Module:** Sorting, Filtering, PivotTables, What-If Analysis, Data Validation
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** You have a dataset with columns: Date, Region, Product, Sales. You want to see total sales per region per product. The fastest approach is:
- A) Use SUMIF for each region-product combination
- B) Create a PivotTable with Region and Product as rows, Sales as values (SUM)
- C) Sort by Region, then manually add subtotals
- D) Create a separate sheet for each region

**2.** You apply a filter to column A showing only "East" and "North". What happens to the row numbers?
- A) They stay the same with gaps
- B) Hidden rows are deleted
- C) They become sequential (1, 2, 3...)
- D) Hidden rows turn gray but numbers stay

**3.** In a PivotTable, you drag "Region" to Rows and "Sales" to Values. Excel defaults the value calculation to:
- A) SUM of Sales
- B) COUNT of Sales
- C) AVERAGE of Sales
- D) It shows a dialog asking you to choose

**4.** You want to restrict a cell so users can only enter numbers between 1 and 100. You should use:
- A) Conditional Formatting
- B) Data Validation with Whole Number between 1 and 100
- C) A formula in the cell
- D) Sheet protection

**5.** What-If Analysis's Goal Seek finds:
- A) The input value needed to achieve a desired output
- B) The maximum possible output
- C) All scenarios for a formula
- D) Errors in your formulas

**6.** You sort a dataset by Date (oldest to newest), then by Region (A to Z). The result:
- A) Sorts by Region first, then Date
- B) Sorts by Date first, then Region within each date
- C) Sorts alphabetically by Region only
- D) Sorts by Date only

**7.** A PivotTable shows quarterly sales. You want to display each quarter as a percentage of the year's total. How?
- A) Add a calculated field with `=Sales/TotalSales`
- B) Right-click a value > Show Values As > % of Grand Total
- C) Manually calculate percentages in a new column
- D) Change the number format to percentage

**8.** You create a custom error message in Data Validation. When a user enters an invalid value, they:
- A) See a red triangle in the cell corner
- B) Get a pop-up error dialog with your message
- C) The cell turns red
- D) The value is silently deleted

**9.** Which Scenario Manager feature lets you compare multiple what-if scenarios side by side?
- A) Scenario Summary
- B) Data Table
- C) Goal Seek
- D) Solver

**10.** You want to filter a PivotTable to show only the top 5 products by sales. How?
- A) Sort descending and delete the bottom rows
- B) Click the Row Labels dropdown > Value Filters > Top 10... then set to Top 5
- C) Use a slicer
- D) Create a separate PivotTable for top 5

**11.** What happens when you add a Slicer to a PivotTable?
- A) It replaces the PivotTable filters
- B) It creates a visual filter panel that can be clicked to filter the PivotTable
- C) It exports the data to a new sheet
- D) It creates a chart

**12.** You have a 2-variable Data Table where the row input cell is interest rate and the column input cell is loan amount. The table shows:
- A) One result per combination of rate and amount
- B) Only the interest rate results
- C) Only the loan amount results
- D) An error unless you use Solver

**13.** You want a data validation dropdown that shows options from a named range called `DeptList`. Which source formula?
- A) `=DeptList`
- B) `=INDIRECT("DeptList")`
- C) You type the range `=$H$1:$H$10`
- D) Both A and B work

**14.** What does "Refresh" do on a PivotTable?
- A) Deletes and recreates the PivotTable
- B) Updates the PivotTable with any changes made to the source data
- C) Clears all filters
- D) Resets the layout to default

**15.** Solver differs from Goal Seek because Solver:
- A) Can optimize with multiple constraints and variables
- B) Only works with one variable
- C) Is built into Excel without add-in
- D) Only finds minimum values

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

**16.** Write a Data Validation custom formula that only allows entries greater than the value in cell A1.

> **Your answer:** `_________________________________`

**17.** You have a PivotTable named `PivotTable1`. Write the GETPIVOTDATA formula to extract the sales amount for Region "East".

> **Your answer:** `_________________________________`

**18.** Describe the steps to create a 1-variable Data Table that shows how monthly payments change for different interest rates (rates in column starting at A2, PMT formula in A1).

> **Your answer:** `_________________________________`

**19.** Write a formula using SUBTOTAL that calculates the average of visible cells in B2:B100 (accounting for filtered rows).

> **Your answer:** `_________________________________`

**20.** You want to add a calculated field to a PivotTable that divides Sales by Quantity to get Unit Price. Describe the steps.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | PivotTables are designed exactly for this cross-tabulation. |
| 2 | **A** | Filters hide rows but don't change row numbers — you see gaps. |
| 3 | **A** | Numeric fields default to SUM in PivotTable Values. |
| 4 | **B** | Data Validation restricts input; conditional formatting only visualizes. |
| 5 | **A** | Goal Seek adjusts one input to reach a target output. |
| 6 | **B** | Multi-level sorts: primary sort first, secondary sort within ties. |
| 7 | **B** | "Show Values As > % of Grand Total" is built into PivotTables. |
| 8 | **B** | Data Validation shows a customizable error alert dialog. |
| 9 | **A** | Scenario Summary creates a comparison table of all scenarios. |
| 10 | **B** | Value Filters > Top 10 lets you specify top N by any measure. |
| 11 | **B** | Slicers are visual, clickable filter buttons. |
| 12 | **A** | A 2-variable data table shows one result per row×column combination. |
| 13 | **D** | Both direct name and INDIRECT work as Data Validation sources. |
| 14 | **B** | Refresh pulls in updated source data without recreating the PivotTable. |
| 15 | **A** | Solver handles multiple variables, constraints, and optimization types. |
| 16 | `=B1>A1` | Apply to the validation range; B1 is the first cell in the range being validated. |
| 17 | `=GETPIVOTDATA("Sales",PivotTable1,"Region","East")` | Extracts a specific value from a PivotTable. |
| 18 | Select range containing PMT formula + rate column → Data tab → What-If Analysis → Data Table → Column input cell = the interest rate cell used in the PMT formula. Leave Row input cell blank. | — |
| 19 | `=SUBTOTAL(1,B2:B100)` | Function 1 = AVERAGE; SUBTOTAL ignores filtered-out rows. |
| 20 | PivotTable Analyze tab → Fields, Items & Sets → Calculated Field → Name: "Unit Price" → Formula: `=Sales/Quantity` → OK. | — |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 5
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 4
