# Excel PivotTable Cheat Sheet — Quick Reference

> **Printable Reference** | Print on A4/Letter | 1-2 pages

---

## CREATING A PIVOTTABLE

| Step | Action |
|------|--------|
| 1 | Click anywhere in your data range |
| 2 | **Insert → PivotTable** (or `Alt + N + V`) |
| 3 | Confirm data range and choose location (New/Existing Worksheet) |
| 4 | Click **OK** — empty PivotTable + Field List appears |

> **Prerequisite**: Data must have headers in row 1, no blank rows/columns, no merged cells.

---

## PIVOTTABLE FIELD LIST LAYOUT

The 4 field areas determine how data is displayed:

| Area | Position | What Goes Here | Example |
|------|----------|---------------|---------|
| **Filters** | Top of report | Field to filter entire table | Year, Region |
| **Columns** | Across top | Field as column headers | Quarter, Product |
| **Rows** | Down left side | Field as row labels | Department, Salesperson |
| **Values** | Center (data) | Field to calculate (sum, count, etc.) | Revenue, Units |

### Drag & Drop Methods
- **Drag** fields from field list to areas
- **Check** a field → Excel auto-assigns (text → Rows, numbers → Values)
- **Right-click** field → Move to → choose area

---

## VALUE FIELD SETTINGS

Right-click any value → **Value Field Settings**:

| Calculation | Use When | Example |
|------------|----------|---------|
| **Sum** | Totalling numeric data | Total revenue |
| **Count** | Counting entries | Number of orders |
| **Average** | Finding mean values | Average order value |
| **Max** | Finding highest value | Highest sale |
| **Min** | Finding lowest value | Lowest price |
| **Count Distinct** *(M365)* | Counting unique values | Unique customers |
| **% of Grand Total** | Showing proportion | % of total sales |
| **% of Column Total** | Column-wise percentage | % within each quarter |
| **Running Total** | Cumulative sum | Year-to-date total |
| **Difference From** | Comparing to a base | Change from prior year |
| **Index** | Relative importance | Cross-tabulation weight |

### Changing Calculation
1. Right-click value cell → **Summarize Values By** → select
2. Or: **Value Field Settings → Summarize Values By** tab

### Number Formatting
- Right-click value → **Number Format** (not Format Cells!)
- This ensures formatting persists after pivot refresh

---

## GROUPING DATA

### Group Dates
1. Right-click any date in Rows/Columns
2. Select **Group**
3. Choose: **Years, Quarters, Months, Days, Hours, Minutes**
4. Can select multiple levels (e.g., Years + Months)

### Group Numbers
1. Right-click any number in Rows
2. Select **Group**
3. Set **Starting at**, **Ending at**, and **By** (interval)

### Group Text (Manual)
1. Select items to group (hold `Ctrl`)
2. Right-click → **Group**
3. Rename the group as needed

### Ungroup
Right-click grouped items → **Ungroup**

---

## CALCULATED FIELDS & ITEMS

### Calculated Field (new field using existing fields)
1. Click PivotTable → **PivotTable Analyze → Fields, Items & Sets → Calculated Field**
2. Name: e.g., `Profit`
3. Formula: `=Revenue - Cost`
4. Click **Add → OK**

| Example Calculated Fields | Formula |
|--------------------------|---------|
| Profit Margin | `=Profit / Revenue` |
| Tax | `=Revenue * 0.08` |
| Avg per Unit | `=Revenue / Units` |

### Calculated Item (new item within an existing field)
1. Click a field label → **Fields, Items & Sets → Calculated Item**
2. Name: e.g., `Total Q1+Q2`
3. Formula: `= Q1 + Q2`

> ⚠️ Calculated items can cause issues with grouping. Use with caution.

---

## SLICERS (Visual Filters)

### Add a Slicer
1. Click PivotTable → **PivotTable Analyze → Insert Slicer**
2. Check fields you want as slicers → **OK**
3. Drag slicers to position them on the worksheet

### Using Slicers
| Action | How |
|--------|-----|
| Select single item | Click the item |
| Select multiple items | `Ctrl + Click` items |
| Clear filter | Click the **clear filter** icon (⊘) in slicer header |
| Multi-select mode | Click the **multi-select** icon (☐) in header |

### Connect Slicer to Multiple PivotTables
1. Right-click slicer → **Report Connections** (or **PivotTable Connections**)
2. Check all PivotTables that should share this slicer

### Slicer Formatting
- **Slicer tab** → choose styles, columns, button size
- Right-click → **Slicer Settings** for sorting and hiding items

---

## REFRESHING PIVOTTABLES

| Action | How |
|--------|-----|
| Refresh one PivotTable | Right-click → **Refresh** |
| Refresh all PivotTables | **PivotTable Analyze → Refresh → Refresh All** |
| Auto-refresh on open | **PivotTable Analyze → Options → Data → Refresh data when opening the file** |
| Change data source | **PivotTable Analyze → Change Data Source** |
| Use Tables as source | Convert data to **Table** (`Ctrl+T`) first — auto-expands on refresh |

---

## PIVOTTABLE OPTIONS & TIPS

| Setting | Where | What It Does |
|---------|-------|-------------|
| Classic layout | Options → Display → **Classic PivotTable layout** | Drag fields directly into grid |
| Repeat item labels | Design → Report Layout → **Repeat All Item Labels** | Fills blank row labels |
| Remove subtotals | Design → Subtotals → **Do Not Show Subtotals** | Cleaner look |
| Show in Tabular Form | Design → Report Layout → **Show in Tabular Form** | One column per field |
| Conditional formatting | Home → Conditional Formatting (on pivot cells) | Visual cues on values |
| Drill down | **Double-click** a value cell | Creates new sheet with underlying data |

---

## 📋 Quick Workflow Tips

1. **Start with clean data** → convert to Table (`Ctrl+T`) first
2. **Rows** = categories going down, **Columns** = categories going across
3. **Right-click** is your best friend — nearly every setting is accessible from context menu
4. **Number Format** via right-click pivot value (not Format Cells) — survives refresh
5. **Group dates** immediately after adding — saves dozens of individual date entries
6. **Slicers > Filters** — more visual and easier for end users
7. **Double-click any number** in a pivot to see the raw data behind it
8. **PivotTable Analyze → Options → Totals & Filters** → uncheck "Show grand totals" to remove
