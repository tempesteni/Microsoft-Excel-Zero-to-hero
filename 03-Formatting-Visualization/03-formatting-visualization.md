# Module 3: Formatting & Visualization in Microsoft Excel

## Table of Contents

1. [Cell Formatting: Font, Size, Color, Borders, Fill Colors](#1-cell-formatting)
2. [Number Formats](#2-number-formats)
3. [Format Painter and Cell Styles](#3-format-painter-and-cell-styles)
4. [Conditional Formatting](#4-conditional-formatting)
5. [Custom Conditional Formatting with Formulas](#5-custom-conditional-formatting-with-formulas)
6. [Tables](#6-tables)
7. [Charts](#7-charts)
8. [Chart Elements](#8-chart-elements)
9. [Chart Formatting and Design Best Practices](#9-chart-formatting-and-design-best-practices)
10. [Sparklines](#10-sparklines)
11. [Slicers](#11-slicers)
12. [Dashboard Design Principles](#12-dashboard-design-principles)
13. [Practice Exercises](#13-practice-exercises)
14. [Video References](#14-video-references)
15. [Sources](#15-sources)

---

## 1. Cell Formatting

### 1.1 Font Formatting

Font formatting controls the appearance of text in cells. Access via **Home → Font** group or **Ctrl+1** (Format Cells dialog).

| Property | Description | Shortcut |
|----------|-------------|----------|
| **Font Family** | Typeface (Calibri, Arial, Times New Roman, etc.) | - |
| **Font Size** | Point size (default: 11pt) | - |
| **Bold** | Makes text thicker | **Ctrl+B** |
| **Italic** | Slants text | **Ctrl+I** |
| **Underline** | Adds underline | **Ctrl+U** |
| **Strikethrough** | Draws line through text | Ctrl+1 → Font → Strikethrough |
| **Font Color** | Text color | Home → Font Color dropdown |
| **Superscript/Subscript** | Raises/lowers text | Ctrl+1 → Font → Superscript/Subscript |

**Example:**
```
To make cell A1 bold with a 14pt Arial font in dark blue:
1. Select cell A1
2. Home → Font → Arial
3. Home → Font Size → 14
4. Ctrl+B (bold)
5. Home → Font Color → Dark Blue
```

### 1.2 Borders

Borders add lines around cells to separate and organize data.

**Access:** Home → Borders dropdown or Ctrl+1 → Border tab

| Border Type | Description |
|-------------|-------------|
| **Bottom Border** | Line below selected cells |
| **Top Border** | Line above selected cells |
| **Left Border** | Line on left side |
| **Right Border** | Line on right side |
| **All Borders** | Lines around every cell in selection |
| **Outside Borders** | Line around the outside of the selection |
| **Thick Box Border** | Thick line around the outside |
| **Bottom Double Border** | Double line below (accounting style) |

**Custom Borders (Ctrl+1 → Border tab):**
- Choose line style (thin, thick, double, dashed, dotted)
- Choose line color
- Click specific border positions in the preview area
- Use presets: None, Outline, Inside

**Example:**
```
To create a formatted header row:
1. Select row 1 (A1:E1)
2. Ctrl+1 → Border tab
3. Choose thick line style
4. Click "Outline" preset
5. Click OK
6. Apply Bold + Fill Color
```

### 1.3 Fill Colors (Background)

Fill colors change the background color of cells.

**Access:** Home → Fill Color dropdown

**Methods:**
1. **Quick Fill:** Home → Fill Color → select color
2. **More Colors:** Fill Color → More Colors → Custom tab (enter RGB values)
3. **Format Cells dialog:** Ctrl+1 → Fill tab → choose color
4. **Pattern Fill:** Ctrl+1 → Fill tab → Pattern Color and Pattern Style

**Pro Tip:** Use `Ctrl+1` → Fill tab to access pattern fills (diagonal stripes, dots, etc.) for printed reports.

---

## 2. Number Formats

Number formats change how numbers appear without changing the underlying value.

### 2.1 Built-in Number Formats

| Format | Input | Display | Description |
|--------|-------|---------|-------------|
| **General** | 1234.5 | 1234.5 | Default, no specific format |
| **Number** | 1234.5 | 1,234.50 | Thousands separator, 2 decimals |
| **Currency** | 1234.5 | $1,234.50 | Currency symbol, thousands separator |
| **Accounting** | 1234.5 | $ 1,234.50 | Currency aligned, parentheses for negatives |
| **Date** | 44927 | 1/1/2023 | Serial number to date |
| **Time** | 0.5 | 12:00:00 PM | Decimal to time |
| **Percentage** | 0.75 | 75% | Multiplied by 100, % symbol |
| **Fraction** | 0.75 | 3/4 | Decimal to fraction |
| **Scientific** | 1234.5 | 1.23E+03 | Scientific notation |
| **Text** | 1234.5 | 1234.5 | Treated as text |

**Access:** Home → Number group dropdown, or Ctrl+1 → Number tab

### 2.2 Custom Number Formats

Custom formats use special codes to control display. Access via **Ctrl+1 → Number → Custom**.

**Format Code Structure:** `positive;negative;zero;text`

#### Common Format Codes

| Code | Description | Example Input | Display |
|------|-------------|---------------|---------|
| `0` | Digit placeholder (shows zeros) | 45.2 | 45 |
| `#` | Digit placeholder (hides zeros) | 45.2 | 45 |
| `?` | Digit placeholder (aligns decimals) | 45.2 | 45.2 |
| `.` | Decimal point | `0.00` | 45.20 |
| `,` | Thousands separator | `#,##0` | 1,234 |
| `%` | Percentage (multiplies by 100) | `0%` | 75% |
| `$` | Dollar sign | `$#,##0` | $1,234 |
| `*` | Repeat character to fill cell | `@ *-` | text---- |
| `_` | Skip width of next character | `_( )` | space for alignment |
| `"text"` | Literal text | `0.0 "ft"` | 45.2 ft |
| `@` | Text placeholder | `@ is text` | Hello is text |
| `[Color]` | Apply color | `[Red]#,##0` | 1,234 in red |

#### Practical Custom Format Examples

| Format Code | Input | Display | Use Case |
|-------------|-------|---------|----------|
| `00000` | 41 | 00041 | Zip codes, product codes |
| `#,##0.00` | 1234567 | 1,234,567.00 | Financial reports |
| `$#,##0.00_);($#,##0.00)` | -500 | ($500.00) | Accounting |
| `0.0,, "M"` | 2500000 | 2.5 M | Millions display |
| `#,##0;[Red]-#,##0` | -500 | -500 in red | Red negatives |
| `0.0%` | 0.856 | 85.6% | Percentage with 1 decimal |
| `mmmm d, yyyy` | 44927 | January 1, 2023 | Full date |
| `hh:mm AM/PM` | 0.75 | 6:00 PM | Time display |
| `[Green]$#,##0;[Red]$(#,##0);"Zero"` | 0 | Zero | Multi-condition |

**Creating a Custom Format:**
```
1. Select the cell(s)
2. Press Ctrl+1
3. Click Number tab → Custom
4. In the Type box, enter your format code
5. Preview in Sample box
6. Click OK
```

---

## 3. Format Painter and Cell Styles

### 3.1 Format Painter

The Format Painter copies formatting from one cell and applies it to another.

**How to Use:**
```
1. Select the cell with the formatting you want to copy
2. On the Home tab, in the Clipboard group, click Format Painter
   (A moving dashed border appears around the source cell)
3. Click the destination cell (or drag across a range)
4. Formatting is applied to the destination
```

**Double-Click Trick:** Double-click the Format Painter button to keep it active. You can then apply the same formatting to multiple cells/ranges. Press **Esc** or click Format Painter again to deactivate.

**What Format Painter Copies:**
- Font formatting (type, size, bold, italic, color)
- Number format
- Alignment
- Borders
- Fill color
- Cell protection settings

### 3.2 Cell Styles

Cell Styles are predefined formatting combinations that you can apply with one click.

**Access:** Home → Styles group → Cell Styles

**Built-in Style Categories:**
- **Good, Bad, Neutral** - green, red, gray backgrounds
- **Data and Model** - input, calculation, check cell, linked cell
- **Titles and Headings** - Heading 1-4, Title, Total
- **Themed Cell Styles** - Accent 1-6 (20% and 40% lighter/darker variants)
- **Number Format** - Comma, Currency, Percent

**Creating a Custom Cell Style:**
```
1. Home → Styles → Cell Styles → New Cell Style
2. Enter a Style name (e.g., "My Header")
3. Click Format button
4. Set Number, Alignment, Font, Border, Fill, Protection
5. Uncheck any categories you don't want to control
6. Click OK twice
7. Your style now appears in the Cell Styles gallery
```

**Modifying a Style:** Right-click any cell style → Modify. Changes apply to ALL cells using that style across the workbook.

**Merging Styles:** You can import styles from another open workbook:
```
1. Open both workbooks
2. Home → Styles → Cell Styles → Merge Styles
3. Select the source workbook
4. Click OK
```

---

## 4. Conditional Formatting

Conditional formatting automatically applies formatting based on cell values or formulas.

**Access:** Home → Styles → Conditional Formatting

### 4.1 Highlight Cells Rules

Highlight cells that meet specific criteria:

| Rule | Description | Example |
|------|-------------|---------|
| **Greater Than** | Values above a threshold | Highlight sales > $1000 |
| **Less Than** | Values below a threshold | Highlight inventory < 10 |
| **Between** | Values within a range | Highlight scores between 80-100 |
| **Equal To** | Exact match | Highlight status = "Complete" |
| **Text That Contains** | Cells containing specific text | Highlight cells with "Urgent" |
| **A Date Occurring** | Date-based rules | Highlight dates in "Last Week" |
| **Duplicate Values** | Find duplicates or unique values | Highlight duplicate entries |

**Example: Highlight cells greater than 80**
```
1. Select the range A1:A10
2. Home → Conditional Formatting → Highlight Cells Rules → Greater Than
3. Enter: 80
4. Choose formatting style (e.g., Light Red Fill with Dark Red Text)
5. Click OK
```

### 4.2 Top/Bottom Rules

| Rule | Description | Example |
|------|-------------|---------|
| **Top 10 Items** | Highlight highest N values | Top 10 sales figures |
| **Top 10 %** | Highlight top N percentage | Top 10% of performers |
| **Bottom 10 Items** | Highlight lowest N values | Bottom 10 scores |
| **Bottom 10 %** | Highlight bottom N percentage | Bottom 10% of results |
| **Above Average** | Values above the mean | Above-average grades |
| **Below Average** | Values below the mean | Below-average grades |

**Example: Highlight above average**
```
1. Select the range A1:A10
2. Home → Conditional Formatting → Top/Bottom Rules → Above Average
3. Choose formatting style
4. Click OK
5. Excel calculates the average (e.g., 42.5) and highlights cells above it
```

### 4.3 Data Bars

Data bars add a visual bar inside each cell, proportional to the value.

```
1. Select the range
2. Home → Conditional Formatting → Data Bars
3. Choose a color style (gradient or solid fill)
4. Bars automatically size relative to min/max values
```

**Data Bar Options (More Rules):**
- Minimum/Maximum type (Lowest/Highest value, Number, Percent, Percentile, Formula)
- Bar color and fill (solid, gradient)
- Show bar only (hide the number)
- Bar direction (left-to-right, right-to-left)
- Negative value handling (separate color, axis)

### 4.4 Color Scales

Color scales apply a two-color or three-color gradient based on cell values.

```
1. Select the range
2. Home → Conditional Formatting → Color Scales
3. Choose a preset:
   - Green-Yellow-Red (3-color)
   - Green-White-Red (3-color)
   - Red-Yellow-Green (3-color)
   - White-Red (2-color)
   - Green-White (2-color)
   - Green-Yellow (2-color)
```

**Use Cases:**
- Heatmaps: Red = high values, Green = low values
- Performance tracking: Color intensity shows relative performance
- Geographic data: Color variation by region/metric

### 4.5 Icon Sets

Icon sets add small icons (arrows, shapes, indicators) to cells based on values.

```
1. Select the range
2. Home → Conditional Formatting → Icon Sets
3. Choose a style:
   - Directional: Arrows (3, 4, or 5 arrows)
   - Shapes: Traffic lights, circles, flags
   - Indicators: Check marks, crosses, exclamation
   - Ratings: Stars, ratings
```

**Customizing Icon Sets (More Rules):**
```
1. Home → Conditional Formatting → Icon Sets → More Rules
2. Set Type to "Percent" or "Number"
3. Set values for each icon threshold:
   - Green arrow: >= 67%
   - Yellow arrow: >= 33%
   - Red arrow: < 33%
4. Reverse icon order if needed
5. Show icon only (hide the number)
```

### 4.6 Managing Conditional Formatting Rules

```
1. Home → Conditional Formatting → Manage Rules
2. The Rules Manager shows all rules for the current selection
3. You can:
   - Add new rules
   - Edit existing rules
   - Delete rules
   - Reorder rules (higher rules take priority)
   - Change the "Applies to" range
```

**Rule Priority:** Rules are evaluated top-to-bottom. Check "Stop If True" to prevent lower-priority rules from applying.

---

## 5. Custom Conditional Formatting with Formulas

Use formulas to create powerful, flexible conditional formatting rules.

**Access:** Home → Conditional Formatting → New Rule → "Use a formula to determine which cells to format"

### Key Rules for Formula-Based Conditional Formatting:
1. The formula must evaluate to **TRUE** or **FALSE**
2. Write the formula for the **upper-left cell** in the selected range
3. Excel automatically copies the formula to other cells (relative references adjust)
4. Use absolute ($) references to lock specific cells

### Common Formula Examples

#### Highlight Entire Row Based on a Cell Value
```
Formula: =$C2="Completed"
Applied to: $A$2:$F$100

Explanation: The $ before C locks the column. As Excel moves
across columns, it still checks column C. The row reference
(2) is relative, so it adjusts for each row.
```

#### Highlight Alternating Rows (Zebra Striping)
```
Formula: =MOD(ROW(),2)=0
Applied to: $A$1:$F$100

Explanation: MOD(ROW(),2)=0 returns TRUE for even rows.
Apply a light gray fill for readability.
```

#### Highlight Cells with Errors
```
Formula: =ISERROR(A1)
Applied to: $A$1:$F$100

Explanation: ISERROR returns TRUE for #N/A, #VALUE!, #REF!, etc.
```

#### Highlight Dates Older Than 30 Days
```
Formula: =TODAY()-A1>30
Applied to: $A$1:$A$100

Explanation: Calculates the difference between today and the date.
```

#### Highlight Top N Values
```
Formula: =A1>=LARGE($A$1:$A$100,5)
Applied to: $A$1:$A$100

Explanation: LARGE returns the 5th largest value.
Cells >= that value get formatted.
```

#### Highlight Cells Based on Another Cell
```
Formula: =$B1>$C$1
Applied to: $A$1:$A$100

Explanation: Compares each cell in column A to a
threshold value in cell C1.
```

#### Highlight Weekends in a Date Column
```
Formula: =OR(WEEKDAY(A1,2)=6, WEEKDAY(A1,2)=7)
Applied to: $A$1:$A$31

Explanation: WEEKDAY with return_type 2 returns
6 for Saturday and 7 for Sunday.
```

#### Dynamic Conditional Formatting with Data Validation
```
Formula: =A1>=$E$1
Applied to: $A$1:$A$100

Where E1 contains a dropdown (Data Validation list).
Change the dropdown to dynamically update formatting.
```

---

## 6. Tables

Excel Tables are structured data ranges with built-in features for sorting, filtering, formatting, and formulas.

### 6.1 Creating a Table

```
1. Click any cell inside your data
2. On the Insert tab, in the Tables group, click Table
   (or press Ctrl+T)
3. Excel auto-detects the data range
4. Check "My table has headers" if your first row contains headers
5. Click OK
```

**Keyboard Shortcut:** `Ctrl+T` creates a table instantly.

### 6.2 Table Features

| Feature | Description |
|---------|-------------|
| **Auto-filter** | Drop-down filters in every header |
| **Banding** | Alternating row colors for readability |
| **Total Row** | Built-in SUBTOTAL calculations |
| **Auto-expand** | Table grows automatically when you add data |
| **Structured References** | Use column names in formulas |
| **Calculated Columns** | Formula fills down automatically |
| **Table Name** | Reference the entire table by name |

### 6.3 Structured References

Tables use structured references instead of cell addresses.

| Reference Style | Example | Description |
|-----------------|---------|-------------|
| **Column Name** | `[@Sales]` | Current row's Sales value |
| **Full Column** | `[Sales]` | Entire Sales column |
| **Table + Column** | `Table1[Sales]` | Sales column from Table1 |
| **Table + Column + @** | `Table1[@Sales]` | Current row's Sales in Table1 |
| **#Headers** | `Table1[#Headers]` | Header row of Table1 |
| **#All** | `Table1[#All]` | Entire table including headers |
| **#Data** | `Table1[#Data]` | Data rows only (no header) |
| **#Totals** | `Table1[#Totals]` | Total row only |

**Example Formulas:**
```
=SUM(Table1[Sales])              Sum of all Sales
=AVERAGE(Table1[Amount])         Average of Amount column
=COUNTA(Table1[Product])         Count of Product entries
=Table1[@Price]*Table1[@Qty]     Price × Quantity in current row
```

### 6.4 Table Styles

```
1. Click inside the table
2. On the Table Design tab, in the Table Styles group:
   - Choose from Light, Medium, or Dark styles
   - Hover to preview
   - Click to apply
3. In Table Style Options:
   ☑ Header Row     - Show/hide header formatting
   ☑ Total Row      - Show/hide total row
   ☑ Banded Rows    - Alternating row colors
   ☑ First Column   - Bold first column
   ☑ Last Column    - Bold last column
   ☑ Banded Columns - Alternating column colors
   ☑ Filter Button  - Show/hide filter dropdowns
```

### 6.5 Total Row

```
1. Click inside the table
2. Table Design → Table Style Options → check Total Row
   (or press Ctrl+Shift+T)
3. Click any cell in the total row
4. Choose a function from the dropdown:
   - Average, Count, Count Numbers, Max, Min, Sum
   - StdDev, Var, More Functions...
```

**Important:** The Total Row uses the SUBTOTAL function (not SUM), which correctly handles filtered data.

### 6.6 Converting Table Back to Range

```
1. Click inside the table
2. Table Design → Tools → Convert to Range
3. Click Yes
```

---

## 7. Charts

### 7.1 Chart Types Overview

| Chart Type | Best For | Example Use |
|------------|----------|-------------|
| **Column** | Comparing values across categories | Sales by product |
| **Bar** | Comparing values (horizontal) | Rankings, long category names |
| **Line** | Trends over time | Monthly revenue trend |
| **Pie** | Part-to-whole relationships | Market share (≤7 slices) |
| **Scatter (XY)** | Correlation between two variables | Height vs. weight |
| **Combo** | Mixed data types on different axes | Revenue (columns) + growth % (line) |
| **Area** | Volume over time | Cumulative sales |
| **Stock** | High-Low-Close financial data | Stock prices |
| **Radar** | Comparison across categories | Feature comparison |
| **Treemap** | Hierarchical data | Sales by region/category |
| **Sunburst** | Multi-level hierarchy | Budget allocation |

### 7.2 Creating a Chart

```
1. Select your data range (including headers)
2. On the Insert tab, in the Charts group:
   - Click Recommended Charts to see Excel's suggestions
   - Or click a specific chart type dropdown
3. Choose a chart subtype
4. The chart appears on your worksheet
5. Use Chart Design and Format tabs to customize
```

### 7.3 Column Charts

**Clustered Column:** Compare values across categories
```
Data Layout:
| Product   | Q1    | Q2    | Q3    | Q4    |
|-----------|-------|-------|-------|-------|
| Widgets   | 1200  | 1500  | 1300  | 1800  |
| Gadgets   | 800   | 900   | 1100  | 1200  |
| Gizmos    | 500   | 600   | 700   | 900   |
```

**Stacked Column:** Show part-to-whole across categories
**100% Stacked Column:** Compare proportions (each column = 100%)

### 7.4 Bar Charts

Bar charts are horizontal columns - ideal when:
- Category labels are long
- You're comparing rankings
- You want a different visual style

### 7.5 Line Charts

**Best for:** Showing trends over time

```
Data Layout:
| Month    | Sales  | Expenses |
|----------|--------|----------|
| Jan      | 10000  | 8000     |
| Feb      | 12000  | 8500     |
| Mar      | 11000  | 9000     |
| Apr      | 15000  | 8200     |
```

**Subtypes:**
- **Line:** Basic trend
- **Line with Markers:** Points at each data value
- **Stacked Line:** Cumulative trends
- **100% Stacked Line:** Proportional trends

### 7.6 Pie Charts

**Best for:** Showing composition (part-to-whole)

**Rules:**
- Use for ≤7 slices (more becomes hard to read)
- All values should be positive
- Slices should sum to a meaningful whole
- Consider a Bar/Column chart if slices are similar in size

**Subtypes:**
- **Pie:** Standard 2D
- **3D Pie:** Adds depth (avoid - distorts perception)
- **Pie of Pie:** Separates small slices into a second pie
- **Bar of Pie:** Separates small slices into a bar chart
- **Doughnut:** Multiple rings for multiple data series

### 7.7 Scatter (XY) Charts

**Best for:** Showing correlation between two variables

```
Data Layout:
| Height | Weight |
|--------|--------|
| 60     | 120    |
| 65     | 150    |
| 70     | 180    |
| 72     | 200    |
```

**Use Cases:**
- Correlation analysis
- Regression visualization
- Scientific data
- Outlier detection

### 7.8 Combo Charts

Combo charts combine two or more chart types on the same axes.

```
Example: Sales (columns) + Profit Margin (line on secondary axis)

1. Select data
2. Insert → Recommended Charts → All Charts → Combo
3. For each data series, choose:
   - Chart type (Column, Line, Area, etc.)
   - Which axis (Primary or Secondary)
4. Click OK
```

**Common Combo:** Clustered Column + Line with Secondary Axis

---

## 8. Chart Elements

### 8.1 Adding Chart Elements

```
1. Click the chart to select it
2. Click the + (plus) icon next to the chart
3. Check/uncheck elements:
   ☑ Chart Title
   ☑ Axis Titles
   ☑ Legend
   ☑ Data Labels
   ☑ Data Table
   ☑ Gridlines
   ☑ Trendline
   ☑ Error Bars
   ☑ Up/Down Bars
   ☑ Lines (Drop Lines, High-Low Lines)
```

### 8.2 Chart Title

```
1. Click the + icon → Chart Title
2. Click the title text to edit
3. Options:
   - Above Chart (default)
   - Centered Overlay (overlaps chart)
   - Use formula: Click title → type = → click cell
```

### 8.3 Axis Titles

```
1. Click + → Axis Titles
2. Click each axis title to edit
3. For 3D charts: also a Depth Axis title
```

### 8.4 Legend

```
1. Click + → Legend
2. Positions: Top, Bottom, Left, Right, Top Right
3. Click and drag to reposition
4. Right-click legend → Select Data to control which series appear
```

### 8.5 Data Labels

```
1. Click + → Data Labels
2. Positions: Center, Inside End, Inside Base, Outside End
3. Right-click labels → Format Data Labels:
   ☑ Value
   ☑ Category Name
   ☑ Series Name
   ☑ Percentage (for pie charts)
   ☑ Separator (comma, semicolon, newline)
```

### 8.6 Trendlines

```
1. Click the data series in the chart
2. Click + → Trendline
3. Choose type:
   - Linear: Straight line (y = mx + b)
   - Exponential: Growth curve (y = ae^(bx))
   - Logarithmic: Log curve (y = a*ln(x) + b)
   - Polynomial: Curve (y = a + bx + cx² + ...)
   - Power: Power curve (y = ax^b)
   - Moving Average: Smoothed trend
4. Right-click trendline → Format Trendline:
   ☑ Display Equation on chart
   ☑ Display R-squared value
   ☑ Set Forecast periods (forward/backward)
```

### 8.7 Error Bars

```
1. Click the data series
2. Click + → Error Bars
3. Options:
   - Standard Error
   - Percentage (e.g., 5%)
   - Standard Deviation (e.g., 1 SD)
   - Fixed value
   - Custom (specify + and - values from cells)
```

---

## 9. Chart Formatting and Design Best Practices

### 9.1 Design Principles

| Principle | Guidance |
|-----------|----------|
| **Simplify** | Remove unnecessary elements (gridlines, borders, 3D effects) |
| **Contrast** | Use contrasting colors for different data series |
| **Consistency** | Same colors/fonts across all charts in a report |
| **Label clearly** | Always include titles and axis labels |
| **Avoid distortion** | Start Y-axis at zero (unless showing small changes) |
| **Right chart type** | Match chart type to the story you're telling |
| **Color-blind friendly** | Don't rely on color alone; use patterns or labels |
| **White space** | Don't overcrowd; let data breathe |

### 9.2 Formatting Tips

```
1. Remove gridlines: Click gridlines → Delete
2. Change colors: Chart Design → Change Colors
3. Format data series: Right-click → Format Data Series
4. Format axis: Right-click axis → Format Axis
   - Set min/max bounds
   - Change number format
   - Adjust tick marks
5. Format plot area: Right-click → Format Plot Area
   - Add/remove border
   - Fill color
6. Add shapes/text: Insert → Shapes/Text Box
```

### 9.3 Common Chart Mistakes to Avoid

| Mistake | Fix |
|---------|-----|
| Using 3D charts | Use 2D for accurate comparison |
| Too many data series | Limit to 4-5 series per chart |
| Pie chart with too many slices | Group small slices into "Other" |
| Y-axis doesn't start at zero | Start at zero unless showing small changes |
| Missing axis labels | Always label both axes |
| Inconsistent colors | Use the same color for the same category |
| No title | Always include a descriptive title |
| Misleading scale | Use consistent intervals |

---

## 10. Sparklines

Sparklines are small, cell-sized charts that show data trends in a compact format.

**Access:** Insert → Sparklines group

### 10.1 Line Sparklines

```
1. Select where you want the sparkline (e.g., E2:E10)
2. Insert → Sparklines → Line
3. Data Range: The values (e.g., A2:D10)
4. Location Range: Where to place sparklines (e.g., E2:E10)
5. Click OK
```

### 10.2 Column Sparklines

```
1. Insert → Sparklines → Column
2. Same process as line sparklines
3. Each column in the sparkline represents one data point
```

### 10.3 Win/Loss Sparklines

```
1. Insert → Sparklines → Win/Loss
2. Shows positive (wins) and negative (losses) values
3. Ideal for: Game scores, profit/loss, +/- changes
```

### 10.4 Customizing Sparklines

```
1. Click the sparkline(s) to select
2. Sparkline Design tab appears:
   - Edit Data: Change data source or location
   - Sparkline Color: Change line/bar color
   - Marker Color: Highlight High Point, Low Point,
     First Point, Last Point, Negative Points, Markers
   - Style: Choose from preset styles
   - Axis: Change min/max bounds, date axis
   - Group: Group/ungroup sparklines for consistent scaling
   - Clear: Remove sparklines
```

**Key Options:**
```
☑ High Point    - Highlights the maximum value
☑ Low Point     - Highlights the minimum value
☑ First Point   - Highlights the first data point
☑ Last Point    - Highlights the last data point
☑ Negative Points - Highlights negative values
☑ Markers       - Shows all data points
```

---

## 11. Slicers

Slicers are visual filter buttons that make it easy to filter data in tables and PivotTables.

### 11.1 Slicers for Tables

```
1. Click inside the table
2. On the Table Design tab, click Insert Slicer
3. Check the columns you want slicers for
4. Click OK
5. Slicers appear as floating panels
6. Click buttons to filter
7. Ctrl+click to select multiple items
8. Click the clear button (×) to remove filter
```

### 11.2 Slicers for PivotTables

```
1. Click inside the PivotTable
2. On the PivotTable Analyze tab, click Insert Slicer
3. Check the fields you want slicers for
4. Click OK
5. Use slicers to filter the PivotTable interactively
```

### 11.3 Formatting Slicers

```
1. Click the slicer to select it
2. Slicer tab appears:
   - Slicer Settings: Name, caption, item sorting
   - Slicer Styles: Choose from preset styles
   - Buttons: Set number of columns, button height/width
   - Size: Set exact dimensions
   - Arrange: Position, alignment, grouping
```

### 11.4 Connecting Slicers to Multiple PivotTables

```
1. Click the slicer
2. On the Slicer tab, click Report Connections
   (or right-click → Report Connections)
3. Check all PivotTables you want connected
4. Click OK
5. Now filtering the slicer affects all connected PivotTables
```

---

## 12. Dashboard Design Principles

### 12.1 What is a Dashboard?

A dashboard is a single-page visual display of the most important information, designed for at-a-glance monitoring.

### 12.2 Design Principles

| Principle | Description |
|-----------|-------------|
| **Know your audience** | Executive vs. operational - adjust detail level |
| **Single page** | Fits on one screen; no scrolling |
| **Key metrics first** | Most important KPIs in the top-left (reading order) |
| **Consistent formatting** | Same fonts, colors, sizes throughout |
| **Minimal text** | Use numbers and visuals, not paragraphs |
| **Color coding** | Red = bad, Yellow = caution, Green = good |
| **Interactive elements** | Slicers, dropdowns, buttons for filtering |
| **White space** | Don't overcrowd - give elements breathing room |
| **Actionable** | Shows what needs attention, not just what happened |

### 12.3 Dashboard Building Blocks in Excel

| Element | Use For |
|---------|---------|
| **PivotTables** | Summarizing large datasets |
| **PivotCharts** | Visual representation of PivotTable data |
| **Sparklines** | Inline trend indicators |
| **Conditional Formatting** | Color-coded KPIs |
| **Slicers** | Interactive filters |
| **Data Validation Dropdowns** | User-controlled parameters |
| **Named Ranges** | Dynamic data sources |
| **CHOOSE/INDEX functions** | Dynamic chart data selection |

### 12.4 Dashboard Layout Tips

```
┌─────────────────────────────────────────────┐
│  KPI 1        KPI 2        KPI 3        KPI 4│
│  $1.2M        847          92%          ↑12% │
├─────────────────────────────────────────────┤
│  [Slicer 1]  [Slicer 2]  [Slicer 3]        │
├────────────────────────┬────────────────────┤
│                        │                    │
│   Chart 1              │   Chart 2          │
│   (Main metric)        │   (Secondary)      │
│                        │                    │
├────────────────────────┼────────────────────┤
│                        │                    │
│   Chart 3              │   Table            │
│   (Trend)              │   (Details)        │
│                        │                    │
└────────────────────────┴────────────────────┘
```

### 12.5 KPI Cards in Excel

Create KPI summary cards using cell formatting:

```
1. Merge cells for the card area
2. Add a subtle border
3. Large font for the KPI value
4. Small font for the KPI label
5. Use conditional formatting for indicators:
   - Green: On target
   - Yellow: Close to target
   - Red: Below target
6. Add a small sparkline for trend
```

---

## 13. Practice Exercises

### Exercise 1: Cell Formatting & Number Formats

**Objective:** Format a sales report with proper number formatting.

**Sample Data:**
| A | B | C | D |
|---|---|---|---|
| Date | Product | Units Sold | Revenue |
| 1/15/2024 | Widget A | 150 | 4500.5 |
| 2/20/2024 | Widget B | 200 | 8000.75 |
| 3/10/2024 | Widget C | 75 | 2250.25 |
| 4/5/2024 | Widget A | 300 | 9000 |
| 5/18/2024 | Widget B | 125 | 5000.5 |

**Tasks:**
1. Format column A as Date (MM/DD/YYYY)
2. Format column D as Currency with 2 decimal places
3. Bold the header row and add a bottom border
4. Apply a fill color to the header row (e.g., dark blue with white text)
5. Auto-fit all columns
6. Apply a custom number format to column D that shows positive values in green and negative in red

### Exercise 2: Conditional Formatting

**Objective:** Apply various conditional formatting rules to highlight important data.

**Sample Data:**
| A | B | C | D |
|---|---|---|---|
| Student | Math | Science | English |
| Alice | 92 | 88 | 95 |
| Bob | 78 | 82 | 71 |
| Carol | 95 | 97 | 89 |
| Dave | 65 | 58 | 72 |
| Eve | 88 | 91 | 84 |
| Frank | 45 | 52 | 61 |

**Tasks:**
1. Highlight cells >90 with green fill, <60 with red fill
2. Apply a color scale (green-yellow-red) to all scores
3. Add data bars to column B (Math scores)
4. Highlight the top 2 scores in each column
5. Create a formula-based rule to highlight entire row if any score < 60
6. Add icon sets (traffic lights) based on: ≥90 green, ≥70 yellow, <70 red

### Exercise 3: Tables & Structured References

**Objective:** Convert data to a table and use structured references.

**Sample Data:**
| A | B | C | D | E |
|---|---|---|---|---|
| OrderID | Product | Category | Price | Quantity |
| 1001 | Laptop | Electronics | 999.99 | 5 |
| 1002 | Desk Chair | Furniture | 249.50 | 10 |
| 1003 | Monitor | Electronics | 399.99 | 8 |
| 1004 | Keyboard | Electronics | 79.99 | 20 |
| 1005 | Bookshelf | Furniture | 149.00 | 3 |

**Tasks:**
1. Convert to a Table (Ctrl+T) named "Orders"
2. Add a Total Row showing the SUM of Quantity
3. Add a calculated column "Total" = Price × Quantity (use structured references)
4. Apply a Medium style table format with banded rows
5. Use the table name in a formula: `=SUM(Orders[Total])`
6. Insert slicers for Category

### Exercise 4: Charts & Visualization

**Objective:** Create and format various chart types.

**Sample Data:**
| A | B | C | D | E |
|---|---|---|---|---|
| Month | Sales | Expenses | Profit | Margin% |
| Jan | 50000 | 35000 | 15000 | 30% |
| Feb | 55000 | 36000 | 19000 | 35% |
| Mar | 48000 | 34000 | 14000 | 29% |
| Apr | 62000 | 38000 | 24000 | 39% |
| May | 70000 | 42000 | 28000 | 40% |
| Jun | 65000 | 40000 | 25000 | 38% |

**Tasks:**
1. Create a clustered column chart for Sales and Expenses
2. Add a line for Profit on a secondary axis (Combo chart)
3. Add a trendline to the Sales series (linear)
4. Format the chart: title, axis labels, legend, data labels
5. Create a pie chart showing expense breakdown by month
6. Add sparklines in column F showing the trend for each metric
7. Apply conditional formatting with data bars to the Profit column

### Exercise 5: Dashboard Building

**Objective:** Create a mini dashboard combining multiple visual elements.

**Steps:**
1. Start with the sample data from Exercise 4
2. Create KPI cards at the top showing:
   - Total Sales (large number)
   - Average Margin %
   - Best Month
   - Worst Month
3. Add a Combo chart (columns + line)
4. Add sparklines for each metric
5. Add a slicer or dropdown for filtering
6. Apply consistent formatting (colors, fonts, spacing)
7. Ensure it fits on one screen

---

## 14. Video References

### Conditional Formatting
- **Excel Conditional Formatting in 10 Minutes** - https://www.youtube.com/watch?v=yVixz6pMk8Q
- **Conditional Formatting with Formulas** - https://www.youtube.com/watch?v=9NUjHBNWe2E
- **Excel Conditional Formatting - Complete Guide** - https://www.youtube.com/watch?v=7Fj4Qb_b7Tw

### Charts & Visualization
- **Excel Charts - The Complete Guide** - https://www.youtube.com/watch?v=ewYorxOGwTk
- **Excel Combo Charts Tutorial** - https://www.youtube.com/watch?v=3k7ZsT3Gk5Y
- **Excel Chart Formatting Tips** - https://www.youtube.com/watch?v=K79UMRWKDF8

### Tables
- **Excel Tables - Complete Guide** - https://www.youtube.com/watch?v=2G4r2r_1T9Y
- **Structured References in Excel Tables** - https://www.youtube.com/watch?v=3b4nF1BFwBg

### Sparklines & Slicers
- **Excel Sparklines Tutorial** - https://www.youtube.com/watch?v=9vJRopuVJ9Q
- **Excel Slicers - Tables & PivotTables** - https://www.youtube.com/watch?v=Zb2h2nUOZuE

### Dashboard Design
- **Excel Dashboard Tutorial** - https://www.youtube.com/watch?v=M2NH_PQuWnE
- **Excel Dashboard Design Tips** - https://www.youtube.com/watch?v=LjWPUqEBZ2E

### Number Formatting
- **Custom Number Formats in Excel** - https://www.youtube.com/watch?v=2f1mSq2sGMk
- **Excel Format Cells Tutorial** - https://www.youtube.com/watch?v=VlIHvXNJKUQ

---

## 15. Sources

### Primary Sources

1. **Microsoft Support - Format Cells**
   https://support.microsoft.com/en-us/excel

2. **Microsoft Support - Conditional Formatting**
   https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f

3. **Microsoft Support - Excel Tables**
   https://support.microsoft.com/en-us/excel

4. **Microsoft Support - Charts**
   https://support.microsoft.com/en-us/excel

5. **Microsoft Support - Sparklines**
   https://support.microsoft.com/en-us/excel

6. **Microsoft Support - Slicers**
   https://support.microsoft.com/en-us/office/use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d

7. **Excel Easy - Format Cells**
   https://www.excel-easy.com/basics/format-cells.html

8. **Excel Easy - Custom Number Format**
   https://www.excel-easy.com/examples/custom-number-format.html

9. **Excel Easy - Format Painter**
   https://www.excel-easy.com/examples/format-painter.html

10. **Excel Easy - Cell Styles**
    https://www.excel-easy.com/examples/cell-styles.html

11. **Excel Easy - Conditional Formatting**
    https://www.excel-easy.com/data-analysis/conditional-formatting.html

12. **Excel Easy - Tables**
    https://www.excel-easy.com/data-analysis/tables.html

13. **Excel Easy - Charts**
    https://www.excel-easy.com/

14. **Exceljet - Conditional Formatting**
    https://exceljet.net/conditional-formatting

15. **Exceljet - Excel Charts**
    https://exceljet.net/charts

16. **Exceljet - Pivot Table Examples**
    https://exceljet.net/pivot-tables

17. **CFI - Conditional Formatting**
    https://corporatefinanceinstitute.com/resources/excel/conditional-formatting/

18. **Chandoo - Excel Charts**
    https://chandoo.org/?s=excel+charts

19. **Contextures - Conditional Formatting**
    https://www.contextures.com/

20. **Peltier Tech Blog - Chart Types**
    https://peltiertech.com/

---

*Last Updated: September 2026*
