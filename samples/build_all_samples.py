#!/usr/bin/env python3
"""Generate sample workbooks 01-06 for the Excel Zero to Hero package.
All data is fictional (YourCompany Inc). Part of the learning material: study this code."""
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.chart import BarChart, LineChart, Reference

OUT = "~/EXCEL Learn/samples"
NAVY = "1F4E79"; SAND = "E7E8D1"; GREY = "D9D9D9"; GREEN = "97BC62"

def header(ws, row, values, start=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start + i, value=v)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center")

def autosize(ws):
    for col in ws.columns:
        w = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(w + 3, 42)

def save(wb, name):
    wb.save(f"{OUT}/{name}")
    print("saved:", name)

# ---------------- sample 01: data entry practice ----------------
wb = Workbook()
ws = wb.active; ws.title = "Instructions"
rows = [
    ["Practice: clean data entry (companion to Module 1)"],
    [""],
    ["1. Open 'Raw-Entry'. Finish the incomplete rows using the patterns above them."],
    ["2. Fix the 3 dates stored as text so Excel recognizes them as real dates."],
    ["3. Make every Order ID match the pattern YC-0001 (zero-padded, 4 digits)."],
    ["4. Open 'Format-Me'. Convert the text-numbers to real numbers (Data > Text to Columns trick)."],
    ["5. Align everything the Excel way: text left, numbers right, headers centered."],
    ["Success check: no green triangles remain in any cell."],
]
for r in rows: ws.append(r)
ws["A1"].font = Font(bold=True, size=14, color=NAVY)
ws2 = wb.create_sheet("Raw-Entry")
header(ws2, 1, ["Order ID", "Date", "Customer", "Product", "Qty", "Unit Price", "Line Total"])
data = [
    ["YC-0001", "2026-01-06", "Northwind Traders", "Widget A", 10, 25.00],
    ["YC-0002", "2026-01-07", "Redwood Labs", "Widget B", 4, 130.50],
    ["YC-3", "03/13/2026", "Harborline Credit Union", "Widget A", 25, 25.00],
    ["YC-0004", "2026-02-02", "Meridian Freight", "Widget C", 2, 480.00],
    ["YC-0005", "2/14/26", "Greenfield Literacy Trust", "Widget B", 6, 130.50],
    ["YC-0006", "2026-02-20", "Northgate Institute", "Widget A", 15, 25.00],
    ["YC-0007", "March 5 2026", "Northwind Traders", "Widget C", 1, 480.00],
    ["YC-0008", "2026-03-11", "VeloCabin", "Widget B", 8, 130.50],
    ["YC-0009", "2026-03-18", "Redwood Labs", "Widget A", None, 25.00],
    ["YC-0010", "2026-03-25", "Meridian Freight", "Widget C", None, 480.00],
    ["YC-0011", "2026-04-02", "Harborline Credit Union", None, 12, None],
    ["YC-0012", "2026-04-08", "Greenfield Literacy Trust", "Widget B", None, 130.50],
]
for row in data: ws2.append(row)
for r in range(2, 2 + len(data)):
    ws2.cell(row=r, column=7, value=f"=E{r}*F{r}")
autosize(ws); autosize(ws2)
ws3 = wb.create_sheet("Format-Me")
header(ws3, 1, ["Item", "Amount (text!)", "Percent (text!)", "Date (text!)"])
fmt = [["Consulting", "1250", "15%", "2026-01-31"], ["Training", "3200", "10%", "2026-02-28"],
       ["Support", "975", "22%", "2026-03-31"], ["Licenses", "15600", "5%", "2026-04-30"]]
for row in fmt: ws3.append(row)
for r in range(2, 6):
    ws3.cell(row=r, column=2).alignment = Alignment(horizontal="right")
save(wb, "sample-01-data-entry-practice.xlsx")

# ---------------- sample 02: formula challenges ----------------
wb = Workbook()
ws = wb.active; ws.title = "Data"
header(ws, 1, ["Region", "Rep", "Product", "Units", "Unit Price", "Revenue", "Commission"])
random.seed(42)
regions = ["North", "South", "East", "West"]; reps = ["Jordan Ade", "Mina Okafor", "Sam Cole", "Riley Chen"]
prods = ["Widget A", "Widget B", "Widget C"]; prices = {"Widget A": 25.0, "Widget B": 130.5, "Widget C": 480.0}
for i in range(40):
    p = random.choice(prods); u = random.randint(1, 30)
    ws.append([random.choice(regions), random.choice(reps), p, u, prices[p], None, None])
ws2 = wb.create_sheet("Challenges")
header(ws2, 1, ["#", "Task (write the formula in Data columns F and G)", "Expected result to match"])
tasks = [
    (1, "Data!F2: revenue for this row (Units * Unit Price), filled down to row 41", "80 rows of revenue"),
    (2, "Data!G2: commission = Revenue * 8% if Units >= 15, else Revenue * 5% (IF), filled down", "Two commission rates"),
    (3, "In I2: total revenue for all rows (SUM)", "Matches SUM of column F"),
    (4, "In I3: average units per order (AVERAGE), 1 decimal", ""),
    (5, "In I4: number of orders with Units > 20 (COUNTIF)", ""),
    (6, "In I5: revenue for Region = North only (SUMIF)", ""),
    (7, "In I6: revenue per region for all 4 regions (SUMIFS grid in I6:L7)", ""),
    (8, "In I8: highest single-order revenue (MAX)", ""),
    (9, "In I9: the Rep name of that highest order (INDEX/MATCH)", ""),
    (10, "In I10: total revenue rounded to whole dollars (ROUND + SUM)", ""),
    (11, "In I11: IFERROR-protected lookup: VLOOKUP 'Widget Z' price from a small table, show 'Not found'", ""),
    (12, "In I12: text join: 'Top rep: ' + the name in I9 (CONCAT)", ""),
]
for t in tasks: ws2.append(list(t))
ws3 = wb.create_sheet("Answer Key")
header(ws3, 1, ["#", "Model formula", "Note"])
answers = [
    (1, "=E2*D2", "Fill down"), (2, '=IF(D2>=15,F2*0.08,F2*0.05)', "IF on Units"),
    (3, "=SUM(F2:F41)", ""), (4, "=ROUND(AVERAGE(D2:D41),1)", ""),
    (5, "=COUNTIF(D2:D41,\">20\")", ""), (6, '=SUMIF(A2:A41,"North",F2:F41)', ""),
    (7, '=SUMIFS($F$2:$F$41,$A$2:$A$41,I$7)', "Region names in I7:L7"),
    (8, "=MAX(F2:F41)", ""), (9, "=INDEX(B2:B41,MATCH(MAX(F2:F41),F2:F41,0))", "Exact match"),
    (10, "=ROUND(SUM(F2:F41),0)", ""), (11, '=IFERROR(VLOOKUP("Widget Z",K2:L4,2,FALSE),"Not found")', "Exact match VLOOKUP"),
    (12, '="Top rep: "&I9', "Concat with &"),
]
for a in answers: ws3.append(list(a))
for s in wb: autosize(s)
save(wb, "sample-02-formula-challenges.xlsx")

# ---------------- sample 03: dashboard before/after ----------------
wb = Workbook()
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
rev = [42000, 51000, 47000, 63000, 58000, 71000]; cost = [31000, 33000, 34000, 36000, 35000, 38000]
ws = wb.active; ws.title = "Before"
ws.append(["YourCompany Inc - H1 2026 (raw export)"]); header(ws, 2, ["Month", "Revenue", "Cost", "Profit"])
for i, m in enumerate(months):
    ws.append([m, rev[i], cost[i], None])
for r in range(3, 9): ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
ws2 = wb.create_sheet("After")
ws2["A1"] = "H1 2026: profit up 6 straight months"
ws2["A1"].font = Font(bold=True, size=16, color=NAVY)
ws2.append([]); header(ws2, 3, ["Month", "Revenue", "Cost", "Profit"])
for i, m in enumerate(months):
    ws2.append([m, rev[i], cost[i], None])
for r in range(4, 10):
    ws2.cell(row=r, column=4, value=f"=B{r}-C{r}")
    for col in (2, 3, 4):
        ws2.cell(row=r, column=col).number_format = '"$"#,##0'
ws2["F3"] = "Total profit"; ws2["F3"].font = Font(bold=True, color=NAVY)
ws2["G3"] = "=SUM(D4:D9)"; ws2["G3"].number_format = '"$"#,##0'
ws2["F4"] = "Best month"; ws2["G4"] = "=INDEX(A4:A9,MATCH(MAX(D4:D9),D4:D9,0))"
ws2.conditional_formatting.add("D4:D9", DataBarRule(start_type="min", end_type="max", color=GREEN))
ws2.conditional_formatting.add("C4:C9", CellIsRule(operator="greaterThan", formula=["36000"],
                                                   fill=PatternFill("solid", fgColor="FFC7CE")))
ch = BarChart(); ch.title = "Revenue vs cost"; ch.y_axis.title = "USD"
ch.add_data(Reference(ws2, min_col=2, max_col=3, min_row=3, max_row=9), titles_from_data=True)
ch.set_categories(Reference(ws2, min_col=1, min_row=4, max_row=9))
ws2.add_chart(ch, "F6")
for s in wb: autosize(s)
save(wb, "sample-03-dashboard-before-after.xlsx")

# ---------------- sample 04: pivottable source data ----------------
wb = Workbook()
ws = wb.active; ws.title = "Sales-Data"
header(ws, 1, ["Date", "Region", "Rep", "Category", "Product", "Units", "Revenue"])
random.seed(7)
cats = {"Office": ["Widget A", "Widget B"], "Tech": ["Widget C", "Widget D"]}
for i in range(300):
    cat = random.choice(list(cats)); p = random.choice(cats[cat])
    d = f"2026-{random.randint(1,6):02d}-{random.randint(1,28):02d}"
    u = random.randint(1, 25)
    ws.append([d, random.choice(regions), random.choice(reps), cat, p, u, u * random.choice([25, 65, 130.5, 480])])
ws2 = wb.create_sheet("Tasks")
for r in [["PivotTable practice (Module 4). Data: 300 clean rows on 'Sales-Data'."], [""],
          ["1. Insert > PivotTable: total Revenue by Region. Which region wins?"],
          ["2. Add Rep as a second row field: who is the top rep in that region?"],
          ["3. Build Revenue by Month (drag Date to rows; Excel groups it)."],
          ["4. Make a PivotChart column chart from task 3."],
          ["5. Add a Slicer on Category and filter to Tech only."],
          ["6. Show values as % of grand total instead of sums."],
          ["Success check: every pivot recalculates when you change one Revenue cell."]]:
    ws2.append(r)
ws2["A1"].font = Font(bold=True, size=13, color=NAVY)
autosize(ws); autosize(ws2)
save(wb, "sample-04-pivottable-source.xlsx")

# ---------------- sample 05: lookup challenges ----------------
wb = Workbook()
ws = wb.active; ws.title = "Products"
header(ws, 1, ["SKU", "Product", "Category", "Price", "Stock"])
for i in range(1, 51):
    ws.append([f"SKU-{i:03d}", f"Widget {chr(65 + i % 4)}-{i}", "Office" if i % 3 else "Tech",
               round(20 + i * 7.5, 2), 50 - i])
ws2 = wb.create_sheet("Orders")
header(ws2, 1, ["Order ID", "SKU", "Qty", "Date", "Product Name (fill)", "Unit Price (fill)", "Order Total (fill)"])
random.seed(11)
for i in range(1, 81):
    ws2.append([f"YC-{i:04d}", f"SKU-{random.randint(1,50):03d}", random.randint(1, 12),
                f"2026-{random.randint(1,6):02d}-{random.randint(1,28):02d}", None, None, None])
ws3 = wb.create_sheet("Challenges")
for r in [["Lookup practice (Module 5). Fill Orders columns E-G using Products."], [""],
          ["1. E2: product name for this SKU with XLOOKUP, filled down."],
          ["2. F2: unit price with XLOOKUP + if_not_found argument 'Discontinued'."],
          ["3. G2: order total = Qty * Unit Price (watch for text results in F!)."],
          ["4. Redo E2 with INDEX/MATCH (exact match) in H2 for comparison."],
          ["5. In J2: count how many orders used a Tech SKU (COUNTIFS + helper col or SUMPRODUCT)."],
          ["6. In J3: total revenue for Office category (SUMIF on the filled Orders)."],
          ["7. In J4: LET formula: average order value with one pass over the data."],
          ["8. Bonus: make task 1 spill for all 80 rows with one dynamic-array formula."]]:
    ws3.append(r)
ws3["A1"].font = Font(bold=True, size=13, color=NAVY)
ws4 = wb.create_sheet("Answer Key")
for r in [["1", '=XLOOKUP(B2,Products!$A:$A,Products!$B:$B)'],
          ["2", '=XLOOKUP(B2,Products!$A:$A,Products!$D:$D,"Discontinued")'],
          ["3", '=IF(ISNUMBER(F2),C2*F2,"check price")'],
          ["4", '=INDEX(Products!$B:$B,MATCH(B2,Products!$A:$A,0))'],
          ["5", '=SUMPRODUCT(--(XLOOKUP(Orders!$B$2:$B$81,Products!$A:$A,Products!$C:$C)="Tech"))'],
          ["6", '=SUMPRODUCT((XLOOKUP(Orders!$B$2:$B$81,Products!$A:$A,Products!$C:$C)="Office")*C2:C81*F2:F81)'],
          ["7", '=LET(p,F2:F81,q,C2:C81,AVERAGE(p*q))'],
          ["8", '=XLOOKUP(B2:B81,Products!$A:$A,Products!$B:$B)']]:
    ws4.append(r)
header(ws4, 1, ["#", "Model formula"])
for s in wb: autosize(s)
save(wb, "sample-05-lookup-challenges.xlsx")

# ---------------- sample 06: power query messy data ----------------
wb = Workbook()
def messy(name, rows, variant):
    s = wb.create_sheet(name)
    heads = [["Date", "Customer Name", "Amount", "Region"],
             ["date ", "Customer name", "AMOUNT ($)", "region"],
             ["Transaction Date", "Client", "Net Amount", "Territory"]][variant]
    header(s, 1, heads)
    for r in rows: s.append(r)
    autosize(s)
    return s
wb.remove(wb.active)
messy("Jan", [
    ["2026-01-05", "Northwind Traders ", 1250, "North"], ["2026-01-09", "Redwood Labs", " 830", "South"],
    ["01/14/2026", "Northwind Traders", 460, "North"], [None, None, None, None],
    ["2026-01-22", "Harborline Credit Union", 2100, " East"], ["2026-01-22", "Harborline Credit Union", 2100, " East"],
    ["2026-01-30", "Meridian Freight", "1,450", "West"], ["2026-01-31", "Redwood Labs", 990, "south"],
], 0)
messy("Feb", [
    ["2/3/2026", "VeloCabin", 725, "north"], ["2026-02-11", " Northgate Institute", 1875, "North"],
    ["2026-02-11", "Northgate Institute", 1875, "North"], [None, None, None, None],
    ["2026-02-18", "Greenfield Literacy Trust", 540, "South"],
    ["Feb 24 2026", "Meridian Freight", 3200, "west"], ["2026-02-27", "Redwood Labs", 815, "South"],
], 1)
messy("Mar", [
    ["2026-03-02", "Northwind Traders", 1990, "North"], ["03/09/2026", "VeloCabin ", 655, "east"],
    ["2026-03-15", "Harborline Credit Union", "2,740", "East"], [None, None, None, None],
    ["2026-03-21", "Meridian Freight", 1120, "West"], ["2026-03-21", "Meridian Freight", 1120, "West"],
    ["2026-03-29", "Greenfield Literacy Trust", 480, "South"],
], 2)
t = wb.create_sheet("Tasks")
for r in [["Power Query bootcamp (Module 6). Three monthly exports, three sets of problems."], [""],
          ["1. Data > Get Data > From Workbook: load all three sheets into Power Query."],
          ["2. Append them into one table 'AllMonths'."],
          ["3. Clean: promote correct headers (they differ per sheet!), trim spaces in names/regions."],
          ["4. Standardize region casing (North/north -> North)."],
          ["5. Fix text-numbers in Amount (' 830', '1,450') to real numbers."],
          ["6. Parse the mixed date formats into one Date type."],
          ["7. Remove blank rows and exact duplicates."],
          ["8. Load the clean result to a new sheet and refresh after editing Jan."]]:
    t.append(r)
t["A1"].font = Font(bold=True, size=13, color=NAVY)
autosize(t)
save(wb, "sample-06-powerquery-messy.xlsx")
print("all sample workbooks generated")
