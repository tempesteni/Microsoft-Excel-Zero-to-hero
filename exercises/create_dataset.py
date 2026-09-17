#!/usr/bin/env python3
"""Create a realistic sales practice dataset for Excel learning."""

import random
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers

random.seed(42)

wb = Workbook()

# ─── Sheet 1: Sales Data ───
ws1 = wb.active
ws1.title = "Sales-Data"

headers = ["Date", "Region", "Salesperson", "Product", "Units Sold", "Unit Price", "Total", "Commission"]
header_font = Font(bold=True, size=11)
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font_white = Font(bold=True, size=11, color="FFFFFF")
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

for col, h in enumerate(headers, 1):
    cell = ws1.cell(row=1, column=col, value=h)
    cell.font = header_font_white
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')
    cell.border = thin_border

regions = ["North", "South", "East", "West"]
salespeople = [
    "Alice Johnson", "Brian Chen", "Carlos Rivera", "Diana Patel",
    "Erik Nguyen", "Fatima Al-Said", "George Kim", "Hannah Okafor",
    "Ian Schwartz", "Jasmine Lee"
]
products = {
    "Laptop": 899.99,
    "Monitor": 349.99,
    "Keyboard": 79.99,
    "Mouse": 29.99,
    "Headset": 149.99,
    "Webcam": 69.99,
}

start_date = datetime.date(2024, 1, 1)
end_date = datetime.date(2024, 12, 31)
date_range_days = (end_date - start_date).days

rows_data = []
for i in range(200):
    day_offset = random.randint(0, date_range_days)
    dt = start_date + datetime.timedelta(days=day_offset)
    region = random.choice(regions)
    person = random.choice(salespeople)
    product = random.choice(list(products.keys()))
    units = random.randint(1, 50)
    price = products[product]
    rows_data.append((dt, region, person, product, units, price))

# Inject intentional issues:
# 1) 3 cells with extra spaces (rows 10, 50, 120 for Salesperson)
rows_data[10] = (rows_data[10][0], rows_data[10][1], "  Alice Johnson ", rows_data[10][3], rows_data[10][4], rows_data[10][5])
rows_data[50] = (rows_data[50][0], " South ", rows_data[50][2], rows_data[50][3], rows_data[50][4], rows_data[50][5])
rows_data[120] = (rows_data[120][0], rows_data[120][1], rows_data[120][2], rows_data[120][3] + "  ", rows_data[120][4], rows_data[120][5])

# 2) 1 blank cell (row 75, Salesperson)
rows_data[75] = (rows_data[75][0], rows_data[75][1], None, rows_data[75][3], rows_data[75][4], rows_data[75][5])

# 3) 2 duplicate rows (duplicate rows 20 and 80, insert after row 100 and 150)
dup1 = rows_data[20]
dup2 = rows_data[80]

for idx, (dt, region, person, product, units, price) in enumerate(rows_data):
    row = idx + 2
    ws1.cell(row=row, column=1, value=dt).number_format = 'YYYY-MM-DD'
    ws1.cell(row=row, column=2, value=region)
    ws1.cell(row=row, column=3, value=person if person else "")
    ws1.cell(row=row, column=4, value=product)
    ws1.cell(row=row, column=5, value=units)
    ws1.cell(row=row, column=6, value=price).number_format = '#,##0.00'
    # Total formula: Units * Price
    ws1.cell(row=row, column=7).value = f"=E{row}*F{row}"
    ws1.cell(row=row, column=7).number_format = '#,##0.00'
    # Commission: 3% of Total
    ws1.cell(row=row, column=8).value = f"=G{row}*0.03"
    ws1.cell(row=row, column=8).number_format = '#,##0.00'
    
    # Apply borders
    for c in range(1, 9):
        ws1.cell(row=row, column=c).border = thin_border

# Add duplicate rows after the 200 data rows
dup_row_1 = 202
for col_idx, val in enumerate([dup1[0], dup1[1], dup1[2], dup1[3], dup1[4], dup1[5]], 1):
    cell = ws1.cell(row=dup_row_1, column=col_idx, value=val)
    cell.border = thin_border
    if col_idx == 1:
        cell.number_format = 'YYYY-MM-DD'
    elif col_idx == 6:
        cell.number_format = '#,##0.00'
ws1.cell(row=dup_row_1, column=7).value = f"=E{dup_row_1}*F{dup_row_1}"
ws1.cell(row=dup_row_1, column=7).number_format = '#,##0.00'
ws1.cell(row=dup_row_1, column=7).border = thin_border
ws1.cell(row=dup_row_1, column=8).value = f"=G{dup_row_1}*0.03"
ws1.cell(row=dup_row_1, column=8).number_format = '#,##0.00'
ws1.cell(row=dup_row_1, column=8).border = thin_border

dup_row_2 = 203
for col_idx, val in enumerate([dup2[0], dup2[1], dup2[2], dup2[3], dup2[4], dup2[5]], 1):
    cell = ws1.cell(row=dup_row_2, column=col_idx, value=val)
    cell.border = thin_border
    if col_idx == 1:
        cell.number_format = 'YYYY-MM-DD'
    elif col_idx == 6:
        cell.number_format = '#,##0.00'
ws1.cell(row=dup_row_2, column=7).value = f"=E{dup_row_2}*F{dup_row_2}"
ws1.cell(row=dup_row_2, column=7).number_format = '#,##0.00'
ws1.cell(row=dup_row_2, column=7).border = thin_border
ws1.cell(row=dup_row_2, column=8).value = f"=G{dup_row_2}*0.03"
ws1.cell(row=dup_row_2, column=8).number_format = '#,##0.00'
ws1.cell(row=dup_row_2, column=8).border = thin_border

# Column widths
ws1.column_dimensions['A'].width = 12
ws1.column_dimensions['B'].width = 10
ws1.column_dimensions['C'].width = 20
ws1.column_dimensions['D'].width = 12
ws1.column_dimensions['E'].width = 12
ws1.column_dimensions['F'].width = 12
ws1.column_dimensions['G'].width = 14
ws1.column_dimensions['H'].width = 14

# Auto-filter
ws1.auto_filter.ref = f"A1:H{dup_row_2}"

# ─── Sheet 2: Employee Data ───
ws2 = wb.create_sheet("Employee-Data")

emp_headers = ["Employee ID", "Name", "Department", "Hire Date", "Salary", "Performance Rating", "Leave Days Taken"]
for col, h in enumerate(emp_headers, 1):
    cell = ws2.cell(row=1, column=col, value=h)
    cell.font = header_font_white
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')
    cell.border = thin_border

departments = ["Sales", "Marketing", "Engineering", "HR", "Finance"]
emp_names = [
    "Amara Obi", "Benjamin Torres", "Chloe Wang", "David Osei", "Elena Volkov",
    "Frank Mensah", "Grace Tanaka", "Hassan Youssef", "Irene Kowalski", "James Okonkwo",
    "Keiko Sato", "Liam O'Brien", "Maya Sharma", "Noah Abrams", "Olivia Dupont",
    "Pedro Gutierrez", "Quinn Foster", "Rosa Martinez", "Samuel Nkrumah", "Tanya Ivanova",
    "Umar Diallo", "Vera Lindberg", "Wendy Chang", "Xavier Moreau", "Yuki Nakamura",
    "Zara Ahmed", "Andre Silva", "Bianca Rossi", "Charles Dubois", "Diana Kozlov",
    "Emmanuel Asante", "Fatou Diallo", "Gabriel Herrera", "Helen Papadopoulos",
    "Isaac Okello", "Julia Bergström", "Kofi Antwi", "Lucia Fernandez", "Michael Eriksen",
    "Nadia Petrov", "Oscar Lindqvist", "Priya Kapoor", "Ricardo Vega", "Sarah Okello",
    "Thomas Fischer", "Ursula Novak", "Victor Adeyemi", "Wanda Kovalenko",
    "Ximena Rojas", "Yusuf Ibrahim"
]

emp_data = []
for i in range(50):
    name = emp_names[i]
    dept = random.choice(departments)
    hire_year = random.randint(2018, 2024)
    hire_month = random.randint(1, 12)
    hire_day = random.randint(1, 28)
    hire_date = datetime.date(hire_year, hire_month, hire_day)
    salary = random.choice(range(30000, 91000, 5000))
    rating = random.randint(1, 5)
    leave = random.randint(0, 25)
    emp_id = f"EMP{str(i+1).zfill(4)}"
    emp_data.append((emp_id, name, dept, hire_date, salary, rating, leave))

for idx, (eid, name, dept, hire_date, salary, rating, leave) in enumerate(emp_data):
    row = idx + 2
    ws2.cell(row=row, column=1, value=eid).border = thin_border
    ws2.cell(row=row, column=2, value=name).border = thin_border
    ws2.cell(row=row, column=3, value=dept).border = thin_border
    ws2.cell(row=row, column=4, value=hire_date).number_format = 'YYYY-MM-DD'
    ws2.cell(row=row, column=4).border = thin_border
    ws2.cell(row=row, column=5, value=salary).number_format = '#,##0'
    ws2.cell(row=row, column=5).border = thin_border
    ws2.cell(row=row, column=6, value=rating).border = thin_border
    ws2.cell(row=row, column=7, value=leave).border = thin_border

# Column widths for Employee-Data
for col_letter, width in [('A', 14), ('B', 22), ('C', 14), ('D', 12), ('E', 12), ('F', 20), ('G', 18)]:
    ws2.column_dimensions[col_letter].width = width

ws2.auto_filter.ref = f"A1:G51"

# Freeze panes for both sheets
ws1.freeze_panes = "A2"
ws2.freeze_panes = "A2"

# Save
output_path = "/home/ennycares/EXCEL Learn/exercises/sales-data-practice.xlsx"
wb.save(output_path)
print(f"✅ Saved to {output_path}")
print(f"   Sheet 'Sales-Data': 200 data rows + 2 duplicates + header = 203 rows")
print(f"   Sheet 'Employee-Data': 50 data rows + header = 51 rows")
print(f"   Intentional issues:")
print(f"     - 3 cells with extra spaces (rows 12, 52, 122)")
print(f"     - 1 blank cell (row 77, Salesperson)")
print(f"     - 2 duplicate rows (rows 202-203, copies of rows 22 and 82)")
