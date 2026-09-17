# Module 7: VBA & Macros — Excel Automation Mastery

> **Module Goal:** Master Visual Basic for Applications (VBA) to automate repetitive Excel tasks, build custom functions, create interactive forms, and develop professional-grade automation solutions.

---

## Table of Contents

1. [What is VBA? Why Automate Excel?](#1-what-is-vba-why-automate-excel)
2. [Developer Tab: Enabling and Using](#2-developer-tab-enabling-and-using)
3. [Recording Macros: Step by Step](#3-recording-macros-step-by-step)
4. [Running Macros](#4-running-macros)
5. [VBA Editor (Alt+F11)](#5-vba-editor-altf11)
6. [VBA Basics](#6-vba-basics)
7. [Working with Worksheets and Workbooks via VBA](#7-working-with-worksheets-and-workbooks-via-vba)
8. [Error Handling](#8-error-handling)
9. [UserForms Basics](#9-userforms-basics)
10. [Common Automation Tasks](#10-common-automation-tasks)
11. [Best Practices for VBA Code](#11-best-practices-for-vba-code)
12. [Practice Exercises](#12-practice-exercises)
13. [Video References](#13-video-references)
14. [Sources & Further Reading](#14-sources--further-reading)

---

## 1. What is VBA? Why Automate Excel?

### What is VBA?

**VBA** stands for **Visual Basic for Applications**. It is a programming language developed by Microsoft and embedded into several Office applications including Excel, Word, Outlook, and Access.

VBA allows you to:

- **Automate repetitive tasks** — Format reports, clean data, generate summaries with one click
- **Create custom functions** — Build functions that don't exist in Excel's built-in library
- **Build interactive applications** — Create UserForms for data entry, dashboards, and mini-apps
- **Control other Office apps** — Send emails via Outlook, create Word documents, interact with Access databases
- **Extend Excel's capabilities** — Perform tasks impossible with formulas alone

### Why Automate Excel?

| Manual Approach | VBA Automation |
|----------------|----------------|
| Format 50 reports manually | One macro formats all 50 in seconds |
| Copy-paste data daily | Auto-import with a single click |
| Build charts one by one | Generate all charts programmatically |
| Send weekly email summaries | VBA sends them automatically via Outlook |
| Clean messy data row by row | Loop through thousands of rows instantly |

### Real-World Impact

- **Finance:** Automated financial models, scenario analysis, report generation
- **HR:** Employee data management, leave tracking, payroll calculations
- **Marketing:** Campaign analysis, data consolidation, automated dashboards
- **Operations:** Inventory tracking, order processing, supply chain monitoring
- **Education:** Grade calculations, attendance tracking, report cards

### VBA vs. Other Approaches

| Feature | Formulas | Power Query | VBA |
|---------|----------|-------------|-----|
| Repetitive formatting | ❌ | ❌ | ✅ |
| Custom dialogs | ❌ | ❌ | ✅ |
| Data transformation | Limited | ✅ | ✅ |
| Conditional logic | Limited | ✅ | ✅ |
| External app control | ❌ | ❌ | ✅ |
| User interaction | ❌ | ❌ | ✅ |
| Learning curve | Low | Medium | Medium-High |

> **Key Insight:** VBA is not a replacement for formulas or Power Query — it complements them. Use formulas for calculations, Power Query for data transformation, and VBA for automation and interaction.

---

## 2. Developer Tab: Enabling and Using

### Why Enable the Developer Tab?

The Developer Tab provides access to:
- **Macro recording and running**
- **VBA Editor** (Visual Basic button)
- **Form Controls** (buttons, checkboxes, dropdowns)
- **ActiveX Controls** (advanced interactive controls)
- **XML and Add-in management**

### How to Enable the Developer Tab

#### Excel 365 / 2021 / 2019 / 2016:

1. Go to **File** → **Options**
2. Click **Customize Ribbon** in the left panel
3. In the right panel, check the box next to **Developer**
4. Click **OK**

#### Excel for Mac:

1. Go to **Excel** → **Preferences** → **Ribbon & Toolbar**
2. In the right panel, check **Developer**
3. Click **Save**

### Developer Tab Components

```
Developer Tab
├── Code Group
│   ├── Visual Basic (opens VBA Editor)
│   ├── Macros (view/run macros)
│   ├── Record Macro
│   ├── Use Relative References
│   └── Macro Security
├── Add-ins Group
│   ├── Excel Add-ins
│   └── COM Add-ins
├── Controls Group
│   ├── Insert (Form & ActiveX controls)
│   ├── Design Mode
│   ├── Properties
│   └── View Code
├── XML Group
│   ├── Source
│   ├── Map
│   └── Expansion Packs
└── Modify Group
    └── Document Panel
```

### Macro Security Settings

1. Go to **Developer** → **Macro Security**
2. Choose your security level:

| Setting | Description | Recommended For |
|---------|-------------|-----------------|
| Disable all macros without notification | Blocks all macros | High-security environments |
| Disable all macros with notification | Shows warning bar | **Most users (recommended)** |
| Disable all macros except digitally signed | Only trusted macros | Corporate environments |
| Enable all macros | Runs everything (not recommended) | Development only |

> **Best Practice:** Use "Disable all macros with notification" — you'll see a yellow security bar when opening files with macros, and can choose to enable them.

---

## 3. Recording Macros: Step by Step

### What is Macro Recording?

Macro recording captures your actions in Excel and converts them into VBA code. It's the easiest way to start learning VBA — record actions, then examine the generated code.

### Step-by-Step: Recording Your First Macro

#### Example: Auto-Format a Sales Report

**Scenario:** You receive a weekly sales report and always format it the same way.

1. **Start Recording:**
   - Go to **Developer** → **Record Macro**
   - Or click the small circle in the bottom-left status bar

2. **Set Macro Properties:**
   - **Macro name:** `FormatSalesReport`
   - **Shortcut key:** `Ctrl+Shift+F`
   - **Store in:** This Workbook
   - **Description:** Formats sales report with standard formatting

3. **Perform Your Actions:**
   ```
   1. Select A1 (title cell)
   2. Bold the text, set font size to 14
   3. Select A1:F1, apply Merge & Center
   4. Select the header row (A2:F2)
   5. Bold, apply background color (blue), white font
   6. Select data range A2:F100
   7. Apply borders (all borders)
   8. Auto-fit columns
   9. Select column D (revenue), apply Currency format
   10. Freeze panes at row 3
   ```

4. **Stop Recording:**
   - Go to **Developer** → **Stop Recording**
   - Or click the square in the status bar

### Examining the Recorded Code

Press **Alt+F11** to open the VBA Editor. The recorded macro looks like this:

```vba
Sub FormatSalesReport()
    ' Formats sales report with standard formatting
    ' Recorded on 2026-09-10
    
    ' Format title
    Range("A1").Select
    Selection.Font.Bold = True
    Selection.Font.Size = 14
    Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .MergeCells = True
    End With
    
    ' Format headers
    Range("A2:F2").Select
    With Selection
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Apply borders
    Range("A2:F100").Select
    Selection.Borders.LineStyle = xlContinuous
    
    ' Auto-fit columns
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    ' Format revenue column as currency
    Range("D2:D100").Select
    Selection.NumberFormat = "$#,##0.00"
    
    ' Freeze panes
    Range("A3").Select
    ActiveWindow.FreezePanes = True
End Sub
```

### Improving Recorded Macros

Recorded macros work but are often inefficient. Here's the improved version:

```vba
Sub FormatSalesReport_Improved()
    ' Improved version - no Select statements
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Format title
    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With
    With ws.Range("A1:F1")
        .HorizontalAlignment = xlCenter
        .MergeCells = True
    End With
    
    ' Format headers
    With ws.Range("A2:F2")
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Apply borders
    With ws.Range("A2:F100").Borders
        .LineStyle = xlContinuous
    End With
    
    ' Auto-fit columns
    ws.Cells.EntireColumn.AutoFit
    
    ' Format revenue column as currency
    ws.Range("D2:D100").NumberFormat = "$#,##0.00"
    
    ' Freeze panes
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True
    
    MsgBox "Report formatted successfully!", vbInformation
End Sub
```

### Macro Recording Best Practices

1. **Plan your actions** before recording
2. **Use relative references** when needed (Developer → Use Relative References)
3. **Keep recordings short** — break complex tasks into multiple macros
4. **Always examine and clean** the generated code
5. **Remove unnecessary `.Select` statements** — they slow down execution
6. **Test immediately** after recording
7. **Add comments** to explain what the macro does

---

## 4. Running Macros

### Method 1: Macro Dialog Box

1. Press **Alt+F8**
2. Select the macro from the list
3. Click **Run**

### Method 2: Keyboard Shortcut

- Assign during recording or via **Alt+F8** → **Options**
- Default: `Ctrl+letter` (e.g., `Ctrl+Shift+F`)

### Method 3: Button on Worksheet

1. Go to **Developer** → **Insert** → **Button (Form Control)**
2. Draw the button on your worksheet
3. In the Assign Macro dialog, select your macro
4. Right-click the button to edit text

```vba
' You can also create buttons programmatically:
Sub CreateMacroButton()
    Dim btn As Button
    Set btn = ActiveSheet.Buttons.Add(10, 10, 120, 30)
    With btn
        .Caption = "Format Report"
        .OnAction = "FormatSalesReport"
        .Font.Size = 11
    End With
End Sub
```

### Method 4: Quick Access Toolbar

1. Click the dropdown arrow on the Quick Access Toolbar
2. Select **More Commands**
3. Choose **Macros** from the "Choose commands from" dropdown
4. Select your macro and click **Add**
5. Click **Modify** to change the icon

### Method 5: Ribbon Button

1. Go to **File** → **Options** → **Customize Ribbon**
2. Create a new tab or group
3. Select **Macros** from the dropdown
4. Add your macro to the custom group

### Method 6: From VBA Editor

- Press **Alt+F11** to open VBA Editor
- Place cursor inside the macro
- Press **F5** or click **Run**

### Method 7: Auto-Run Macros

```vba
' Runs automatically when the workbook opens
Private Sub Workbook_Open()
    MsgBox "Welcome! Workbook is ready."
End Sub

' Runs automatically when a cell changes
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        MsgBox "Cell A1 changed to: " & Target.Value
    End If
End Sub
```

---

## 5. VBA Editor (Alt+F11)

### Opening the VBA Editor

Press **Alt+F11** from Excel to open the Visual Basic Editor.

### VBA Editor Layout

```
┌─────────────────────────────────────────────────────────────┐
│  VBA Editor - Microsoft Visual Basic for Applications       │
├─────────────────────────────────────────────────────────────┤
│ File  Edit  View  Insert  Format  Debug  Run  Tools  Add-Ins│
├──────────┬──────────────────────────────────────────────────┤
│ Project  │  Code Window                                     │
│ Explorer │  ┌──────────────────────────────────────────────┐│
│          │  │  (Your VBA code goes here)                   ││
│ VBAProj  │  │                                              ││
│  │       │  │                                              ││
│  ├─Sheet1 │  │                                              ││
│  ├─Sheet2 │  │                                              ││
│  ├─ThisWb │  │                                              ││
│  └─Module1│  │                                              ││
│          │  └──────────────────────────────────────────────┘│
│──────────┤──────────────────────────────────────────────────┤
│Properties│  Immediate Window (Ctrl+G)                       │
│          │  ┌──────────────────────────────────────────────┐│
│ Name:    │  │  ? Range("A1").Value                         ││
│ Module1  │  │  Hello World                                 ││
│          │  └──────────────────────────────────────────────┘│
└──────────┴──────────────────────────────────────────────────┘
```

### Project Explorer (Ctrl+R)

Shows the hierarchy of all open workbooks and their components:

```
VBAProject (Book1.xlsm)
├── Microsoft Excel Objects
│   ├── Sheet1 (Sheet1)
│   ├── Sheet2 (Sheet2)
│   └── ThisWorkbook
├── Forms
│   └── UserForm1
└── Modules
    ├── Module1
    └── Module2
```

**Where to place code:**

| Location | Use For |
|----------|---------|
| **Module** | General-purpose macros and functions |
| **Sheet object** | Sheet-specific event code (e.g., Worksheet_Change) |
| **ThisWorkbook** | Workbook-level events (e.g., Workbook_Open) |
| **UserForm** | Form control event code |
| **Class Module** | Custom objects and events |

### Properties Window (F4)

Displays properties of the selected object. You can change:
- **Name** — The code name of the object
- **Visible** — Show/hide sheets (xlSheetVisible, xlSheetHidden, xlSheetVeryHidden)
- **Tab Color** — Sheet tab color

### Code Window

Where you write and edit VBA code. Features:
- **Syntax highlighting** — Keywords, comments, strings in different colors
- **Auto-complete** — Type `Range("` and see suggestions
- **Procedure dropdown** — Jump to any Sub/Function in the module
- **Split bar** — Divide the window to view two code sections

### Immediate Window (Ctrl+G)

A powerful debugging tool:

```vba
' Test expressions immediately (prefix with ?)
? Range("A1").Value          ' Returns: Hello World
? ActiveSheet.Name           ' Returns: Sheet1
? Cells(1, 1).Address        ' Returns: $A$1

' Execute statements
Range("A1").Value = "Updated"
Sheets("Sheet1").Activate

' Debug output
Debug.Print "Variable x = " & x
```

### Essential VBA Editor Shortcuts

| Shortcut | Action |
|----------|--------|
| Alt+F11 | Toggle between Excel and VBA Editor |
| F5 | Run the current procedure |
| F8 | Step through code (one line at a time) |
| F9 | Toggle breakpoint |
| Ctrl+G | Show Immediate Window |
| Ctrl+R | Show Project Explorer |
| F4 | Show Properties Window |
| Ctrl+Space | Auto-complete |
| Ctrl+J | List members |
| Ctrl+Shift+I | Quick Info |
| Ctrl+Z | Undo |
| Ctrl+F | Find |
| Ctrl+H | Find and Replace |

---

## 6. VBA Basics

### 6.1 Sub and Function Procedures

#### Sub Procedures

A **Sub** performs an action but doesn't return a value:

```vba
Sub SayHello()
    MsgBox "Hello, World!"
End Sub

' Sub with parameters
Sub GreetUser(userName As String)
    MsgBox "Hello, " & userName & "!"
End Sub

' Call it:
' GreetUser "Alice"
```

#### Function Procedures

A **Function** performs a calculation and returns a value:

```vba
' Custom function usable in worksheets
Function CalculateTax(price As Double, taxRate As Double) As Double
    CalculateTax = price * taxRate
End Function

' Usage in a cell: =CalculateTax(100, 0.08) returns 8

' Function with optional parameter
Function FullName(firstName As String, lastName As String, _
                  Optional middleName As String = "") As String
    If middleName = "" Then
        FullName = firstName & " " & lastName
    Else
        FullName = firstName & " " & middleName & " " & lastName
    End If
End Function
```

#### Difference Between Sub and Function

| Feature | Sub | Function |
|---------|-----|----------|
| Returns value | No | Yes |
| Called from worksheet | No | Yes (=FunctionName()) |
| Called from VBA | `Call SubName` or `SubName` | `result = FunctionName()` |
| Can modify cells | Yes | Yes (but not recommended) |
| Entry point | Can be assigned to buttons | Used in formulas |

### 6.2 Variables and Data Types

#### Declaring Variables

```vba
' Explicit declaration (recommended)
Dim employeeName As String
Dim age As Integer
Dim salary As Double
Dim isActive As Boolean
Dim startDate As Date

' Multiple declarations on one line
Dim x As Integer, y As Integer, z As Integer

' Module-level variable (available to all procedures in the module)
Dim totalCount As Long

' Public variable (available to all modules)
Public companyName As String
```

#### VBA Data Types

| Data Type | Storage | Range | Example |
|-----------|---------|-------|---------|
| **Byte** | 1 byte | 0 to 255 | `Dim b As Byte` |
| **Boolean** | 2 bytes | True/False | `Dim flag As Boolean` |
| **Integer** | 2 bytes | -32,768 to 32,767 | `Dim i As Integer` |
| **Long** | 4 bytes | -2.1B to 2.1B | `Dim l As Long` |
| **LongLong** | 8 bytes | Very large numbers | `Dim ll As LongLong` |
| **Single** | 4 bytes | Decimal (7 digits) | `Dim s As Single` |
| **Double** | 8 bytes | Decimal (15 digits) | `Dim d As Double` |
| **Currency** | 8 bytes | Financial precision | `Dim c As Currency` |
| **Date** | 8 bytes | Dates and times | `Dim dt As Date` |
| **String** | Varies | Text (up to 2GB) | `Dim str As String` |
| **Variant** | Varies | Any type (default) | `Dim v As Variant` |
| **Object** | 4 bytes | Object reference | `Dim ws As Worksheet` |

#### Choosing the Right Data Type

```vba
' Use Long instead of Integer (faster on modern systems)
Dim rowNumber As Long  ' NOT Integer

' Use Double for calculations
Dim revenue As Double  ' NOT Single (precision matters)

' Use String for text
Dim customerName As String

' Use Date for dates
Dim orderDate As Date

' Avoid Variant when possible (uses more memory)
' BAD:  Dim x As Variant  (or just Dim x)
' GOOD: Dim x As Long
```

### 6.3 Option Explicit

**Always use Option Explicit!** It forces you to declare all variables, preventing typos from creating new variables.

```vba
Option Explicit  ' Must be at the TOP of every module

Sub Example()
    Dim counter As Integer
    conuter = 5  ' TYPO! Without Option Explicit, this creates a new variable
    ' With Option Explicit: ERROR! "Variable not defined"
    Debug.Print counter  ' Returns 0 (not 5!) without Option Explicit
End Sub
```

**To set Option Explicit automatically:**
1. In VBA Editor, go to **Tools** → **Options**
2. Check **Require Variable Declaration**
3. Click **OK**

### 6.4 MsgBox and InputBox

#### MsgBox

```vba
' Simple message
MsgBox "Operation completed!"

' With title
MsgBox "Data saved successfully.", vbInformation, "Success"

' With buttons and capturing response
Dim response As VbMsgBoxResult
response = MsgBox("Do you want to continue?", vbYesNoCancel + vbQuestion, "Confirm")

Select Case response
    Case vbYes
        ' User clicked Yes
        MsgBox "You chose Yes"
    Case vbNo
        ' User clicked No
        MsgBox "You chose No"
    Case vbCancel
        ' User clicked Cancel
        MsgBox "Operation cancelled"
End Select

' MsgBox buttons and icons:
' vbOKOnly (0)           vbOKCancel (1)
' vbAbortRetryIgnore (2) vbYesNoCancel (3)
' vbYesNo (4)            vbRetryCancel (5)
'
' vbCritical (16)        vbQuestion (32)
' vbExclamation (48)     vbInformation (64)
```

#### InputBox

```vba
' Simple input
Dim userName As String
userName = InputBox("Enter your name:", "User Input")

' With default value
Dim city As String
city = InputBox("Enter your city:", "Location", "New York")

' Validate input
Dim age As String
age = InputBox("Enter your age:", "Age")
If age = "" Then
    MsgBox "No input provided."
ElseIf Not IsNumeric(age) Then
    MsgBox "Please enter a valid number."
Else
    MsgBox "You are " & age & " years old."
End If

' Application.InputBox (more powerful, allows range selection)
Dim rng As Range
On Error Resume Next
Set rng = Application.InputBox("Select a range:", Type:=8)  ' Type 8 = Range
On Error GoTo 0
If rng Is Nothing Then
    MsgBox "No range selected."
Else
    MsgBox "You selected: " & rng.Address
End If
```

### 6.5 Range and Cells Object Model

#### The Range Object

```vba
' Referencing ranges
Range("A1")                    ' Single cell
Range("A1:D10")                ' Range of cells
Range("A1, C1, E1")           ' Non-contiguous range
Range("A1:" & "D" & lastRow)  ' Dynamic range

' Using Cells (row, column) - great for loops
Cells(1, 1)        ' A1
Cells(5, 3)        ' C5
Cells(Rows.Count, 1)  ' Last cell in column A

' Combining Range and Cells
Range(Cells(1, 1), Cells(10, 5))  ' A1:E10

' Named ranges
Range("SalesData")

' Entire rows and columns
Range("A:A")           ' Entire column A
Range("1:1")           ' Entire row 1
Columns("A:D")         ' Columns A through D
Rows("1:5")            ' Rows 1 through 5
```

#### Working with Range Properties

```vba
' Value
Range("A1").Value = "Hello"
Dim val As String
val = Range("A1").Value

' Address
Debug.Print Range("A1").Address          ' $A$1
Debug.Print Range("A1").Address(False, False)  ' A1 (relative)

' Offset
Range("A1").Offset(1, 0)   ' A2 (down 1, right 0)
Range("A1").Offset(0, 2)   ' C1 (down 0, right 2)

' Resize
Range("A1").Resize(3, 4)   ' A1:D3

' End (like Ctrl+Arrow)
Range("A1").End(xlDown)     ' Last filled cell down from A1
Range("A1").End(xlToRight)  ' Last filled cell right from A1

' Find last row and column
Dim lastRow As Long
Dim lastCol As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
lastCol = Cells(1, Columns.Count).End(xlToLeft).Column

' CurrentRegion (like Ctrl+Shift+*)
Range("A1").CurrentRegion.Select

' SpecialCells
' Find all cells with formulas
ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas).Select
' Find all blank cells
ActiveSheet.Cells.SpecialCells(xlCellTypeBlanks).Select
```

#### Range Actions

```vba
' Clear contents and formatting
Range("A1:D10").Clear           ' Clear everything
Range("A1:D10").ClearContents   ' Clear values only
Range("A1:D10").ClearFormats    ' Clear formatting only

' Copy and Paste
Range("A1:A10").Copy Destination:=Range("C1")
Range("A1:A10").Copy
Range("C1").PasteSpecial xlPasteValues

' Sort
Range("A1:D100").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

' Filter
Range("A1:D100").AutoFilter Field:=1, Criteria1:="Sales"

' Find
Dim foundCell As Range
Set foundCell = Range("A:A").Find(What:="Target", LookIn:=xlValues)
If Not foundCell Is Nothing Then
    MsgBox "Found at: " & foundCell.Address
End If
```

### 6.6 With...End With

```vba
' Without With (repetitive)
Range("A1").Font.Bold = True
Range("A1").Font.Size = 14
Range("A1").Font.Color = RGB(255, 0, 0)
Range("A1").Font.Name = "Arial"

' With With (cleaner)
With Range("A1").Font
    .Bold = True
    .Size = 14
    .Color = RGB(255, 0, 0)
    .Name = "Arial"
End With

' Nested With
With Range("A1:D1")
    .Font.Bold = True
    .Interior.Color = RGB(0, 51, 102)
    .Font.Color = RGB(255, 255, 255)
    With .Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End With
```

### 6.7 Conditional Statements

#### If...Then...Else

```vba
' Simple If
If Range("A1").Value > 100 Then
    MsgBox "Greater than 100"
End If

' If...Else
If Range("A1").Value > 100 Then
    Range("B1").Value = "High"
Else
    Range("B1").Value = "Low"
End If

' If...ElseIf...Else
Dim score As Double
score = Range("A1").Value

If score >= 90 Then
    Range("B1").Value = "A"
ElseIf score >= 80 Then
    Range("B1").Value = "B"
ElseIf score >= 70 Then
    Range("B1").Value = "C"
ElseIf score >= 60 Then
    Range("B1").Value = "D"
Else
    Range("B1").Value = "F"
End If

' Single-line If (for simple cases)
If Range("A1").Value = "" Then Range("A1").Value = "N/A"

' If with logical operators
If Range("A1").Value > 0 And Range("B1").Value > 0 Then
    MsgBox "Both positive"
End If

If Range("A1").Value = "Yes" Or Range("A1").Value = "Y" Then
    MsgBox "Confirmed"
End If
```

#### Select Case

```vba
' Select Case is cleaner for multiple conditions
Dim dept As String
dept = Range("A1").Value

Select Case dept
    Case "Sales"
        Range("B1").Value = "Revenue Generation"
    Case "Marketing"
        Range("B1").Value = "Brand Awareness"
    Case "IT"
        Range("B1").Value = "Technology Support"
    Case "HR"
        Range("B1").Value = "People Management"
    Case "Finance"
        Range("B1").Value = "Financial Operations"
    Case Else
        Range("B1").Value = "Unknown Department"
End Select

' Select Case with ranges
Select Case score
    Case Is >= 90
        grade = "A"
    Case 80 To 89
        grade = "B"
    Case 70 To 79
        grade = "C"
    Case 60 To 69
        grade = "D"
    Case Else
        grade = "F"
End Select

' Select Case with multiple values
Select Case weekday
    Case 1, 7
        dayType = "Weekend"
    Case 2 To 6
        dayType = "Weekday"
End Select
```

### 6.8 Loops

#### For...Next

```vba
' Basic For loop
Dim i As Long
For i = 1 To 10
    Cells(i, 1).Value = i * 10
Next i

' With Step
For i = 1 To 100 Step 5
    Cells(i, 1).Value = "Every 5th row"
Next i

' Reverse loop
For i = 100 To 1 Step -1
    If Cells(i, 1).Value = "" Then
        Cells(i, 1).EntireRow.Delete
    End If
Next i

' Loop through rows with data
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow  ' Start at 2 to skip header
    If Cells(i, 3).Value > 1000 Then
        Cells(i, 3).Interior.Color = RGB(0, 255, 0)  ' Highlight
    End If
Next i
```

#### For Each...Next

```vba
' Loop through cells in a range
Dim cell As Range
For Each cell In Range("A1:A100")
    If cell.Value < 0 Then
        cell.Font.Color = RGB(255, 0, 0)  ' Red for negatives
    End If
Next cell

' Loop through all worksheets
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    Debug.Print ws.Name
    ws.Range("A1").Value = "Updated"
Next ws

' Loop through all open workbooks
Dim wb As Workbook
For Each wb In Application.Workbooks
    Debug.Print wb.Name
Next wb

' Loop through selected cells
For Each cell In Selection
    cell.Value = UCase(cell.Value)  ' Convert to uppercase
Next cell
```

#### Do While / Do Until

```vba
' Do While - runs while condition is True
Dim row As Long
row = 1
Do While Cells(row, 1).Value <> ""
    Cells(row, 2).Value = Cells(row, 1).Value * 2
    row = row + 1
Loop

' Do Until - runs until condition becomes True
row = 1
Do Until IsEmpty(Cells(row, 1))
    Cells(row, 2).Value = Cells(row, 1).Value * 2
    row = row + 1
Loop

' Do...Loop While (runs at least once)
Do
    Cells(row, 2).Value = Cells(row, 1).Value * 2
    row = row + 1
Loop While Cells(row, 1).Value <> ""

' Exit Do - break out of loop
Do While True
    If Cells(row, 1).Value = "STOP" Then Exit Do
    row = row + 1
Loop
```

### 6.9 Arrays

#### Static Arrays

```vba
' Declare a static array
Dim names(1 To 5) As String
names(1) = "Alice"
names(2) = "Bob"
names(3) = "Charlie"
names(4) = "Diana"
names(5) = "Eve"

' Multi-dimensional array
Dim matrix(1 To 3, 1 To 3) As Double
matrix(1, 1) = 1: matrix(1, 2) = 2: matrix(1, 3) = 3
matrix(2, 1) = 4: matrix(2, 2) = 5: matrix(2, 3) = 6
matrix(3, 1) = 7: matrix(3, 2) = 8: matrix(3, 3) = 9

' Initialize with Array function
Dim colors As Variant
colors = Array("Red", "Green", "Blue", "Yellow")
```

#### Dynamic Arrays

```vba
' Dynamic array - size determined at runtime
Dim data() As Variant
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Read range into array (FAST!)
data = Range("A1:D" & lastRow).Value

' Process array (much faster than cell-by-cell)
Dim i As Long
For i = 1 To UBound(data, 1)
    data(i, 4) = data(i, 2) * data(i, 3)  ' Calculate total
Next i

' Write array back to range (FAST!)
Range("A1:D" & lastRow).Value = data
```

#### Array Performance Tip

```vba
' SLOW - reading/writing cells one by one
Sub SlowMethod()
    Dim i As Long
    For i = 1 To 10000
        Cells(i, 2).Value = Cells(i, 1).Value * 2
    Next i
End Sub

' FAST - using arrays
Sub FastMethod()
    Dim data() As Variant
    Dim result() As Variant
    Dim i As Long
    
    ' Read all data at once
    data = Range("A1:A10000").Value
    
    ' Prepare result array
    ReDim result(1 To UBound(data, 1), 1 To 1)
    
    ' Process in memory
    For i = 1 To UBound(data, 1)
        result(i, 1) = data(i, 1) * 2
    Next i
    
    ' Write all results at once
    Range("B1:B10000").Value = result
End Sub
```

---

## 7. Working with Worksheets and Workbooks via VBA

### Worksheet Operations

```vba
' Reference worksheets
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sales")          ' By name
Set ws = ThisWorkbook.Sheets(1)                ' By index
Set ws = ActiveSheet                            ' Currently active

' Create new worksheet
Dim newWs As Worksheet
Set newWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
newWs.Name = "NewSheet"

' Delete worksheet (with confirmation)
Application.DisplayAlerts = False
ThisWorkbook.Sheets("TempSheet").Delete
Application.DisplayAlerts = True

' Copy worksheet
ThisWorkbook.Sheets("Template").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

' Move worksheet
ThisWorkbook.Sheets("Sheet1").Move Before:=ThisWorkbook.Sheets(1)

' Hide/Show worksheets
ws.Visible = xlSheetHidden       ' Hidden (user can unhide)
ws.Visible = xlSheetVeryHidden   ' Very Hidden (only via VBA)
ws.Visible = xlSheetVisible      ' Visible

' Rename worksheet
ws.Name = "Q1 Sales"

' Protect/Unprotect
ws.Protect Password:="myPassword", AllowFiltering:=True
ws.Unprotect Password:="myPassword"

' Loop through all sheets
Dim sheet As Worksheet
For Each sheet In ThisWorkbook.Worksheets
    Debug.Print sheet.Name & " - " & sheet.Cells(Rows.Count, 1).End(xlUp).Row & " rows"
Next sheet
```

### Workbook Operations

```vba
' Create new workbook
Dim newWb As Workbook
Set newWb = Workbooks.Add

' Open workbook
Dim existingWb As Workbook
Set existingWb = Workbooks.Open("C:\Data\Sales.xlsx")

' Save workbook
ThisWorkbook.Save
existingWb.SaveAs "C:\Data\Sales_Backup.xlsx", xlOpenXMLWorkbook

' Close workbook
existingWb.Close SaveChanges:=True

' ThisWorkbook vs ActiveWorkbook
' ThisWorkbook = workbook containing the VBA code
' ActiveWorkbook = currently active workbook (may be different!)

' Get workbook path
Debug.Print ThisWorkbook.Path        ' C:\Users\Name\Documents
Debug.Print ThisWorkbook.FullName    ' C:\Users\Name\Documents\Book1.xlsm
Debug.Print ThisWorkbook.Name        ' Book1.xlsm

' Copy data between workbooks
SourceWorkbook.Sheets("Data").Range("A1:D100").Copy _
    Destination:=ThisWorkbook.Sheets("Import").Range("A1")
```

---

## 8. Error Handling

### Why Error Handling Matters

Without error handling, VBA crashes with a confusing dialog. With proper error handling, you can:
- Display meaningful error messages
- Log errors for debugging
- Recover gracefully from unexpected situations
- Prevent data corruption

### On Error GoTo

```vba
Sub ProcessData()
    On Error GoTo ErrorHandler
    
    ' Code that might fail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("NonExistentSheet")  ' This will error!
    
    ' More processing...
    ws.Range("A1").Value = "Done"
    
    Exit Sub  ' Important: exit before error handler
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "In procedure: ProcessData", vbCritical, "Error"
    
    ' Optional: log error to a file
    Open ThisWorkbook.Path & "\error.log" For Append As #1
    Print #1, Now & " - Error " & Err.Number & ": " & Err.Description
    Close #1
End Sub
```

### On Error Resume Next

```vba
Sub SafeDeleteSheet()
    On Error Resume Next  ' Ignore errors temporarily
    
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("TempData").Delete
    Application.DisplayAlerts = True
    
    If Err.Number <> 0 Then
        Debug.Print "Sheet didn't exist or couldn't be deleted"
        Err.Clear  ' Reset error
    End If
    
    On Error GoTo 0  ' Resume normal error handling
End Sub

' Check specific error numbers
Sub ReadFile()
    On Error Resume Next
    
    Dim filePath As String
    filePath = "C:\NonExistent\file.xlsx"
    
    Workbooks.Open filePath
    
    Select Case Err.Number
        Case 0
            MsgBox "File opened successfully"
        Case 1004
            MsgBox "File not found: " & filePath
        Case Else
            MsgBox "Unexpected error: " & Err.Description
    End Select
    
    On Error GoTo 0
End Sub
```

### Defensive Programming

```vba
Sub ProcessWorksheet(sheetName As String)
    ' Check if sheet exists before processing
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbExclamation
        Exit Sub
    End If
    
    ' Check if range has data
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        MsgBox "Sheet '" & sheetName & "' is empty!", vbExclamation
        Exit Sub
    End If
    
    ' Safe to proceed
    ' ... process data ...
End Sub
```

### Error Handling Best Practices

1. **Always use error handling** in production code
2. **Exit Sub before ErrorHandler** to prevent fall-through
3. **Use specific error handling** (Resume Next) sparingly
4. **Clear errors** after handling with `Err.Clear`
5. **Log errors** for debugging
6. **Don't ignore errors** — at minimum, log them
7. **Test error paths** — don't just test the happy path

---

## 9. UserForms Basics

### What is a UserForm?

A UserForm is a custom dialog box that provides a user-friendly interface for data input, selection, and interaction.

### Creating a UserForm

1. In VBA Editor, go to **Insert** → **UserForm**
2. Add controls from the **Toolbox**
3. Set properties in the **Properties Window**
4. Write code for control events

### Common Controls

| Control | Purpose | Key Properties |
|---------|---------|---------------|
| **Label** | Display text | Caption, Font |
| **TextBox** | Text input | Value, MaxLength |
| **ComboBox** | Drop-down list | List, Value, Style |
| **ListBox** | List selection | List, MultiSelect |
| **CommandButton** | Click action | Caption, Default |
| **CheckBox** | Toggle option | Value (True/False) |
| **OptionButton** | Radio selection | Value, GroupName |
| **Frame** | Group controls | Caption |
| **SpinButton** | Numeric up/down | Value, Min, Max |
| **Image** | Display images | Picture |

### Example: Employee Data Entry Form

```vba
' UserForm Code (frmEmployee)
' Controls: txtName, txtEmail, cboDepartment, txtSalary, btnSubmit, btnCancel

' Initialize form - populate department dropdown
Private Sub UserForm_Initialize()
    With cboDepartment
        .AddItem "Sales"
        .AddItem "Marketing"
        .AddItem "IT"
        .AddItem "Finance"
        .AddItem "HR"
        .AddItem "Operations"
    End With
End Sub

' Submit button click
Private Sub btnSubmit_Click()
    ' Validate inputs
    If Trim(txtName.Value) = "" Then
        MsgBox "Please enter employee name.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtEmail.Value) = "" Then
        MsgBox "Please enter email address.", vbExclamation
        txtEmail.SetFocus
        Exit Sub
    End If
    
    If cboDepartment.ListIndex = -1 Then
        MsgBox "Please select a department.", vbExclamation
        cboDepartment.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtSalary.Value) Or Val(txtSalary.Value) <= 0 Then
        MsgBox "Please enter a valid salary.", vbExclamation
        txtSalary.SetFocus
        Exit Sub
    End If
    
    ' Find next empty row
    Dim nextRow As Long
    nextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    ' Write data
    With Cells
        Cells(nextRow, 1).Value = txtName.Value
        Cells(nextRow, 2).Value = txtEmail.Value
        Cells(nextRow, 3).Value = cboDepartment.Value
        Cells(nextRow, 4).Value = CDbl(txtSalary.Value)
        Cells(nextRow, 5).Value = Date
    End With
    
    MsgBox "Employee added successfully!", vbInformation
    
    ' Clear form for next entry
    Call ClearForm
End Sub

' Cancel button
Private Sub btnCancel_Click()
    Unload Me
End Sub

' Helper to clear form
Private Sub ClearForm()
    txtName.Value = ""
    txtEmail.Value = ""
    cboDepartment.ListIndex = -1
    txtSalary.Value = ""
    txtName.SetFocus
End Sub
```

### Showing the Form

```vba
' Show form modally (blocks Excel interaction)
Sub ShowEmployeeForm()
    frmEmployee.Show vbModal
End Sub

' Show form modeless (allows Excel interaction)
Sub ShowEmployeeFormModeless()
    frmEmployee.Show vbModeless
End Sub
```

---

## 10. Common Automation Tasks

### 10.1 Auto-Format Reports

```vba
Sub AutoFormatReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    With ws
        ' Title formatting
        With .Range(.Cells(1, 1), .Cells(1, lastCol))
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = RGB(0, 51, 102)
            .HorizontalAlignment = xlCenter
            .MergeCells = True
        End With
        
        ' Header formatting
        With .Range(.Cells(2, 1), .Cells(2, lastCol))
            .Font.Bold = True
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Data formatting
        With .Range(.Cells(2, 1), .Cells(lastRow, lastCol))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 20
        End With
        
        ' Alternating row colors
        Dim i As Long
        For i = 3 To lastRow
            If i Mod 2 = 1 Then
                .Range(.Cells(i, 1), .Cells(i, lastCol)).Interior.Color = RGB(240, 240, 240)
            End If
        Next i
        
        ' Auto-fit columns
        .Cells.EntireColumn.AutoFit
    End With
    
    MsgBox "Report formatted!", vbInformation
End Sub
```

### 10.2 Loop Through Files in a Folder

```vba
Sub ProcessAllFilesInFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim summaryWs As Worksheet
    Dim summaryRow As Long
    
    ' Select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    ' Setup summary sheet
    Set summaryWs = ThisWorkbook.Sheets("Summary")
    summaryRow = 2
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Loop through all Excel files
    fileName = Dir(folderPath & "*.xlsx")
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        
        ' Process each workbook
        With wb.Sheets(1)
            summaryWs.Cells(summaryRow, 1).Value = fileName
            summaryWs.Cells(summaryRow, 2).Value = .Range("B5").Value  ' e.g., total sales
            summaryWs.Cells(summaryRow, 3).Value = .Range("B10").Value ' e.g., profit
        End With
        
        wb.Close SaveChanges:=False
        summaryRow = summaryRow + 1
        fileName = Dir()  ' Get next file
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Processed " & (summaryRow - 2) & " files.", vbInformation
End Sub
```

### 10.3 Send Emails from Excel (Outlook)

```vba
Sub SendBulkEmails()
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("EmailList")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Create Outlook application
    Set outlookApp = CreateObject("Outlook.Application")
    
    For i = 2 To lastRow  ' Skip header
        If ws.Cells(i, 3).Value = "Pending" Then  ' Column C = Status
            Set outlookMail = outlookApp.CreateItem(0)  ' 0 = olMailItem
            
            With outlookMail
                .To = ws.Cells(i, 1).Value           ' Column A = Email
                .Subject = ws.Cells(i, 2).Value      ' Column B = Subject
                .HTMLBody = "<h3>Hello " & ws.Cells(i, 4).Value & "</h3>" & _
                           "<p>Please find attached the Q3 report.</p>" & _
                           "<p>Best regards,<br>Your Team</p>"
                .Attachments.Add ThisWorkbook.Path & "\Reports\Q3_Report.pdf"
                .Send  ' Or .Display to review before sending
            End With
            
            ' Update status
            ws.Cells(i, 3).Value = "Sent"
            ws.Cells(i, 4).Value = Now
            
            Set outlookMail = Nothing
        End If
    Next i
    
    Set outlookApp = Nothing
    
    MsgBox "Emails sent successfully!", vbInformation
End Sub
```

### 10.4 Create Charts Programmatically

```vba
Sub CreateSalesChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As chart
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Delete existing charts
    Dim c As ChartObject
    For Each c In ws.ChartObjects
        c.Delete
    Next c
    
    ' Create new chart
    Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=10, Width:=500, Height:=300)
    Set chart = chartObj.chart
    
    With chart
        .ChartType = xlColumnClustered
        
        ' Set data source
        .SetSourceData Source:=ws.Range("A1:B" & lastRow)
        
        ' Format chart
        .HasTitle = True
        .ChartTitle.Text = "Monthly Sales Report"
        .ChartTitle.Font.Size = 14
        
        ' Format axes
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Revenue ($)"
        
        ' Add data labels
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.NumberFormat = "$#,##0"
        
        ' Apply style
        .ChartStyle = 42
        
        ' Legend
        .HasLegend = False
    End With
    
    MsgBox "Chart created!", vbInformation
End Sub
```

### 10.5 Import/Export Data

```vba
' Import CSV file
Sub ImportCSV()
    Dim filePath As String
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV File")
    
    If filePath = "False" Then Exit Sub
    
    ' Method 1: Open as workbook
    Dim csvWb As Workbook
    Set csvWb = Workbooks.Open(filePath)
    csvWb.Sheets(1).UsedRange.Copy ThisWorkbook.Sheets("Import").Range("A1")
    csvWb.Close False
    
    MsgBox "CSV imported!", vbInformation
End Sub

' Export range to CSV
Sub ExportToCSV()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\Export_" & Format(Now, "yyyymmdd") & ".csv"
    
    ' Copy to new workbook and save as CSV
    Dim exportWb As Workbook
    Set exportWb = Workbooks.Add
    ThisWorkbook.Sheets("Export").UsedRange.Copy exportWb.Sheets(1).Range("A1")
    
    Application.DisplayAlerts = False
    exportWb.SaveAs filePath, xlCSV
    exportWb.Close False
    Application.DisplayAlerts = True
    
    MsgBox "Exported to: " & filePath, vbInformation
End Sub

' Export to PDF
Sub ExportToPDF()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\Report_" & Format(Now, "yyyymmdd") & ".pdf"
    
    ThisWorkbook.Sheets("Report").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False
    
    MsgBox "PDF exported to: " & filePath, vbInformation
End Sub
```

---

## 11. Best Practices for VBA Code

### Code Structure

```vba
Option Explicit  ' ALWAYS use this

' Module-level declarations
Private Const MAX_ROWS As Long = 100000
Private Const REPORT_DATE As String = "2026-09-10"

' Main procedure - entry point
Sub MainProcess()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Call helper procedures
    Call ValidateData
    Call ProcessData
    Call GenerateReport
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Process complete!", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

### Performance Optimization

```vba
Sub OptimizedCode()
    ' 1. Turn off screen updating
    Application.ScreenUpdating = False
    
    ' 2. Turn off automatic calculation
    Application.Calculation = xlCalculationManual
    
    ' 3. Disable events
    Application.EnableEvents = False
    
    ' 4. Use arrays instead of cell-by-cell
    Dim data() As Variant
    data = Range("A1:D10000").Value
    ' Process data in array...
    Range("A1:D10000").Value = data
    
    ' 5. Avoid Select/Activate
    ' BAD:  Range("A1").Select: Selection.Value = 10
    ' GOOD: Range("A1").Value = 10
    
    ' 6. Use With blocks
    With Range("A1").Font
        .Bold = True
        .Size = 12
    End With
    
    ' 7. Use specific objects, not Variant
    Dim ws As Worksheet  ' NOT: Dim ws As Object
    
    ' 8. Disable alerts for batch operations
    Application.DisplayAlerts = False
    ' ... delete sheets, etc.
    Application.DisplayAlerts = True
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

### Naming Conventions

```vba
' Variables: camelCase
Dim lastRow As Long
Dim customerName As String
Dim isValid As Boolean

' Constants: UPPER_CASE
Const MAX_RETRIES As Long = 3
Const TAX_RATE As Double = 0.08

' Procedures: PascalCase
Sub ProcessSalesData()
Function CalculateCommission()

' Controls: prefix + name
' txtName, cboDepartment, btnSubmit, lblTitle

' Meaningful names (avoid: x, tmp, data1, val)
' GOOD: employeeName, totalRevenue, isFirstRun
' BAD:  x, temp, d, v
```

### Documentation

```vba
' ============================================================
' Procedure: ProcessSalesData
' Purpose:   Processes raw sales data and calculates totals
' Author:    John Smith
' Date:      2026-09-10
' Parameters:
'   ws - Source worksheet containing sales data
'   startDate - Filter start date
'   endDate - Filter end date
' Returns:   Number of records processed
' ============================================================
Function ProcessSalesData(ws As Worksheet, startDate As Date, endDate As Date) As Long
    ' ... implementation
End Function
```

### Complete Checklist

- [ ] `Option Explicit` at top of every module
- [ ] All variables declared with specific types
- [ ] Error handling in every procedure
- [ ] ScreenUpdating/Calculation/Events managed
- [ ] No unnecessary Select/Activate
- [ ] Meaningful variable and procedure names
- [ ] Comments explaining "why", not "what"
- [ ] Arrays for large data operations
- [ ] Constants for magic numbers
- [ ] Modular code (small, focused procedures)

---

## 12. Practice Exercises

### Exercise 1: Hello World

Create a macro that displays a message box with your name.

```vba
Sub HelloWorld()
    Dim name As String
    name = InputBox("What is your name?", "Hello")
    If name <> "" Then
        MsgBox "Hello, " & name & "! Welcome to VBA!", vbInformation, "Greeting"
    End If
End Sub
```

### Exercise 2: Range Operations

Create a macro that fills cells A1:A10 with numbers 1-10, then highlights even numbers in green.

```vba
Sub HighlightEvenNumbers()
    Dim i As Long
    
    For i = 1 To 10
        Cells(i, 1).Value = i
        If i Mod 2 = 0 Then
            Cells(i, 1).Interior.Color = RGB(0, 200, 0)
        Else
            Cells(i, 1).Interior.ColorIndex = xlNone
        End If
    Next i
    
    MsgBox "Done! Even numbers highlighted.", vbInformation
End Sub
```

### Exercise 3: Data Validation Loop

Loop through column A and mark rows where sales exceed $1000.

```vba
Sub MarkHighSales()
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow  ' Skip header
        If IsNumeric(Cells(i, 2).Value) Then
            If Cells(i, 2).Value > 1000 Then
                Cells(i, 3).Value = "High"
                Cells(i, 3).Font.Color = RGB(0, 128, 0)
            Else
                Cells(i, 3).Value = "Normal"
            End If
        End If
    Next i
End Sub
```

### Exercise 4: Sheet Consolidator

Combine data from multiple sheets into one summary sheet.

```vba
Sub ConsolidateSheets()
    Dim summaryWs As Worksheet
    Dim ws As Worksheet
    Dim summaryRow As Long
    Dim lastRow As Long
    
    ' Create or clear summary sheet
    On Error Resume Next
    Set summaryWs = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0
    
    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        summaryWs.Name = "Summary"
    End If
    
    summaryWs.Cells.Clear
    summaryWs.Range("A1:D1").Value = Array("Sheet", "Row", "Column A", "Column B")
    summaryRow = 2
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Summary" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow > 1 Then
                ws.Range("A2:B" & lastRow).Copy summaryWs.Cells(summaryRow, 3)
                summaryWs.Range(summaryWs.Cells(summaryRow, 1), _
                    summaryWs.Cells(summaryRow + lastRow - 2, 1)).Value = ws.Name
                summaryRow = summaryRow + lastRow - 1
            End If
        End If
    Next ws
    
    summaryWs.Columns.AutoFit
    MsgBox "Consolidated " & (summaryRow - 2) & " rows!", vbInformation
End Sub
```

### Exercise 5: Random Password Generator (Function)

```vba
Function GeneratePassword(length As Long) As String
    Dim chars As String
    Dim password As String
    Dim i As Long
    
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%"
    password = ""
    
    Randomize  ' Seed random number generator
    
    For i = 1 To length
        password = password & Mid(chars, Int(Rnd() * Len(chars)) + 1, 1)
    Next i
    
    GeneratePassword = password
End Function

' Usage: =GeneratePassword(12) in a cell, or call from another Sub
Sub TestPassword()
    MsgBox "Generated password: " & GeneratePassword(16)
End Sub
```

### Exercise 6: Interactive Form

Build a simple calculator using a UserForm with two text boxes, an operator dropdown, and a result display.

```vba
' UserForm: frmCalculator
' Controls: txtNum1, txtNum2, cboOperator, btnCalculate, lblResult

Private Sub UserForm_Initialize()
    cboOperator.AddItem "+"
    cboOperator.AddItem "-"
    cboOperator.AddItem "*"
    cboOperator.AddItem "/"
    cboOperator.ListIndex = 0
End Sub

Private Sub btnCalculate_Click()
    If Not IsNumeric(txtNum1.Value) Or Not IsNumeric(txtNum2.Value) Then
        MsgBox "Please enter valid numbers.", vbExclamation
        Exit Sub
    End If
    
    Dim num1 As Double, num2 As Double, result As Double
    num1 = CDbl(txtNum1.Value)
    num2 = CDbl(txtNum2.Value)
    
    Select Case cboOperator.Value
        Case "+": result = num1 + num2
        Case "-": result = num1 - num2
        Case "*": result = num1 * num2
        Case "/"
            If num2 = 0 Then
                MsgBox "Cannot divide by zero!", vbExclamation
                Exit Sub
            End If
            result = num1 / num2
    End Select
    
    lblResult.Caption = "Result: " & result
End Sub
```

---

## 13. Video References

### Wise Owl Tutorials (YouTube)
- Excel VBA Introduction: https://www.youtube.com/watch?v=6t8M1bB6r7I
- VBA for Beginners Full Course: https://www.youtube.com/watch?v=GhGg4JkxKEM
- Excel VBA UserForms: https://www.youtube.com/watch?v=KHOJ5_pb5yc

### ExcelMacroMastery (YouTube)
- VBA Tutorial for Beginners: https://www.youtube.com/watch?v=y5Y1Jd_STCE
- VBA Arrays Guide: https://www.youtube.com/watch?v=hyGc9I6xJXg

### Chandoo.org (YouTube)
- Excel VBA Crash Course: https://www.youtube.com/watch?v=kS3fwbJ1cOA
- VBA Loops Tutorial: https://www.youtube.com/watch?v=Aw2bJjj8c5E

### TrumpExcel (YouTube)
- Excel VBA Tutorial for Beginners: https://www.youtube.com/watch?v=Wwg_KbxWfyk
- Recording Macros: https://www.youtube.com/watch?v=JkDKRg_jGgM
- VBA UserForm Tutorial: https://www.youtube.com/watch?v=9WN-2NknhzA

### Leila Gharani (YouTube)
- Excel VBA Tutorial: https://www.youtube.com/watch?v=BaKdFe1JY4I

### Kevin Stratvert (YouTube)
- Excel VBA for Beginners: https://www.youtube.com/watch?v=BaKdFe1JY4I

---

## 14. Sources & Further Reading

### Websites

- **TrumpExcel** — Excel VBA Programming: https://trumpexcel.com/excel-vba/
- **ExcelMacroMastery** — VBA Articles: https://excelmacromastery.com/vba-articles/
- **Chandoo.org** — VBA Tutorials: https://chandoo.org/wp/vba/
- **Microsoft Docs** — VBA Language Reference: https://docs.microsoft.com/en-us/office/vba/api/overview/excel
- **Excel Easy** — VBA Tutorial: https://www.excel-easy.com/vba.html
- **Automate Excel** — VBA Tutorials: https://www.automateexcel.com/vba/

### Books

- *Excel VBA Programming for Dummies* by Michael Alexander & John Walkenbach
- *Power Programming with VBA* by John Walkenbach
- *Excel 2019 Power Programming with VBA* by Michael Alexander & Dick Kusleika
- *Professional Excel Development* by Rob Bovey, Dennis Wallentin, Stephen Bullen

### Free Resources

- ExcelMacroMastery VBA Cheat Sheet: https://excelmacromastery.com/
- Chandoo.org VBA Classes: https://chandoo.org/wp/vba-classes/
- TrumpExcel Free VBA Course: https://trumpexcel.com/excel-vba/

---

> **Next Module:** [Module 8: Real-World Projects](../08-Real-World-Projects/08-real-world-projects.md) — Apply everything you've learned in practical, industry-relevant projects.
