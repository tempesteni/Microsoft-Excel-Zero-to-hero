# Quiz 07 — VBA Macros

**Module:** VBA Macros, Loops, UserForms, Error Handling
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** What keyboard shortcut opens the VBA Editor?
- A) `Alt + F11`
- B) `Ctrl + M`
- C) `F5`
- D) `Alt + F8`

**2.** You record a macro that formats cell A1 bold. You then run it. Cell B5:
- A) Also becomes bold
- B) Stays unchanged — the macro only affects A1
- C) Gets deleted
- D) Shows an error

**3.** In VBA, what does `MsgBox "Hello"` display?
- A) A text box on the worksheet
- B) A pop-up dialog box with the text "Hello"
- C) Writes "Hello" to cell A1
- D) Prints "Hello" to the Immediate window

**4.** Which loop structure runs at least once before checking the condition?
- A) `For...Next`
- B) `Do While...Loop`
- C) `Do...Loop While`
- D) `For Each...Next`

**5.** You want to loop through all cells in column A from A1 to A100. Which is correct?
- A) `For i = 1 To 100: Cells(i, 1).Value = ... : Next i`
- B) `For Each cell In Range("A1:A100"): cell.Value = ... : Next`
- C) Both A and B
- D) Neither — you must use `Do While`

**6.** What does `Application.ScreenUpdating = False` do?
- A) Hides Excel completely
- B) Prevents screen redraws during macro execution, speeding it up
- C) Disables all user input
- D) Turns off the formula bar

**7.** A UserForm has a TextBox named `txtName` and a CommandButton. To read the user's input in the button's click event:
- A) `MsgBox txtName.Value`
- B) `MsgBox txtName.Text`
- C) Both A and B work
- D) `MsgBox txtName.Caption`

**8.** What does `On Error Resume Next` do?
- A) Stops the macro on any error
- B) Skips the error-causing line and continues execution
- C) Shows a custom error message
- D) Deletes the error

**9.** Which statement correctly declares a variable in VBA?
- A) `Dim x As Integer`
- B) `var x = Integer`
- C) `Integer x`
- D) `Declare x As Integer`

**10.** You want to call a macro named `ProcessData` from another macro. The correct syntax is:
- A) `Call ProcessData`
- B) `ProcessData`
- C) `Run ProcessData`
- D) Both A and B

**11.** What is the difference between `Sub` and `Function` in VBA?
- A) There is no difference
- B) A Sub performs actions; a Function returns a value
- C) A Function is faster
- D) A Sub can't be called from a worksheet cell

**12.** You want a macro to run automatically when the workbook opens. Name the sub:
- A) `Auto_Open()`
- B) `Workbook_Open()`
- C) `Sub Start()`
- D) Both A and B work

**13.** What does `Cells(Rows.Count, 1).End(xlUp).Row` return?
- A) Always row 1
- B) The last used row in column A
- C) The total number of rows
- D) An error

**14.** To add a dropdown list to a UserForm, you use:
- A) TextBox
- B) ComboBox or ListBox
- C) CheckBox
- D) Label

**15.** What does `Err.Number` give you after an error occurs?
- A) The line number where the error happened
- B) The error code identifying the type of error
- C) The number of errors so far
- D) A random number

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

Write the VBA code that accomplishes the task.

**16.** Write a Sub that loops through cells A1:A10 and makes any cell with a value greater than 100 bold.

> **Your answer:** `_________________________________`

**17.** Write a `For Each` loop that clears all cells in the range B2:B50 that contain the text "TBD".

> **Your answer:** `_________________________________`

**18.** Write a Function called `DoubleIt` that takes a number as input and returns double the value.

> **Your answer:** `_________________________________`

**19.** Write the error handling block that shows a custom message box when an error occurs, then exits the Sub.

> **Your answer:** `_________________________________`

**20.** Write VBA code that creates a new worksheet named "Summary" at the end of the workbook (handling the case where it already exists).

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **A** | `Alt+F11` opens VBA Editor; `Alt+F8` opens the macro list dialog. |
| 2 | **B** | Recorded macros use absolute references by default — only the recorded cells are affected. |
| 3 | **B** | MsgBox displays a modal dialog box to the user. |
| 4 | **C** | `Do...Loop While` executes the body first, then checks the condition. |
| 5 | **C** | Both `For i` with `Cells()` and `For Each` with a range work. |
| 6 | **B** | Disabling screen updating prevents flicker and speeds up macros significantly. |
| 7 | **C** | Both `.Value` and `.Text` return the TextBox content. |
| 8 | **B** | `On Error Resume Next` suppresses errors and continues to the next line. |
| 9 | **A** | `Dim x As Integer` is the standard VBA variable declaration. |
| 10 | **D** | Both `Call ProcessData` and just `ProcessData` invoke the sub. |
| 11 | **B** | Subs execute code; Functions return values and can be used in expressions. |
| 12 | **D** | Both `Auto_Open()` and `Workbook_Open()` run on open (Workbook_Open is preferred). |
| 13 | **B** | Starting from the bottom of column A and pressing Ctrl+Up gives the last used row. |
| 14 | **B** | ComboBox (editable dropdown) or ListBox (scrollable list) provides dropdown behavior. |
| 15 | **B** | `Err.Number` is the numeric error code (e.g., 1004 for application-defined errors). |
| 16 | ```vba Sub BoldOver100() Dim cell As Range For Each cell In Range("A1:A10") If cell.Value > 100 Then cell.Font.Bold = True Next cell End Sub``` | — |
| 17 | ```vba Sub ClearTBD() Dim cell As Range For Each cell In Range("B2:B50") If cell.Value = "TBD" Then cell.ClearContents Next cell End Sub``` | — |
| 18 | ```vba Function DoubleIt(num As Double) As Double DoubleIt = num * 2 End Function``` | — |
| 19 | ```vba Sub MyMacro() On Error GoTo ErrHandler ' ... code ... Exit Sub ErrHandler: MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical Exit Sub End Sub``` | — |
| 20 | ```vba Sub CreateSummary() Dim ws As Worksheet On Error Resume Next Set ws = ThisWorkbook.Sheets("Summary") On Error GoTo 0 If ws Is Nothing Then Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)) ws.Name = "Summary" End If End Sub``` | — |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 8
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 7
