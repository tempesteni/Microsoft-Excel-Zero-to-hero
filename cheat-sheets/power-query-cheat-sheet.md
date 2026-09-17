# Power Query Cheat Sheet — M Language & Transformations

> **Quick Reference** | Print on A4/Letter | 1-2 pages

---

## GETTING STARTED

| Action | How |
|--------|-----|
| Open Power Query Editor | **Data → Get Data → Launch Power Query Editor** |
| Import from file | **Data → Get Data → From File** (CSV, Excel, JSON, XML, etc.) |
| Import from database | **Data → Get Data → From Database** (SQL Server, Access, etc.) |
| Import from web | **Data → Get Data → From Web** |
| Append queries | **Home → Append Queries** (stack tables vertically) |
| Merge queries | **Home → Merge Queries** (join tables like VLOOKUP) |
| Close & Load | **Home → Close & Load** (returns data to Excel) |

---

## TOP 20 M LANGUAGE FUNCTIONS

### Table Functions

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `Table.SelectRows` | `Table.SelectRows(table, condition)` | Filter rows by condition | `= Table.SelectRows(Source, each [Status] = "Active")` |
| `Table.SelectColumns` | `Table.SelectColumns(table, columns)` | Keep only listed columns | `= Table.SelectRows(Source, {"Name","Sales"})` |
| `Table.RemoveColumns` | `Table.RemoveColumns(table, columns)` | Remove listed columns | `= Table.RemoveColumns(Source, {"Temp","Debug"})` |
| `Table.AddColumn` | `Table.AddColumn(table, name, function)` | Add computed column | `= Table.AddColumn(Source, "Tax", each [Price] * 0.08)` |
| `Table.RenameColumns` | `Table.RenameColumns(table, list)` | Rename columns | `= Table.RenameColumns(Source, {{"old","new"}})` |
| `Table.Sort` | `Table.Sort(table, ordering)` | Sort by columns | `= Table.Sort(Source, {{"Sales", Order.Descending}})` |
| `Table.Group` | `Table.Group(table, key, agg)` | Group and aggregate | `= Table.Group(Source, {"Region"}, {{"Total", each List.Sum([Sales]), type number}})` |
| `Table.TransformColumnTypes` | `Table.TransformColumnTypes(table, types)` | Set column data types | `= Table.TransformColumnTypes(Source, {{"Date", type date}})` |
| `Table.NestedJoin` | `Table.NestedJoin(t1, key1, t2, key2, name)` | Left join two tables | `= Table.NestedJoin(Orders, {"ID"}, Customers, {"ID"}, "Customer")` |
| `Table.Combine` | `Table.Combine({table1, table2, …})` | Append/stack tables | `= Table.Combine({Q1, Q2, Q3, Q4})` |

### List Functions

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `List.Sum` | `List.Sum(list)` | Sum of list values | `= List.Sum({1, 2, 3})` → 6 |
| `List.Average` | `List.Average(list)` | Average of list | `= List.Average({10, 20, 30})` → 20 |
| `List.Count` | `List.Count(list)` | Count items | `= List.Count({1, 2, 3})` → 3 |
| `List.Min` / `List.Max` | `List.Min(list)` | Minimum / Maximum value | `= List.Max({5, 2, 8})` → 8 |
| `List.Distinct` | `List.Distinct(list)` | Remove duplicates | `= List.Distinct({"A","B","A"})` → {"A","B"} |

### Text Functions

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `Text.Length` | `Text.Length(text)` | Character count | `= Text.Length("Hello")` → 5 |
| `Text.Split` | `Text.Split(text, delimiter)` | Split text into list | `= Text.Split("a-b-c", "-")` → {"a","b","c"} |
| `Text.Combine` | `Text.Combine(list, delimiter)` | Join list into text | `= Text.Combine({"a","b"}, ",")` → "a,b" |
| `Text.Lower` / `Upper` | `Text.Lower(text)` | Change case | `= Text.Upper("hello")` → "HELLO" |
| `Text.Trim` | `Text.Trim(text)` | Remove leading/trailing spaces | `= Text.Trim(" hi ")` → "hi" |

---

## COMMON TRANSFORMATIONS (GUI → M Code)

### Filter Rows
```
// Via GUI: Click column header filter dropdown → select values
// M Code generated:
= Table.SelectRows(Source, each [Region] = "East")

// Multiple conditions:
= Table.SelectRows(Source, each [Region] = "East" and [Sales] > 1000)

// Contains text:
= Table.SelectRows(Source, each Text.Contains([Name], "Corp"))

// Not null:
= Table.SelectRows(Source, each [Email] <> null)
```

### Add Conditional Column
```
// GUI: Add Column → Conditional Column
// M Code:
= Table.AddColumn(Source, "Tier", each
    if [Sales] >= 10000 then "Gold"
    else if [Sales] >= 5000 then "Silver"
    else "Bronze"
)
```

### Merge (Join) Tables
```
// GUI: Home → Merge Queries → select key columns → Join kind
// M Code (Left Outer Join):
= Table.NestedJoin(Orders, {"CustomerID"}, Customers, {"CustomerID"}, "Customer", JoinKind.LeftOuter)

// Join types: LeftOuter, RightOuter, FullOuter, Inner, LeftAnti, RightAnti
```

### Pivot / Unpivot
```
// Unpivot (wide → long): Select columns → Transform → Unpivot Columns
= Table.UnpivotOtherColumns(Source, {"ID","Name"}, "Attribute", "Value")

// Pivot (long → wide): Select column → Transform → Pivot Column
= Table.Pivot(Source, List.Distinct(Source[Month]), "Month", "Sales", List.Sum)
```

### Replace Values
```
// GUI: Right-click column → Replace Values
= Table.ReplaceValue(Source, "old", "new", Replacer.ReplaceText, {"ColumnName"})

// Replace nulls:
= Table.ReplaceValue(Source, null, 0, Replacer.ReplaceValue, {"Sales"})
```

### Custom Column with Text Extraction
```
// Extract domain from email:
= Table.AddColumn(Source, "Domain", each Text.AfterDelimiter([Email], "@"))

// Extract first 3 characters:
= Table.AddColumn(Source, "Code", each Text.Start([ID], 3))

// Parse JSON column:
= Table.AddColumn(Source, "Parsed", each Json.Document([JsonCol]))
```

---

## POWER QUERY PARAMETERS & VARIABLES

### Define a Parameter
1. **Home → Manage Parameters → New Parameter**
2. Set Name, Type, Default Value
3. Reference in code: `= Table.SelectRows(Source, each [Date] >= Parameter1)`

### Reference Previous Steps
Each step is a variable. Reference by name:
```
let
    Source = Excel.Workbook(File.Contents("data.xlsx")),
    Sheet1 = Source{[Name="Sheet1"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet1),
    Filtered = Table.SelectRows(Promoted, each [Status] = "Active")
in
    Filtered
```

---

## ERROR HANDLING

| Pattern | M Code |
|---------|--------|
| Replace errors with value | `= Table.ReplaceErrorValues(Source, {"Col", 0})` |
| Try/catch | `= try [risky_operation] otherwise "default"` |
| Remove error rows | `= Table.RemoveRowsWithErrors(Source, {"Col"})` |
| Check for null | `= if [Value] = null then 0 else [Value]` |

---

## PERFORMANCE TIPS

| Tip | Why |
|-----|-----|
| **Filter early** | Remove rows/columns as first steps — reduces data flowing through pipeline |
| **Disable auto data type detection** | File → Options → uncheck "Auto-detect column types" (prevents auto-changes) |
| **Use `Table.Buffer`** | `= Table.Buffer(Source)` — caches result in memory to avoid re-evaluation |
| **Fold queries** | Ensure filters push to source (check via View → Query Dependencies → right-click → View native query) |
| **Avoid row-by-row** | Prefer column-level transforms over `Table.AddColumn` with expensive per-row logic |
| **Reference, don't duplicate** | Base new queries on existing steps rather than re-loading source data |

---

## 📋 Quick Tips

- **`fx` button** in formula bar: Insert step → write M code
- **Applied Steps pane** (right): Click any step to preview at that point
- **Advanced Editor**: Home → Advanced Editor — see/edit full M code
- **Duplicate query**: Right-click query → Duplicate (for variations without reloading)
- **`Ctrl+Click`** column headers to select multiple columns for batch transforms
- All M functions are **case-sensitive**: `Table.Group` ≠ `table.group`
- Use `each` shorthand: `each [Column] > 10` is shorthand for `(_ as record) => _[Column] > 10`
