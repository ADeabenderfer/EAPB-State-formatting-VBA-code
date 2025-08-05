'Macro created by Amber Deabenderfer in 2020
Sub BalAdjState()

'Deliminate
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array( _
        25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), _
        Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array( _
        38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), _
        Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array(49, 1), Array(50, 1), Array( _
        51, 1), Array(52, 1), Array(53, 1), Array(54, 1), Array(55, 1), Array(56, 1), Array(57, 1), _
        Array(58, 1), Array(59, 1), Array(60, 1), Array(61, 1), Array(62, 1), Array(63, 1), Array( _
        64, 1), Array(65, 1), Array(66, 1), Array(67, 1), Array(68, 1), Array(69, 1), Array(70, 1), _
        Array(71, 1), Array(72, 1), Array(73, 1), Array(74, 1), Array(75, 1), Array(76, 1), Array( _
        77, 1), Array(78, 1), Array(79, 1), Array(80, 1), Array(81, 1), Array(82, 1), Array(83, 1), _
        Array(84, 1), Array(85, 1), Array(86, 1), Array(87, 1), Array(88, 1), Array(89, 1), Array( _
        90, 1), Array(91, 1), Array(92, 1), Array(93, 1), Array(94, 1), Array(95, 1), Array(96, 1), _
        Array(97, 1), Array(98, 1), Array(99, 1), Array(100, 1), Array(101, 1), Array(102, 1), _
        Array(103, 1), Array(104, 1), Array(105, 1), Array(106, 1), Array(107, 1), Array(108, 1), _
        Array(109, 1), Array(110, 1), Array(111, 1), Array(112, 1), Array(113, 1), Array(114, 1), _
        Array(115, 1), Array(116, 1), Array(117, 1), Array(118, 1), Array(119, 1), Array(120, 1), _
        Array(121, 1), Array(122, 1), Array(123, 1), Array(124, 1), Array(125, 1), Array(126, 1), _
        Array(127, 1), Array(128, 1), Array(129, 1), Array(130, 1), Array(131, 1), Array(132, 1)) _
        , TrailingMinusNumbers:=True

'Freeze top view
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

'Add Filters
    Rows("1:1").Select
    Selection.AutoFilter

'Center/Align headers
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireColumn.AutoFit
    Range("K2:OS1048576").Select
    Selection.ColumnWidth = 9.86

'Hide front columns
Dim Col As String, cfind As Range
Worksheets(1).Activate

On Error Resume Next
Col = "Payroll Name"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
cfind.Select
Selection.EntireColumn.Hidden = True

Col = "Balance Dimension"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
cfind.Select
Selection.EntireColumn.Hidden = True

Col = "Run Type"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
cfind.Select
Selection.EntireColumn.Hidden = True

Col = "Process Type"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
cfind.Select
Selection.EntireColumn.Hidden = True

Col = "State Worked Hours"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
cfind.Select
Selection.EntireColumn.Hidden = True

'Size # columns
  Dim Lrow As Long
  Dim lColumn As Long
  Dim lCell As Range

Col = "Process Date"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)

  Lrow = Range("A1").End(xlDown).Row
  lColumn = Range("A1").End(xlToRight).column

Set lCell = Cells(Lrow, lColumn)

Range(cfind, lCell).Select
Selection.ColumnWidth = 9.86

'START OF FILTERING
Col = "State"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)

    ActiveWorkbook.Worksheets("Employee Active Payroll Balance").AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Employee Active Payroll Balance").AutoFilter.Sort. _
        SortFields.Add2 Key:=cfind, SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Employee Active Payroll Balance").AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Col = "Process Date"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
cfind.EntireColumn.Select

    ActiveWorkbook.Worksheets("Employee Active Payroll Balance").AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Employee Active Payroll Balance").AutoFilter.Sort. _
        SortFields.Add2 Key:=cfind, SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Employee Active Payroll Balance").AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


'Insert SUI Exempt column
Dim Found As Range, Found1 As Range, Found2 As Range, i As Integer, j As Integer
Dim LR As Long
Set Found1 = Rows(1).Find(What:="SUI Employer Gross", LookIn:=xlValues, LookAt:=xlWhole)
Set Found2 = Rows(1).Find(What:="SUI Employer Subject Withholdable", LookIn:=xlValues, LookAt:=xlWhole)

Set Found = Rows(1).Find(What:="SUI Employer Taxable", LookIn:=xlValues, LookAt:=xlWhole)
If Found Is Nothing Then Exit Sub
LR = Cells(Rows.Count, Found.column).End(xlUp).Row
Found.Offset(, 1).EntireColumn.Insert
Cells(1, Found.column + 1).Value = "SUI Employer Exempt (Gross - Subject Withholdable)"

i = Range(Found, Found1).Cells.Count
j = Range(Found, Found2).Cells.Count
Formula = "=IF(RC[-" & i & "]="""","""",RC[-" & i & "]-RC[-" & j & "])"
Range(Cells(2, Found.column + 1), Cells(LR, Found.column + 1)).FormulaR1C1 = Formula


'Format cells
    Cells.Select
    Selection.NumberFormat = "#,##0.00"

    Col = "Process Date"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    cfind.ColumnWidth = 10.55
    cfind.EntireColumn.Select
    Selection.NumberFormat = "m/d/yyyy"

    Col = "Person Number"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    cfind.ColumnWidth = 8.36
    cfind.EntireColumn.Select
    Selection.NumberFormat = "0000000"

    Col = "Payroll Relationship Number"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    cfind.ColumnWidth = 8.36
    cfind.EntireColumn.Select
    Selection.NumberFormat = "0000000"

    Col = "State"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    cfind.ColumnWidth = 4.45


'Add rows for SUI ER Excess/Taxable
    Rows("1:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

'Add Metro Report Alert
Col = "Metro Employee Reduced Subject Withholdable"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole).Offset(-1, 0)
cfind.Select
ActiveCell.FormulaR1C1 = "Pull Metro Report!"
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True

'Calculate SUI ER Excess/Taxable
Col = "SUI Employer Excess"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole).Offset(-2, 0)
cfind.Select
ActiveCell.FormulaR1C1 = "SUI Excess Total"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole).Offset(-1, 0)
cfind.Select
ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[4999]C)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

Col = "SUI Employer Taxable"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole).Offset(-2, 0)
cfind.Select
ActiveCell.FormulaR1C1 = "SUI Taxable Total"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole).Offset(-1, 0)
cfind.Select
ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[99999]C)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
On Error GoTo 0


'locate 0 balance columns
    Rows("1:1").Select
    Range("M1").Activate
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove


'locate pretax and 0 balance columns
    Dim formStart As Range, formEnd As Range

    Col = "Process Date"
    Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
    Set formStart = cfind.Offset(-3, 1)
    Set formEnd = cfind.End(xlToRight).Offset(-3, 0)

    formStart.Select
    ActiveCell.FormulaR1C1 = "=IF((IFERROR(SEARCH(""other pretax"",R[3]C,1),-1))>0,IF(SUM(R[4]C:R[99999]C)<>0,SUM(R[4]C:R[99999]C),MAX(R[4]C:R[99999]C)),IF((IFERROR(SEARCH(""pretax"",R[3]C,1),-1))>0,0,IF(SUM(R[4]C:R[99999]C)<>0,SUM(R[4]C:R[99999]C),MAX(R[4]C:R[99999]C))))"

    formStart.Select
    Selection.AutoFill Destination:=Range(formStart, formEnd), Type:=xlFillDefault


    Rows("1:1").Select
    Selection.NumberFormat = "#,##0.00"

'Hide 0 balances
Dim cell As Range
For Each cell In ActiveWorkbook.ActiveSheet.Rows("1").Cells
If cell.Value = "0" Then
cell.EntireColumn.Hidden = True
End If
Next cell

'Remove row 1 calculations
    Rows("1:1").Select
    Range("W1").Activate
    Selection.Delete Shift:=xlUp
    Range("E3").Select


'Color Headers
    Rows("3:3").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Family Leave Insurance Employee", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Family Leave Insurance Employer", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Long Term Care Employee", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Medical Leave Insurance Employee", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "Medical Leave Insurance Employer", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False


    Selection.FormatConditions.Add Type:=xlTextString, String:="Metro Employee" _
        , TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="SDI Employee", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="SDI Employer", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    
    Selection.FormatConditions.Add Type:=xlTextString, String:="SIT", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False


    Selection.FormatConditions.Add Type:=xlTextString, String:="SUI Employee", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="SUI Employer", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Selection.FormatConditions.Add Type:=xlTextString, String:= _
        "State Transit Tax", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="VPDI Employee" _
        , TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99481185338908E-02
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="VPDI Employer" _
        , TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249946592608417
    End With
    Selection.FormatConditions(1).StopIfTrue = False

'Hide State Worked Hours
On Error Resume Next
Col = "State Worked Hours"
Set cfind = Cells.Find(What:=Col, LookAt:=xlWhole)
cfind.Select
Selection.EntireColumn.Hidden = True

'Rename Tab
    Sheets("Employee Active Payroll Balance").Select
    Sheets("Employee Active Payroll Balance").Name = "State"
    Range("B4").Select


End Sub
