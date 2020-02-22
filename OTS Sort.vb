Sub Sort()
'
' Sort Macro
'
' Keyboard Shortcut: Ctrl+r
'
    Sheets("ASN").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "ASN"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Comment"
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ASN").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("C:C").Select
    Selection.Copy
    Columns("I:I").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("ASN").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ASN").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "I1:I517576"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ASN").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Application.CutCopyMode = False
    Columns("C:C").Select
    Selection.Copy
    Columns("I:I").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("ASN").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ASN").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "I1:I406043"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ASN").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    
    Sheets("OrderCurrentStatus").Select
    Range("A1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("E1:E63"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A$1:$I$63").RemoveDuplicates Columns:=2, Header:=xlNo
    ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("C1:C33"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("F1:F33"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    Range("H2").Select
End Sub
