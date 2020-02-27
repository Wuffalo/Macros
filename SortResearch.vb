Sub Sort()
'
' Sort Macro
'
' Keyboard Shortcut: Ctrl+r
'
    OCS_length = WorksheetFunction.CountA(Sheets("OrderCurrentStatus").Range("A:A"))
    ASN_length = WorksheetFunction.CountA(Sheets("ASN").Range("A:A"))
    Sheets("OrderCurrentStatus").Select
    Columns("C:C").Delete Shift:=xlToLeft
    Columns("F:F").Delete Shift:=xlToLeft
    Range("H1").Value = "ASN"
    Range("I1").Value = "Comment"
    Cells.EntireColumn.AutoFit
    Sheets("ASN").Select
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
        "I1:I" & ASN_length), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
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
        Add2 Key:=Range("E1:E" & OCS_length), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A$1:$I$" & OCS_length).RemoveDuplicates Columns:=2, Header:=xlNo
    ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort.SortFields. _
        Clear
    OCS_length = WorksheetFunction.CountA(Sheets("OrderCurrentStatus").Range("A:A")) ' Set new reduced size
    ActiveWorkbook.Worksheets("OrderCurrentStatus").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("C1:C" & OCS_length), SortOn:=xlSortOnValues, Order:=xlAscending, _
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
        Add2 Key:=Range("F1:F" & OCS_length), SortOn:=xlSortOnValues, Order:=xlAscending, _
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
    Range("H2:H" & OCS_length).Formula = "=VLOOKUP(B2,ASN!F:I,4,0)"
    Range("H2:H" & OCS_length).Value = Range("H2:H" & OCS_length).Value
End Sub


