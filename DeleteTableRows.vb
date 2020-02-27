'   https://stackoverflow.com/questions/20663491/delete-all-data-rows-from-an-excel-table-apart-from-the-first

Sub DeleteTableRows(ByRef Table As ListObject)
    On Error Resume Next
    '~~> Clear Header Row `IF` it exists
    Table.DataBodyRange.Rows(1).ClearContents
    '~~> Delete all the other rows `IF `they exist
    Table.DataBodyRange.Offset(1, 0).Resize(Table.DataBodyRange.Rows.Count - 1, _
    Table.DataBodyRange.Columns.Count).Rows.Delete
    On Error GoTo 0
End Sub


https://www.thespreadsheetguru.com/blog/2014/6/20/the-vba-guide-to-listobject-excel-tables
--Delete all rows except 1st row 

Sub ResetTable()

Dim tbl As ListObject

Set tbl = ActiveSheet.ListObjects("Table1")
'Delete all table rows except first row
  With tbl.DataBodyRange
    If .Rows.Count > 1 Then
      .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows.Delete
    End If
  End With
'Clear out data from first table row
  tbl.DataBodyRange.Rows(1).ClearContents
End Sub

--same but preserve formulas
Sub ResetTable()

Dim tbl As ListObject

Set tbl = ActiveSheet.ListObjects("Table1")
'Delete all table rows except first row
  With tbl.DataBodyRange
    If .Rows.Count > 1 Then
      .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows.Delete
    End If
  End With
'Clear out data from first table row (retaining formulas)
  tbl.DataBodyRange.Rows(1).SpecialCells(xlCellTypeConstants).ClearContents
End Sub

--
Mine
Sub DeleteTableRows(ByRef Table As ListObject)
    On Error Resume Next
    '~~> Clear Header Row `IF` it exists
    Table.DataBodyRange.Rows(1).SpecialCells(xlCellTypeConstants).ClearContents
    '~~> Delete all the other rows `IF `they exist
    Table.DataBodyRange.Offset(1, 0).Resize(Table.DataBodyRange.Rows.Count - 1, _
    Table.DataBodyRange.Columns.Count).Rows.Delete
    On Error GoTo 0
End Sub