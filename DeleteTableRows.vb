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
