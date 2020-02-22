Public TodayD As String
Public TomorrowD As String


Sub SortReport()
'
' SortReport Macro
'
' Keyboard Shortcut: Ctrl+r
'
    '''Dim TodayD As String
    '''Dim TomorrowD As String
    Dim AgedTabName As String
    Dim DSLCTab As String
    Dim TableOTS As String
    Dim TableOTSp1 As String
    Dim TableOTS2 As String
    Dim TableOTSp12 As String
    Dim TableOTS3 As String
    
    TodayD = "0207"                                             'CHANGE THIS TO TODAY'S TAB NAME"
    TomorrowD = "0210"                                          'CHANGE THIS TO TOMORROW'S TAB NAME"
    AgedTabName = "Aged"                                            'CHANGE THIS TO THE NAME OF THE AGED TAB
    DSLCTab = "DSLC"
    
    TableOTS2 = "OTSTable"                                          'CHANGE THIS TO CURRENT TABLE # FOR OTS TAB
    TableOTSp12 = "OTSTableT"                                       'CHANGE THIS TO CURRENT TABLE # FOR OTS+1 TAB
    
    
    Dim OTS As Object
    Dim OTSp1 As Object
    Dim Aged As Object
    Dim DSLC As Object
    
    Set OTS = ActiveWorkbook.Worksheets(TodayD)
    Set OTSp1 = ActiveWorkbook.Worksheets(TomorrowD)
    Set Aged = ActiveWorkbook.Worksheets(AgedTabName)
    Set DSLC = ActiveWorkbook.Worksheets(DSLCTab)
    
    Sheets(TodayD).Select
    ActiveSheet.ListObjects(TableOTS2).Range.AutoFilter Field:=4, Criteria1:= _
        Array("#N/A", "In Picking", "Part Allocated", "Loaded", "Not Started", "Pack Ready", "Released", "Allocated", "Created Externally", "In Packing"), _
        Operator:=xlFilterValues
    Application.CutCopyMode = False
    OTS.ListObjects(TableOTS2).Sort.SortFields.Clear
    OTS.ListObjects(TableOTS2).Sort.SortFields.Add _
        Key:=Range("OTSTable[WMS Status]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    OTS.ListObjects(TableOTS2).Sort.SortFields.Add _
        Key:=Range("OTSTable[Carrier]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    OTS.ListObjects(TableOTS2).Sort.SortFields.Add _
        Key:=Range("OTSTable[Customer]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    OTS.ListObjects(TableOTS2).Sort.SortFields.Add( _
        Range("OTSTable[Comments]"), xlSortOnCellColor, xlAscending, , xlSortNormal). _
        SortOnValue.Color = RGB(255, 255, 204)
    With OTS.ListObjects(TableOTS2).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'    Sheets("0705").Select
'    ActiveSheet.ListObjects("DSLCTable").Range.AutoFilter Field:=4, Criteria1:= _
'        Array("Allocated", "Created Externally", "In Picking", "Loaded", "Not Started", _
'        "Pack Ready", "Released"), Operator:=xlFilterValues
'    DSLC.ListObjects("DSLCTable").Sort.SortFields.Clear
'    DSLC.ListObjects("DSLCTable").Sort.SortFields.Add _
'        Key:=Range("DSLCTable[WMS Status]"), SortOn:=xlSortOnValues, Order:= _
'        xlAscending, DataOption:=xlSortNormal
'    DSLC.ListObjects("DSLCTable").Sort.SortFields.Add _
'        Key:=Range("DSLCTable[Carrier]"), SortOn:=xlSortOnValues, Order:=xlAscending _
'        , DataOption:=xlSortNormal
'    DSLC.ListObjects("DSLCTable").Sort.SortFields.Add _
'        Key:=Range("DSLCTable[Customer]"), SortOn:=xlSortOnValues, Order:= _
'        xlAscending, DataOption:=xlSortNormal
'    DSLC.ListObjects("DSLCTable").Sort.SortFields.Add( _
'        Range("DSLCTable[Comments]"), xlSortOnCellColor, xlAscending, , xlSortNormal). _
'        SortOnValue.Color = RGB(255, 255, 204)
'    With DSLC.ListObjects("DSLCTable").Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
    Sheets(TomorrowD).Select
    ActiveSheet.ListObjects("OTSTableT").Range.AutoFilter Field:=4, Criteria1:= _
        Array("Allocated", "Created Externally", "Part Allocated", "In Packing", "In Picking", "Loaded", _
        "Not Started", "Pack Ready"), Operator:=xlFilterValues
    OTSp1.ListObjects("OTSTableT").Sort.SortFields.Clear
    OTSp1.ListObjects("OTSTableT").Sort.SortFields.Add _
        Key:=Range("OTSTableT[WMS Status]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    OTSp1.ListObjects("OTSTableT").Sort.SortFields.Add _
        Key:=Range("OTSTableT[Carrier]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    OTSp1.ListObjects("OTSTableT").Sort.SortFields.Add _
        Key:=Range("OTSTableT[Customer]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    OTSp1.ListObjects("OTSTableT").Sort.SortFields.Add( _
        Range("OTSTableT[Comments]"), xlSortOnCellColor, xlAscending, , xlSortNormal). _
        SortOnValue.Color = RGB(255, 255, 204)
    With OTSp1.ListObjects("OTSTableT").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Aged").Select
    ActiveSheet.ListObjects("AGEDTable").Range.AutoFilter Field:=4, Criteria1:= _
        Array("Allocated", "Created Externally", "In Packing", "Part Allocated", "In Picking", "Loaded", "Not Started", _
        "Pack Ready", "Released"), Operator:=xlFilterValues
    Aged.ListObjects("AGEDTable").Sort.SortFields.Clear
    Aged.ListObjects("AGEDTable").Sort.SortFields.Add _
        Key:=Range("AGEDTable[WMS Status]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    Aged.ListObjects("AGEDTable").Sort.SortFields.Add _
        Key:=Range("AGEDTable[Carrier]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    Aged.ListObjects("AGEDTable").Sort.SortFields.Add _
        Key:=Range("AGEDTable[Customer]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    Aged.ListObjects("AGEDTable").Sort.SortFields.Add( _
        Range("AGEDTable[Comments]"), xlSortOnCellColor, xlAscending, , xlSortNormal). _
        SortOnValue.Color = RGB(255, 255, 204)
    With Aged.ListObjects("AGEDTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets(TodayD).Select
    Range("OTSTable[[#Headers],[SO-SS]]").Select
End Sub

