Sub ImportFiles()
'
' ImportFiles Macro
'
    
    Workbooks.Open ("C:\Users\WMINSKEY\OneDrive - Schenker AG\Documents\OTS Research\Files\Find ASN.csv")
    Windows("Find ASN.csv").Activate
    ActiveWorkbook.Sheets("Find ASN").UsedRange.Copy
    Windows("Sort_UOV.xlsb").Activate
    ActiveWorkbook.Sheets("ASN").Select
    ActiveWorkbook.Sheets("ASN").Range("A1").Select
    ActiveWorkbook.Sheets("ASN").Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A1").Select
    Windows("Find ASN.csv").Close SaveChanges = False

    Workbooks.Open ("C:\Users\WMINSKEY\OneDrive - Schenker AG\Documents\OTS Research\Files\ordercurrentstatus.csv")
    Windows("ordercurrentstatus.csv").Activate
    ActiveWorkbook.Sheets("ordercurrentstatus").UsedRange.Copy
    Windows("Sort_UOV.xlsb").Activate
    ActiveWorkbook.Sheets("OrderCurrentStatus").Select
    ActiveWorkbook.Sheets("OrderCurrentStatus").Range("A1").Select
    ActiveWorkbook.Sheets("OrderCurrentStatus").Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A1").Select
    Windows("ordercurrentstatus.csv").Close SaveChanges = False

End Sub