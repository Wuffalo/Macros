
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

On Error GoTo ErrCode1:

Workbooks.Open ("S:\Operations\Operations Lead\Pick by Aisle.xlsb")
Windows("Pick by Aisle.xlsb").Activate
Range("a1:ag30000").Copy
Windows("Operations Dashboard.xlsb").Activate
Sheets("DATABASE").Select
Range("d1").Select
Selection.PasteSpecial Paste:=xlValues
Application.CutCopyMode = False

Windows("Shipment Order Summary (PICK ZONE).csv").Close SaveChanges = False

Workbooks.Open ("S:\Operations\Data\Comments&Priority.xlsb")
Windows("Comments&Priority.xlsb").Activate
Windows("Comments&Priority.xlsb").Close SaveChanges = False


ErrCode1:
MsgBox "Could not find 'Shipment Order Summary (PICK ZONE).csv' in file path 'S:\Operations\Data'"

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

---

Workbooks.Open filename:="X:\business\2014\Easy*"
Workbooks.Open filename:=ActiveWorkbook.Path & "\302113*"

---

If Range("D11") = 0 Then

Code 1

ElseIf Range("D11") = 1 Then

Code 2

ElseIf Range("D11") > 1 Then

Code 3

End If
