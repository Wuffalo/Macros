Sub FormulaLocationAging()
'
' FormulaLocationAging Macro
'
'
    
    SheetLength = WorksheetFunction.CountA(Sheets("Inventory").Range("A:A"))
    IQC = "P"
    OIC = "Q"
    NIC = "R"

    Range("P1").Value = "Inventory QTY"
    Range("P2:P" & SheetLength).Formula = "=COUNTIF($E$2:$E$"& SheetLength &",E2)"
    Range("P2:P" & SheetLength).Value = Range("P2:P" & SheetLength).Value

    Range("Q1").Value = "Oldest Inventory"
    Range("Q2:Q" & SheetLength).Formula = "=MINIFS($I$2:$I$"& SheetLength &",$E$2:$E$"& SheetLength &",E2)"
    Range("Q2:Q" & SheetLength).Value = Range("Q2:Q" & SheetLength).Value
    Range("Q2:Q" & SheetLength).NumberFormat = "mm/dd/yy h:mm;@"

    Range("R1").Value = "Newest Inventory"
    Range("R2:R" & SheetLength).Formula = "=MAXIFS($I$2:$I$"& SheetLength &",$E$2:$E$"& SheetLength &",E2)"
    Range("R2:R" & SheetLength).Value = Range("R2:R" & SheetLength).Value
    Range("R2:R" & SheetLength).NumberFormat = "mm/dd/yy h:mm;@"

End Sub