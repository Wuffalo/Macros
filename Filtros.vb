
Sheets("Milestone Data").Range("AM2:AM" & ContadorData).Formula = "=IF(AND(C2=""S"",OR(G2=""CA"",G2=""US"")),""CCD"",IF(AND(C2=""C"",OR(G2=""CA"",G2=""US"")),""DD"",IF(OR(D2=""FJZ"",D2=""FTX"",D2=""TAU""),""3B18PR"",IF(C2=""C"",""3B2SC"",IF(C2=""S"",""DOCGEN"",)))))"
Sheets("Milestone Data").Range("AM2:AM" & ContadorData) = Sheets("Milestone Data").Range("AM2:AM" & ContadorData).Value

Consol
Sheets("Milestone Data").Range("O2:O" & ContadorData).Formula = "=IF(COUNTIF(N:N,N2)=0,""No"",COUNTIF(N:N,N2))"
Sheets("Milestone Data").Range("O2:O" & ContadorData) = Sheets("Milestone Data").Range("O2:O" & ContadorData).Value

