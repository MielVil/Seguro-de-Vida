Attribute VB_Name = "Module4"
Sub CRE()
Attribute CRE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CRE Macro
'

'
Dim i As Integer
    i = 0
Sheets("ER").Select
n = Worksheets("Parametros").Cells(9, 3).Value
 Range("D55").Select
    ActiveCell.FormulaR1C1 = "=+R[-41]C"
    Range("E55").Select
    ActiveCell.FormulaR1C1 = _
        "=+R[-41]C*IF(Parametros!R15C3=""MX"",(1+VLOOKUP(ER!R[-1]C,Tabla5,2,0))^(-ER!R[-1]C),IF(Parametros!R15C3=""US"",(1+VLOOKUP(ER!R[-1]C,Tabla5,3,0))^(-ER!R[-1]C),(1+VLOOKUP(ER!R[-1]C,Tabla5,4,0)^(-ER!R[-1]C))))"
    Range("E55").Select
    Selection.AutoFill Destination:=Range(Cells(55, 5), Cells(55, 5 + (n - 2))), Type:=xlFillDefault
    Cells(55, 4 + n).Select
    For i = 0 To n - 1
    Cells(200, 4 + i).Select
    ActiveCell.FormulaR1C1 = "=+R[-145]C"
    Cells(201, 4 + i).Select
    ActiveCell.FormulaR1C1 = "=+R[-145]C"
    Next i
    Range("D203") = WorksheetFunction.Sum(Range(Cells(200, 4), Cells(200, 4 + (n - 1))))
    Cells(55, 4 + (n - 1)).Select
    Range("D56").Select
    ActiveCell.FormulaR1C1 = "=+R[-14]C"
    Range("E56").Select
    ActiveCell.FormulaR1C1 = _
        "=+R[-14]C*IF(Parametros!R15C3=""MX"",(1+VLOOKUP(ER!R[-2]C,Tabla5,2,0))^(-ER!R[-2]C),IF(Parametros!R15C3=""US"",(1+VLOOKUP(ER!R[-2]C,Tabla5,3,0))^(-ER!R[-2]C),(1+VLOOKUP(ER!R[-2]C,Tabla5,4,0)^(-ER!R[-2]C))))"
    Range("E56").Select
    Selection.AutoFill Destination:=Range(Cells(56, 5), Cells(56, 5 + (n - 2))), Type:=xlFillDefault
    Range("D204") = WorksheetFunction.Sum(Range(Cells(201, 4), Cells(201, 4 + (n - 1))))
    Range("D203").Select
    Selection.Copy
    Range("N55").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D204").Select
    Selection.Copy
    Range("N56").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
