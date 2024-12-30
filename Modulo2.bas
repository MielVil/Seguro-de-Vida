Attribute VB_Name = "Module2"
Sub PP()
Attribute PP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PP Macro
'

'
n = Worksheets("Parametros").Cells(9, 3).Value
a = Worksheets("Parametros").Cells(4, 7).Value
Sheets("PP").Select
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=+R[-1]C[9]"
    Range("D3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 4), Cells(3 + (n - a - 1), 4))
    Range("E3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 5), Cells(3 + (n - a - 1), 5))
    Range("F3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 6), Cells(3 + (n - a - 1), 6))
    Range("G3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 7), Cells(3 + (n - a - 1), 7))
    Range("H3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 8), Cells(3 + (n - a - 1), 8))
    Range("I3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 9), Cells(3 + (n - a - 1), 9))
    Range("J3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 10), Cells(3 + (n - a - 1), 10))
    Range("K3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 11), Cells(3 + (n - a - 1), 11))
    Range("L3").Select
    Selection.AutoFill Destination:=Range(Cells(3, 12), Cells(3 + (n - a - 1), 12))
    Range("C4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 3), Cells(4 + (n - a - 2), 3))
    Sheets("Parametros").Select
End Sub
