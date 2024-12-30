Attribute VB_Name = "Module5"
Sub Mpf()
Attribute Mpf.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Mpf Macro
'

'
Sheets("ER").Select
n = Worksheets("Parametros").Cells(9, 3).Value
a = Worksheets("Parametros").Cells(4, 7).Value
Range("D50").Select
ActiveCell.FormulaR1C1 = "=+R[-6]C*R49C3"
Range("D50").Select
Selection.AutoFill Destination:=Range(Cells(50, 4), Cells(50, 4 + (n - a - 1))), Type:=xlFillDefault
Range("D12").Select
ActiveCell.FormulaR1C1 = "=+R[38]C"
Range("D12").Select
Selection.AutoFill Destination:=Range(Cells(12, 4), Cells(12, 4 + (n - a - 1))), Type:=xlFillDefault
Sheets("avr").Select
    Dim i As Integer
    i = 0
    For i = 0 To n - a - 1
    Sheets("avr").Select
    Cells(119, 1 + i).Select
    Selection.Copy
    Sheets("ER").Select
    Cells(44, 4 + i).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Next i
    
Sheets("Parametros").Select
End Sub
