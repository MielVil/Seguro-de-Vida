Attribute VB_Name = "Module7"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "=+[@MRI]*VLOOKUP([@A–os],Tabla6,2,0)"
    Range("D5").Select
End Sub
