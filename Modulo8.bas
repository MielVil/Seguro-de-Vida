Attribute VB_Name = "Module8"
Sub mat()
Attribute mat.VB_ProcData.VB_Invoke_Func = " \n14"
'
' mat Macro
'

'
    Range("M23").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=+IF(Parametros!R[-16]C[-10]=""DOT"",VLOOKUP(Parametros!R9C3-1,Table2,11,0)*Parametros!R6C3,0)"
    Range("M24").Select
End Sub
Sub Calculo()
Attribute Calculo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Calculo Macro
'

'
Call PP
Call CV
Call ER

End Sub
