Attribute VB_Name = "Module3"
Sub ER()
Attribute ER.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ER Macro
'

'

    Sheets("ER").Select
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "=+IF(Parametros!RC[3]=0,Parametros!R[9]C[-1],0)"  'Trae la Prima Inicial de Parametros
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=+IF(Parametros!R[-1]C[3]=0,0,Parametros!R[8]C[-1])"
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "=R6C4*VLOOKUP(R[-2]C,Table2,2,0)" 'Calcula la prima de renovacion
    Range("D6").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C+R[-2]C" 'calcula total de prima
    'Range("E5").Select
    
 'Copiar y pegar formulas
 Dim i As Integer
i = 1
n = Worksheets("Parametros").Cells(9, 3).Value
a = Worksheets("Parametros").Cells(4, 7).Value
Sheets("ER").Select

'Copia y pega las formulas de prima de renovacion
For i = 1 To (n - a - 2)
    Cells(5, 4 + i).Select
    Selection.Copy
    Cells(5, 5 + i).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Next i
    
   'Copia y pega las formulas de Total de Prima
   For i = 1 To (n - a - 1)
      Cells(6, 3 + i).Select
    Selection.Copy
    Cells(6, 4 + i).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Next i
    'PRIMA CEDIDA
Sheets("ER").Select
    Range("D9").Select
    ActiveCell.FormulaR1C1 = "=+R[-3]C*R9C3"
    Range("D9").Select
    Selection.AutoFill Destination:=Range(Cells(9, 4), Cells(9, 4 + (n - a - 1))), Type:=xlFillDefault
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "=+R[-4]C-R[-1]C"
    Range("D10").Select
    Selection.AutoFill Destination:=Range(Cells(10, 4), Cells(10, 4 + (n - a - 1))), Type:=xlFillDefault
    'Total ingresos
    Sheets("ER").Select
    Range("D14").Select
    ActiveCell.FormulaR1C1 = "=+R[-4]C+R[-2]C"
    Range("D14").Select
    Selection.AutoFill Destination:=Range(Cells(14, 4), Cells(14, 4 + (n - a - 1))), Type:=xlFillDefault
    'Siniestros
    Sheets("ER").Select
    Range("D19").Select
    ActiveCell.FormulaR1C1 = "=+VLOOKUP(R[-16]C,Table2,4,0)*Parametros!R6C3"
    Range("D19").Select
    Selection.AutoFill Destination:=Range(Cells(19, 4), Cells(19, 4 + (n - a - 1))), Type:=xlFillDefault
    'Recuperacion de siniestros
    Sheets("ER").Select
    Range("D20").Select
    ActiveCell.FormulaR1C1 = "=+R[-1]C*R20C3"
    Range("D20").Select
    Selection.AutoFill Destination:=Range(Cells(20, 4), Cells(20, 4 + (n - a - 1))), Type:=xlFillDefault
    'Total siniestros
    Sheets("ER").Select
    Range("D21").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=+R[-2]C-R[-1]C"
    Range("D21").Select
    Selection.AutoFill Destination:=Range(Cells(21, 4), Cells(21, 4 + (n - a - 1))), Type:=xlFillDefault
    'Caducidad
    Sheets("ER").Select
    Range("D22").Select
    ActiveCell.FormulaR1C1 = _
        "=+VLOOKUP(R[-19]C,Table2,10,0)*VLOOKUP(R[-19]C,Tabla1,3,0)"
    Range("D22").Select
    Selection.AutoFill Destination:=Range(Cells(22, 4), Cells(22, 4 + (n - a - 1))), Type:=xlFillDefault
    'Maturity
    Sheets("ER").Select
    Range("D23").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D23").Select
    Selection.AutoFill Destination:=Range(Cells(23, 4), Cells(23, 4 + (n - a - 2))), Type:=xlFillDefault
    Cells(23, 4 + (n - a - 1)).Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(Parametros!R[-16]C[-10]=""DOT"",VLOOKUP(Parametros!R9C3-1,Table2,11,0)*Parametros!R6C3,0)"
    'Total DB
    Sheets("ER").Select
    Range("D24").Select
    ActiveCell.FormulaR1C1 = "=+SUM(R[-3]C:R[-1]C)"
    Range("D24").Select
    Selection.AutoFill Destination:=Range(Cells(24, 4), Cells(24, 4 + (n - a - 1))), Type:=xlFillDefault
    'Con in ag
    Sheets("ER").Select
    Application.CutCopyMode = False
    Range("D27").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(Parametros!R[-23]C[3]=0,VLOOKUP(R[-24]C,Tabla3,IF(Parametros!R[-20]C[-1]=""DOT"",2,IF(Parametros!R[-20]C[-1]=""OV"",3,4)))*ER!R[-21]C,0)"
    Range("E27").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E27").Select
    Selection.AutoFill Destination:=Range(Cells(27, 5), Cells(27, 4 + (n - a - 1))), Type:=xlFillDefault
    'Com R A
    Sheets("ER").Select
    Range("D28").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(Parametros!R[-24]C[3]=0,0,VLOOKUP(R[-25]C,Tabla3,IF(Parametros!R[-21]C[-1]=""DOT"",2,IF(Parametros!R[-21]C[-1]=""OV"",3,4)))*ER!R[-22]C)"
    Range("E28").Select
    ActiveCell.FormulaR1C1 = _
        "=+VLOOKUP(R[-25]C,Tabla3,IF(Parametros!R7C3=""DOT"",2,IF(Parametros!R7C3=""OV"",3,4)))*ER!R[-23]C"
    Range("E28").Select
    Selection.AutoFill Destination:=Range(Cells(28, 5), Cells(28, 4 + (n - a - 1))), Type:=xlFillDefault
    'Bono Agente
    Sheets("ER").Select
    Range("D29").Select
    ActiveCell.FormulaR1C1 = "=+IF(Parametros!R[-25]C[3]=0,RC[-1]*R[-23]C,0)"
    Range("E29").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E29").Select
    Selection.AutoFill Destination:=Range(Cells(29, 5), Cells(29, 4 + (n - a - 1))), Type:=xlFillDefault
    'Con in P
    Sheets("ER").Select
    Range("D30").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(Parametros!R[-26]C[3]=0,VLOOKUP(R[-27]C,Tabla4,IF(Parametros!R[-23]C[-1]=""DOT"",2,IF(Parametros!R[-23]C[-1]=""OV"",3,4)))*ER!R[-24]C,0)"
    Range("E30").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E30").Select
    Selection.AutoFill Destination:=Range(Cells(30, 5), Cells(30, 4 + (n - a - 1))), Type:=xlFillDefault
    'Com R P
    Sheets("ER").Select
    Range("D31").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(Parametros!R[-27]C[3]=0,0,VLOOKUP(R[-28]C,Tabla4,IF(Parametros!R[-24]C[-1]=""DOT"",2,IF(Parametros!R[-24]C[-1]=""OV"",3,4)))*ER!R[-25]C)"
    Range("E31").Select
    ActiveCell.FormulaR1C1 = _
        "=+VLOOKUP(R[-28]C,Tabla4,IF(Parametros!R7C3=""DOT"",2,IF(Parametros!R7C3=""OV"",3,4)),0)*ER!R[-25]C"
    Range("E31").Select
    Selection.AutoFill Destination:=Range(Cells(31, 5), Cells(31, 4 + (n - a - 1))), Type:=xlFillDefault
    'Total comisiones
    Sheets("ER").Select
    Range("D32").Select
    ActiveCell.FormulaR1C1 = "=+SUM(R[-5]C:R[-1]C)"
    Range("D32").Select
    Selection.AutoFill Destination:=Range(Cells(32, 4), Cells(32, 4 + (n - a - 1))), Type:=xlFillDefault
    'Gastos
    Sheets("ER").Select
    Range("D35").Select
    ActiveCell.FormulaR1C1 = "=+R[-29]C*R35C3"
    Range("D35").Select
    Selection.AutoFill Destination:=Range(Cells(35, 4), Cells(35, 4 + (n - a - 1))), Type:=xlFillDefault
    Range("D36").Select
    ActiveCell.FormulaR1C1 = "=+R[-30]C*R36C3"
    Range("D36").Select
    Selection.AutoFill Destination:=Range(Cells(36, 4), Cells(36, 4 + (n - a - 1))), Type:=xlFillDefault
    Range("D37").Select
    ActiveCell.FormulaR1C1 = "=+SUM(R[-2]C:R[-1]C)"
    Range("D37").Select
    Selection.AutoFill Destination:=Range(Cells(37, 4), Cells(37, 4 + (n - a - 1))), Type:=xlFillDefault
    'Costo de reaseguro
    Sheets("ER").Select
    Range("D40").Select
    ActiveCell.FormulaR1C1 = "=+R40C3*R[-34]C"
    Range("D40").Select
    Selection.AutoFill Destination:=Range(Cells(40, 4), Cells(40, 4 + (n - a - 1))), Type:=xlFillDefault
    'Total egresos
    Sheets("ER").Select
    Range("D42").Select
    ActiveCell.FormulaR1C1 = "=+R[-18]C+R[-10]C+R[-5]C+R[-2]C"
    Range("D42").Select
    Selection.AutoFill Destination:=Range(Cells(42, 4), Cells(42, 4 + (n - a - 1))), Type:=xlFillDefault
    'Utilidad
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
    Dim s As Integer
    s = 0
    For s = 0 To n - a - 1
    Sheets("avr").Select
    Cells(119, 1 + s).Select
    Selection.Copy
    Sheets("ER").Select
    Cells(44, 4 + s).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Next s
    
    'Calculo de reserva
    Dim k As Integer
    k = 0
Sheets("ER").Select
n = Worksheets("Parametros").Cells(9, 3).Value
 Range("D55").Select
    ActiveCell.FormulaR1C1 = "=+R[-41]C"
    Range("E55").Select
    ActiveCell.FormulaR1C1 = _
        "=+R[-41]C*IF(Parametros!R15C3=""MX"",(1+VLOOKUP(ER!R[-1]C,Tabla5,2,0))^(-ER!R[-1]C),IF(Parametros!R15C3=""US"",(1+VLOOKUP(ER!R[-1]C,Tabla5,3,0))^(-ER!R[-1]C),(1+VLOOKUP(ER!R[-1]C,Tabla5,4,0)^(-ER!R[-1]C))))"
    Range("E55").Select
    Selection.AutoFill Destination:=Range(Cells(55, 5), Cells(55, 5 + (n - a - 2))), Type:=xlFillDefault
    Cells(55, 4 + n).Select
    For k = 0 To n - a - 1
    Cells(200, 4 + k).Select
    ActiveCell.FormulaR1C1 = "=+R[-145]C"
    Cells(201, 4 + k).Select
    ActiveCell.FormulaR1C1 = "=+R[-145]C"
    Next k
    Range("D203") = WorksheetFunction.Sum(Range(Cells(200, 4), Cells(200, 4 + (n - a - 1))))
    Cells(55, 4 + (n - a - 1)).Select
    Range("D56").Select
    ActiveCell.FormulaR1C1 = "=+R[-14]C"
    Range("E56").Select
    ActiveCell.FormulaR1C1 = _
        "=+R[-14]C*IF(Parametros!R15C3=""MX"",(1+VLOOKUP(ER!R[-2]C,Tabla5,2,0))^(-ER!R[-2]C),IF(Parametros!R15C3=""US"",(1+VLOOKUP(ER!R[-2]C,Tabla5,3,0))^(-ER!R[-2]C),(1+VLOOKUP(ER!R[-2]C,Tabla5,4,0)^(-ER!R[-2]C))))"
    Range("E56").Select
    Selection.AutoFill Destination:=Range(Cells(56, 5), Cells(56, 5 + (n - a - 2))), Type:=xlFillDefault
    Range("D204") = WorksheetFunction.Sum(Range(Cells(201, 4), Cells(201, 4 + (n - a - 1))))
    Range("D203").Select
    Selection.Copy
    Cells(55, 4 + n - a).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D204").Select
    Selection.Copy
    Cells(56, 4 + n - a).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    
    
    Sheets("Parametros").Select
End Sub
