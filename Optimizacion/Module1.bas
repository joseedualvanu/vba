Attribute VB_Name = "Module1"
Option Explicit

'Declaraciones

Dim PathDIR As String
Dim PathListaDemanda As String
Dim PathSolver As String
Dim PathTabDistro As String
Dim PathImportTATA1 As String
Dim PathValidacion As String
Dim LastRow As Long
Dim LastRowDistro As Long
Dim LastRowRes As Long
Dim closedBook
Dim Sheet As Worksheet
Dim i As Integer
Dim Estadistico As Long
Dim SolverBook As Workbook
Dim DistroBook As Workbook
Dim r As Range ' screenshot
Dim oCht As Chart 'screenshot
Dim Timestamp As String
Dim wb As Workbook
Dim MacroBook As Workbook
Dim miArray() As Variant
Dim contador As Integer
Dim resultado As Workbook

Sub run()
Call Initialize
Call CopiarListaDemanda
Call ProcesoLD
Call Solver
Call Terminado
Call clean
End Sub

Private Sub Initialize()
Set MacroBook = ActiveWorkbook
If Range("C6").Value = "X" And Range("C12").Value = "" Then
    PathDIR = Sheets("Home").Range("B7").Value
    PathListaDemanda = PathDIR + Sheets("Home").Range("B8").Value
    PathSolver = PathDIR + Sheets("Home").Range("B9").Value
    PathTabDistro = PathDIR + Sheets("Home").Range("B10").Value
    PathImportTATA1 = PathDIR + Sheets("Home").Range("B11").Value
    
ElseIf Range("C6").Value = "" And Range("C12").Value = "X" Then
    PathDIR = Sheets("Home").Range("B14").Value
    PathListaDemanda = PathDIR + Sheets("Home").Range("B15").Value
    PathSolver = PathDIR + Sheets("Home").Range("B16").Value
    PathTabDistro = PathDIR + Sheets("Home").Range("B17").Value
    PathImportTATA1 = PathDIR + Sheets("Home").Range("B18").Value
    PathValidacion = PathDIR + Sheets("Home").Range("B19").Value´
Else
    MacroBook.Activate
    Sheets("Home").Select
    Range("B21").Value = "Path no seteados correctamente, revisar y volver a correr"
    'MsgBox ("Path no seteados correctamente, revisar y volver a correr"), , "Macro detenida"
End
End If

Application.DisplayAlerts = False 'Apagar ventanitas de Excel porque es un gorra
End Sub

Sub CopiarListaDemanda()
Application.ScreenUpdating = False

    'Borra pestañas preexistentes
    
    For Each Sheet In ActiveWorkbook.Worksheets
         If Sheet.Name = "ListaDemanda" Then
         Sheets("ListaDemanda").Delete
         Else
         End If
    Next Sheet

    Set closedBook = Workbooks.Open(PathListaDemanda)
    closedBook.Sheets("lisbasefrescos").Copy Before:=ThisWorkbook.Sheets(1)
    closedBook.Close SaveChanges:=False

Application.ScreenUpdating = True
Sheets("lisbasefrescos").Name = "ListaDemanda"
End Sub

Sub ProcesoLD()

'elimino la primera y segunda row
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
'separo en columnas
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1)), _
        TrailingMinusNumbers:=True
        
'elimino columnas innecesarias
    Range("V:V,U:U,T:T,S:S,R:R,P:P,O:O,N:N,M:M,K:K,J:J,I:I,H:H,G:G,F:F,E:E,D:D,C:C" _
    ).Select
    Range("C1").Activate
    Selection.Delete Shift:=xlToLeft
'defino ultima row
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   Range("A1").Select
' No hay datos?
If LastRow = 1 Then
'MsgBox ("Lista demanda está vacía, no se puede continuar"), , "Error"
    MacroBook.Activate
    Sheets("Home").Select
    Range("B21").Value = "Lista demanda está vacía, no se puede continuar"
End
End If
'filtro por estado =! 0
    ActiveSheet.Range("$A$1:$D$" & LastRow).AutoFilter Field:=4, Criteria1:="<>0", _
    Operator:=xlAnd
'Elimino lo filtrado
    ActiveSheet.Range("$A$1:$D$" & LastRow).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.ShowAllData
'copio
 LastRow = Cells(Rows.Count, 1).End(xlUp).Row
 Range("A2:C" & LastRow).Copy
 
End Sub

Sub Solver()
Dim Result As OpenSolverResult

ReDim miArray(60, 25)

Application.Wait (Now + TimeValue("0:00:2"))
Workbooks.Open Filename:=PathSolver
Set SolverBook = ActiveWorkbook
Application.WindowState = xlMaximized 'maximize Excel
ActiveWindow.WindowState = xlMaximized 'maximize the workbook in Excel
Worksheets("Demanda").Activate
Range("A2").Select
ActiveSheet.Paste
Sheets("Solver").Select
ActiveWorkbook.RefreshAll
    Workbooks.Open Filename:=PathTabDistro
    Set DistroBook = ActiveWorkbook
    Cells.Replace What:="n", Replacement:="250", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'Defino ultimafila en listadistro
    LastRowDistro = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A1").Select
'Empiezo el loop

For i = 2 To LastRowDistro
    On Error GoTo -1
    DistroBook.Activate
    Estadistico = Range("C" & i).Value
    'Voy a celda B3 (solver) e inserto el primer valor (estadistico) de la columna C de la Tabla de distribución
    SolverBook.Activate
        Sheets("Solver").Select
       ' On Error Resume Next
        On Error GoTo NoEstaEstadistico
        'Numero de estadistico
        miArray(i, 0) = Estadistico
        
        ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("ESTADISTICO"). _
        ClearAllFilters
        ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("ESTADISTICO"). _
        CurrentPage = Estadistico ' Agarrar error si no llega a encontrar el estadistico
        On Error GoTo 0
    'correr solver
        Application.Wait (Now + TimeValue("0:00:3"))
    
        'Reseteo los unos en la planilla
        Worksheets("Solver").Activate
        Range("D6").Select
        ActiveCell.FormulaR1C1 = "0"
        Range("D6").Select
        Selection.AutoFill Destination:=Range("D6:D50")
        Range("D6:D95").Select
        
        Range("E6").Select
        ActiveCell.FormulaR1C1 = "0"
        Range("E6").Select
        Selection.AutoFill Destination:=Range("E6:G6"), Type:=xlFillDefault
        Range("E6:G6").Select
        Selection.AutoFill Destination:=Range("E6:G50"), Type:=xlFillDefault
        Range("E6:G95").Select
        
        'OPCION 1: correr solver con todas las definiciones
        'SolverReset
        'Precision y timeout
        'SolverOptions precision:=0.001, MaxTime:=180
        'Funcion Objetivo
        'SetCell valor objetivo, MaxMinVal = 2 minimiza, ValueOf valor para minimizar, ByChange celdas con variables
        'Luego defino el solver
        'SolverOk SetCell:="$T$10", _
        'MaxMinVal:=2, _
        'ByChange:="$D$6:$G$50", _
        'Engine:=3, EngineDesc:="Evolutionary"
        'Restriccion 1: variables binarias
        'SolverAdd cellRef:=Range("$D$6:$G$50"), _
         relation:=5 'binaria
        'Restriccion 2: 'todos los locales tienen un folio
        'SolverAdd cellRef:=Range("$H$6:$H$50"), _
        ' relation:=1, _
        ' formulaText:=1 '<=
        'Restriccion 3: cap de bultos max 1, diferencia mayor o igual a 0
        'SolverAdd cellRef:=Range("$L$15"), _
        ' relation:=3, _
        ' formulaText:=0 ' >=
        'Restriccion 4: cap de bultos max 2, diferencia mayor o igual a 0
        'SolverAdd cellRef:=Range("$L$16"), _
        ' relation:=3, _
        ' formulaText:=0 '>=
        'Restriccion 5: cap de bultos max 3, diferencia mayor o igual a 0
        'SolverAdd cellRef:=Range("$L$17"), _
        ' relation:=3, _
         formulaText:=0 '>=
        'Restriccion 6: cap de bultos max 4, diferencia mayor o igual a 0
        'SolverAdd cellRef:=Range("$L$18"), _
        ' relation:=3, _
        ' formulaText:=0 '>=
        'Restriccion 7: valor objetivo mayor o igual a cero
        ' SolverAdd cellRef:=Range("$T$10"), _
        'relation:=3, _
        ' formulaText:=0  '>=
        'Resuelvo y no muestro el cartel
        'x = SolverSolve(True, True)
        'Guardo el solver
        'SolverSave SaveArea:=Range("L25")
        
        'OPCION 2: Ejecuto el solver1 guardado en ese lugar
        'SolverLoad loadArea:=Range("N13:N23")
        'x = SolverSolve(True, True)
        'Application.Wait (Now + TimeValue("0:00:3"))
        'x = SolverSolve(True, True)
        
        'OPCION 3: Ejecuto el solver2 guardado en ese lugar
        'SolverLoad loadArea:=Range("P13:P23")
        'x = SolverSolve(True, True)
        
        'OPCION 4: opensolver
        'Si el folio 1 tiene el 100% de la demanda y puede cumplir con toda la demanda, le asigno todo los folios
        If Range("L2").Value = "100" And Range("M2").Value > Range("M5").Value Then
        Worksheets("Solver").Activate
        Range("D6").Select
        ActiveCell.FormulaR1C1 = "1"
        Range("D6").Select
        Selection.AutoFill Destination:=Range("D6:D50")
        Range("D6:D95").Select
        
        Else
        Result = RunOpenSolver(False, True) ' do not relax IP, do hide dialogs
        
        End If

        'Prueba con relax IP
        'Result = RunOpenSolver(True, True) ' do not relax IP, do hide dialogs
      
    
    Application.Wait (Now + TimeValue("0:00:3"))
    
    contador = i
    'Obtengo los datos
    'Numero de estadistico
    miArray(i, 0) = Estadistico
    'Folio principal
    miArray(i, 1) = Range("J2")
    
    'Folio 1
    'Numero de folio1
    miArray(i, 2) = Range("K2")
    'Porcentaje del pedido1
    miArray(i, 3) = Range("L2")
    'Porcentaje obtenido
    miArray(i, 4) = Range("M9")
    'Bultos Max
    miArray(i, 5) = Range("M2")
    'Suma bultos1: resultado solver
    miArray(i, 6) = Range("M8")
    
    'Folio2
    'Numero de folio2
    miArray(i, 7) = Range("N2")
    'Porcentaje del pedido1
    miArray(i, 8) = Range("O2")
    'Porcentaje obtenido
    miArray(i, 9) = Range("O9")
    'Bultos Max
    miArray(i, 10) = Range("P2")
    'Suma bultos1: resultado solver
    miArray(i, 11) = Range("O8")
    
    'Folio3
    'Numero de folio3
    miArray(i, 12) = Range("Q2")
    'Porcentaje del pedido1
    miArray(i, 13) = Range("R2")
    'Porcentaje obtenido
    miArray(i, 14) = Range("Q9")
    'Bultos Max
    miArray(i, 15) = Range("S2")
    'Suma bultos1: resultado solver
    miArray(i, 16) = Range("Q8")
    
    'Folio4
    'Numero de folio4
    miArray(i, 17) = Range("T2")
    'Porcentaje del pedido1
    miArray(i, 18) = Range("U2")
    'Porcentaje obtenido
    miArray(i, 19) = Range("S9")
    'Bultos Max
    miArray(i, 20) = Range("V2")
    'Suma bultos1: resultado solver
    miArray(contador, 21) = Range("S8")
    
    'Suma total bultos
    miArray(i, 22) = Range("M5")
    'Celda objetivo
    miArray(i, 23) = Range("T10")
    'Salida del solver
    miArray(i, 24) = Result
            
Nohayscreen:
Application.DisplayAlerts = False
'oCht.Delete
Application.DisplayAlerts = True
    'copíar hoja asignacion a resultado
         Application.CutCopyMode = False
         Sheets("Asignacion").Select
         Range("S2:V46").Copy
         Sheets("Resultado").Select
         LastRowRes = Cells(Rows.Count, 1).End(xlUp).Row
         LastRowRes = LastRowRes + 1
         Range("A" & LastRowRes).PasteSpecial Paste:=xlPasteValues

NoEstaEstadistico:
'MsgBox ("No encuentra estadistico" & Estadistico), , "No encontrado"
Next i
    
End Sub
Sub Terminado()
'defino ultima row
   Sheets("Resultado").Select
   Application.DisplayAlerts = False
   'PathImportTATA1 = "\\srv-cnd-fs01.tata.com.uy\Publicas\Sistemas\RPA\Sync\1-BOT-Sync\DFV-TATA1-FruVer\Output\Lista Importacion Tata1.xlsx"
   Set SolverBook = ActiveWorkbook
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   If LastRow = 1 Then
    'MsgBox ("Estadisticos no coinciden, validar listas"), , "Error"
    MacroBook.Activate
    Sheets("Home").Select
    Range("B21").Value = "Estadisticos no coinciden, validar listas"
    End
    End If
   Range("A1").Select
'filtro por estado = borrar
    ActiveSheet.Range("$A$1:$D$" & LastRow).AutoFilter Field:=4, Criteria1:="Borrar*", _
    Operator:=xlAnd
'Elimino lo filtrado
    ActiveSheet.Range("$A$1:$D$" & LastRow).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.ShowAllData
    
    Set wb = Workbooks.Add
    SolverBook.Sheets("Resultado").Copy Before:=wb.Sheets(1)
    wb.Activate
    
'Ordenar la lista de output para mejorar performance en tata1
    With ActiveSheet.Sort
         .SortFields.Add Key:=Range("B1"), Order:=xlAscending
         .SetRange Range("$A$1:$D$" & Cells(Rows.Count, 1).End(xlUp).Row)
         .Header = xlYes
         .Apply
    End With

    wb.SaveAs PathImportTATA1
    wb.Close SaveChanges:=False
    SolverBook.Close SaveChanges:=False
    DistroBook.Close SaveChanges:=False
    
    'Guardo la validación
    Set resultado = Workbooks.Add
    resultado.Activate
    Range("A1:Y60") = miArray
     
    'Encabezados
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Estadistico"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Folio ppal"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Folio1"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "% ped1"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "% obtenido1"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Bultos max1"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "Suma bultos1"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "Folio2"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "% ped2"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "% obtenido2"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "Bultos max2"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "Suma bultos2"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "Folio3"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "% ped3"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "% obtenido3"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "Bultos max3"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "Suma bultos3"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "Folio4"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "% ped4"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "% obtenido4"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "Bultos max4"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "Suma bultos4"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "Suma Bultos Total"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "Celda Objetivo"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "Resultado Solver"
    Range("G3").Select
    
    'Formato Tabla
    Worksheets("Hoja1").Activate
    Worksheets("Hoja1").ListObjects.Add(xlSrcRange, Range("$A$2:$Y$" & Cells(Rows.Count, 1).End(xlUp).Row), , xlYes).Name = "Table1"
    Worksheets("Hoja1").ListObjects("Table1").TableStyle = "TableStyleMedium20"

    resultado.SaveAs PathValidacion
    resultado.Close SaveChanges:=True
    'SolverBook.Close SaveChanges:=False
    
'MsgBox ("Proceso terminado"), , "Fin"
End Sub

'Funcion que copia formulas + copia y pega valores +en cada fila con datos
Private Sub AplicarFormulaAlaColumnaSoloTexto(Formula As String, FromRow As Integer, FormulaColumn As Integer, LastRowColumn As Integer, Optional xlPasteValus As Boolean = True)
    Range(Cells(FromRow, FormulaColumn), Cells(Cells(Rows.Count, LastRowColumn).End(xlUp).Row, FormulaColumn)).Formula = Formula
    If xlPasteValus Then
    Range(Cells(FromRow, FormulaColumn), Cells(Cells(Rows.Count, LastRowColumn).End(xlUp).Row, FormulaColumn)).Copy
    Cells(FromRow, FormulaColumn).PasteSpecial xlPasteValues
    End If
    
End Sub

'Copia la pestaña de excepciones desde el archivo previamente descargado
Sub CopiarExcepcionesSinAbrirFile()
Application.ScreenUpdating = False

    Set closedBook = Workbooks.Open(PathExcepciones)
    closedBook.Sheets("Hoja1").Copy Before:=ThisWorkbook.Sheets(1)
    closedBook.Close SaveChanges:=False

Application.ScreenUpdating = True
Sheets("Hoja1").Name = "Excepciones"
End Sub



'Funcion que copia Formulas + copia y pega solamente los valores, elimina la formula de la celda + en cada fila con datos
Private Sub AplicarFormulaAlaColumna(Formula As String, FromRow As Integer, FormulaColumn As Integer, LastRowColumn As Integer, Optional xlPasteValus As Boolean = True)
    Range(Cells(FromRow, FormulaColumn), Cells(Cells(Rows.Count, LastRowColumn).End(xlUp).Row, FormulaColumn)).Formula = Formula
End Sub


Sub clean()

For Each Sheet In ActiveWorkbook.Worksheets
     If Sheet.Name = "ListaDemanda" Then
     Sheets("ListaDemanda").Delete
     Else
     End If
Next Sheet



End Sub