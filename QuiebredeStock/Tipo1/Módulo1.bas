Attribute VB_Name = "Módulo1"

Sub aTratamientoDatos()
Attribute aTratamientoDatos.VB_ProcData.VB_Invoke_Func = "a\n14"
'
'Borro las primeras dos filas 
Rows("1:2").Select
Selection.Delete Shift:=xlUp

'Paso de una columna a varias
'Elijo la primer columna
    Columns("A:A").Select
'Cambio el formato para columnas
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array( _
        25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), _
        Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array( _
        38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), _
        Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array(49, 1), Array(50, 1), Array( _
        51, 1), Array(52, 1), Array(53, 1), Array(54, 1), Array(55, 1), Array(56, 1), Array(57, 1), _
        Array(58, 1), Array(59, 1), Array(60, 1), Array(61, 1), Array(62, 1), Array(63, 1), Array( _
        64, 1), Array(65, 1), Array(66, 1), Array(67, 1), Array(68, 1), Array(69, 1)), _
        DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:= _
        True
        
'
' Columnas Macro
'Elijo las columnas a eliminar
    Range("D1,G1,I1,K1:P1,R1,U1,X1:AA1,AC1:AD1,AG1:AN1,AP1:BJ1,BM1:BP1").Select
'Elimino
    Selection.EntireColumn.Delete
    
'
' Filtro10 Macro

'Creo la columna Filtro10
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Filtro10"
    Range("S2").Select
'Aplico la formula
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]<RC[-6],""elimino"",""dejo"")"
    Range("S2").Select
'Aplico para todas las filas hasta la S62000 - cambiar para que tome ultima fila
    Selection.AutoFill Destination:=Range("S2:S62000")
    Range("S2:S62000").Select

'Ordeno en funcion de stock sucursal creciente para que eliminar con el primer filtro sea mas eficiente
    Range("A1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Hoja").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Hoja").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "E1:E62000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Hoja").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
End Sub

