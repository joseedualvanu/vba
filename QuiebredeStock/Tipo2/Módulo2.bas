Attribute VB_Name = "Módulo2"
Sub bFiltro()
Attribute bFiltro.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' Filtro Macro
'
    Range("A1").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("A1:P80000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("Filtro"), Unique:=False
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("A1").Select
    ActiveSheet.ShowAllData
    
End Sub

Sub cMovimiento()
Attribute cMovimiento.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' Comentarios Macro

'Borro los datos anteriores
    Sheets("HojaComentarios").Select
    Range("P2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
'Elijo datos a pasar
    Sheets("Hoja").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("HojaComentarios").Select
    Range("A2").Select
    ActiveSheet.Paste
'Copio y pego en formato de valor para quedarme con la concatenación
    Columns("T:T").Select
    Selection.Copy
    Range("U1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    
End Sub

Sub dBorrarColumnas()
Attribute dBorrarColumnas.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' BorrarColumnas Macro
'
    Range("Q1,R1,S1,T1").Select
    Selection.EntireColumn.Delete
    Range("A1").Select
End Sub
