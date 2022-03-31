Attribute VB_Name = "Módulo2"

Sub bFiltro()
Attribute bFiltro.VB_ProcData.VB_Invoke_Func = "s\n14"
' Filtro
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A1:S62000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("Filtro"), Unique:=False
        
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ShowAllData
    
End Sub
Sub cFiltroSuper()
Attribute cFiltroSuper.VB_ProcData.VB_Invoke_Func = "d\n14"
' Filtro FiltroSuper
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A1:S62000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("FiltroSuper"), Unique:=False
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ShowAllData

End Sub
Sub dFiltro11()
Attribute dFiltro11.VB_ProcData.VB_Invoke_Func = "f\n14"
' Filtro11
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A1:S62000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("Filtro11"), Unique:=False
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ShowAllData

End Sub

Sub eFiltro12()
Attribute eFiltro12.VB_ProcData.VB_Invoke_Func = "g\n14"
' Filtro12
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A1:S62000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("Filtro12"), Unique:=False
        
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ShowAllData

End Sub

Sub fFiltro13()
Attribute fFiltro13.VB_ProcData.VB_Invoke_Func = "h\n14"

' Filtro13
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A1:S62000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("Filtro13"), Unique:=False
        
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ShowAllData

End Sub

Sub gFiltro14()
Attribute gFiltro14.VB_ProcData.VB_Invoke_Func = "j\n14"

' Filtro14
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A1:S62000").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("Filtro14"), Unique:=False
        
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ShowAllData

End Sub

Sub hComentarios()
Attribute hComentarios.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' Comentarios Macro
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Comentarios"
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A1").Select
    
End Sub
