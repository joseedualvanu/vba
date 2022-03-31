Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
    Range("A2:C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'
    Windows("Solver.xlsx").Activate
    Sheets("Solver").Select
    ActiveWorkbook.RefreshAll
    Windows("DFV-TATA1-FruVer-Macro.xlsb").Activate
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'
    Windows("Planilla Distr. FyV Final.xlsx").Activate
    Cells.Replace What:="n", Replacement:="250", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Windows("DFV-TATA1-FruVer-Macro.xlsb").Activate
End Sub
