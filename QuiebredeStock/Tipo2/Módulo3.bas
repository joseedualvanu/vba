Attribute VB_Name = "Módulo3"
Sub eActualizarTabla()
Attribute eActualizarTabla.VB_ProcData.VB_Invoke_Func = "g\n14"
'
' eActualizarTabla Macro
'
' Acceso directo: CTRL+g
'
    Sheets("Tabla dinamica").Select
    ActiveSheet.PivotTables("TablaDinámica").PivotCache.Refresh
End Sub

Sub fBorrarHojas()
Attribute fBorrarHojas.VB_ProcData.VB_Invoke_Func = "h\n14"
'
' fBorrarHojas Macro
'
    Application.DisplayAlerts = False
    Sheets("Hoja").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Filtros").Select
    ActiveWindow.SelectedSheets.Delete
    Range("B15").Select
    Application.DisplayAlerts = True
    Sheets("HojaComentarios").Select
    
End Sub

