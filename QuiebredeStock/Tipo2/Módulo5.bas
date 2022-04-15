Attribute VB_Name = "Módulo5"
Sub gSaveAsXlsx()
Dim myPath As String

'myPath = Application.ActiveWorkbook.Path
myPath = ActiveWorkbook.Path

Worksheets(Array("HojaComentarios", "Tabla dinamica")).Copy

ActiveWorkbook.SaveAs Filename:=myPath & "\" & "QuiebredeStockBAS.xlsx"

ActiveWorkbook.Close SaveChanges:=True

End Sub

Sub Formato_Tabla()

Worksheets("HojaComentarios").Activate
Worksheets("HojaComentarios").ListObjects.Add(xlSrcRange, Range("$A$1:$Q$" & Cells(Rows.Count, 1).End(xlUp).Row), , xlYes).Name = "Table1"
Worksheets("HojaComentarios").ListObjects("Table1").TableStyle = "TableStyleMedium21"

End Sub

