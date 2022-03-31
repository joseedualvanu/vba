Attribute VB_Name = "Módulo4"
Sub iSaveAsXlsx()
Dim myPath As String

myPath = ActiveWorkbook.Path

Worksheets(Array("Hoja")).Copy

ActiveWorkbook.SaveAs Filename:=myPath & "\" & "QuiebredeStock.xlsx"

ActiveWorkbook.Close SaveChanges:=True

End Sub

Sub Formato_Tabla()

Worksheets("Hoja").Activate
Worksheets("Hoja").ListObjects.Add(xlSrcRange, Range("$A$1:$S$" & Cells(Rows.Count, 1).End(xlUp).Row), , xlYes).Name = "Table1"
Worksheets("Hoja").ListObjects("Table1").TableStyle = "TableStyleMedium20"

End Sub
