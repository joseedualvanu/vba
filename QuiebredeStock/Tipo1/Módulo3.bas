Attribute VB_Name = "Módulo3"

Dim PathStockGral As String
Dim closedBook
Dim Sheet As Worksheet

'Copia la pestaña de Stock General desde el archivo previamente descargado
Sub aaCopiar()
Application.DisplayAlerts = False 'Apagar ventanitas de Excel

'Borro si existe una Hoja "Hoja"
For Each Sheet In ActiveWorkbook.Worksheets
     If Sheet.Name = "Hoja" Then
     Sheets("Hoja").Delete
     Else
     End If
Next Sheet

PathStockGral = "direccion del archivo a copiar"
Application.ScreenUpdating = False

    Set closedBook = Workbooks.Open(PathStockGral)
    closedBook.Sheets("lisloc01-ncr").Copy Before:=ThisWorkbook.Sheets(1)
    closedBook.Close SaveChanges:=False

Application.ScreenUpdating = True
Sheets("lisloc01-ncr").Name = "Hoja"
End Sub





