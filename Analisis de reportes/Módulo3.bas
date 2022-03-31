Attribute VB_Name = "Módulo3"

Sub aCopiar1()
Dim directory1 As String, directory2 As String, fileName1 As String, fileName2 As String, Sheet As Worksheet, total As Integer

Application.DisplayAlerts = False 'Apagar ventanitas de Excel porque es un gorra
'Application.ScreenUpdating = False

For Each Sheet In ActiveWorkbook.Worksheets
     If Sheet.Name = "HojaA" Then
     Sheets("HojaA").Delete
     Else
     End If
Next Sheet

'Parametros
directory1 = "direccion1"
directory2 = "direccion2"

fileName1 = Dir(directory1 & "ped-entrega-ncr.xls")
fileName2 = Dir(directory2 & "macro1 V2.xlsm")

'Abre FichaHoraria.CSV
Workbooks.Open (directory1 & fileName1)
For Each Sheet In Workbooks(fileName1).Worksheets
    total = Workbooks(fileName2).Worksheets.Count
    Workbooks(fileName1).Worksheets(Sheet.Name).Copy _
    after:=Workbooks(fileName2).Worksheets(total)
Next Sheet

'Cierra el archivo FichaHoraria.CSV
Workbooks(fileName1).Close

'renombra esta hoja porque es variable y la estandariza
ActiveSheet.Name = "HojaA"

End Sub


Sub aCopiar2()
Dim directory1 As String, directory2 As String, fileName1 As String, fileName2 As String, Sheet As Worksheet, total As Integer

Application.DisplayAlerts = False 'Apagar ventanitas de Excel porque es un gorra
'Application.ScreenUpdating = False

For Each Sheet In ActiveWorkbook.Worksheets
     If Sheet.Name = "HojaB" Then
     Sheets("HojaB").Delete
     Else
     End If
Next Sheet

'Parametros
directory1 = "direccion1"
directory2 = "direccion2"

fileName1 = Dir(directory1 & "ListadoFrescos.csv")
fileName2 = Dir(directory2 & "macro1 V2.xlsm")

'Abre FichaHoraria.CSV
Workbooks.Open (directory1 & fileName1)
For Each Sheet In Workbooks(fileName1).Worksheets
    total = Workbooks(fileName2).Worksheets.Count
    Workbooks(fileName1).Worksheets(Sheet.Name).Copy _
    after:=Workbooks(fileName2).Worksheets(total)
Next Sheet

'Cierra el archivo FichaHoraria.CSV
Workbooks(fileName1).Close

'renombra esta hoja porque es variable y la estandariza
ActiveSheet.Name = "HojaB"

End Sub



Sub aCopiar3()
Dim directory1 As String, directory2 As String, fileName1 As String, fileName2 As String, Sheet As Worksheet, total As Integer

Application.DisplayAlerts = False 'Apagar ventanitas de Excel porque es un gorra
'Application.ScreenUpdating = False

For Each Sheet In ActiveWorkbook.Worksheets
     If Sheet.Name = "HojaC" Then
     Sheets("HojaC").Delete
     Else
     End If
Next Sheet

'Parametros
directory1 = "direccion1"
directory2 = "direccion2"

fileName1 = Dir(directory1 & "ListadoSecos.csv")
fileName2 = Dir(directory2 & "macro1 V2.xlsm")

'Abre FichaHoraria.CSV
Workbooks.Open (directory1 & fileName1)
For Each Sheet In Workbooks(fileName1).Worksheets
    total = Workbooks(fileName2).Worksheets.Count
    Workbooks(fileName1).Worksheets(Sheet.Name).Copy _
    after:=Workbooks(fileName2).Worksheets(total)
Next Sheet

'Cierra el archivo FichaHoraria.CSV
Workbooks(fileName1).Close

'renombra esta hoja porque es variable y la estandariza
ActiveSheet.Name = "HojaC"

End Sub











