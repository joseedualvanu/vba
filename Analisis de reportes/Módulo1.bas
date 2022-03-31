Attribute VB_Name = "Módulo1"

Sub Formato_Tabla1()

Worksheets("HojaA").Activate
Worksheets("HojaA").ListObjects.Add(xlSrcRange, Range("$A$1:$S$" & Cells(Rows.Count, 1).End(xlUp).Row), , xlYes).Name = "Table1"
Worksheets("HojaA").ListObjects("Table1").TableStyle = "TableStyleMedium20"

End Sub

Sub Formato_Tabla2()

Worksheets("HojaB").Activate
Worksheets("HojaB").ListObjects.Add(xlSrcRange, Range("$A$1:$V$" & Cells(Rows.Count, 1).End(xlUp).Row), , xlYes).Name = "Table2"
Worksheets("HojaB").ListObjects("Table2").TableStyle = "TableStyleMedium21"

End Sub

Sub Formato_Tabla3()

Worksheets("HojaC").Activate
Worksheets("HojaC").ListObjects.Add(xlSrcRange, Range("$A$1:$C$" & Cells(Rows.Count, 1).End(xlUp).Row), , xlYes).Name = "Table3"
Worksheets("HojaC").ListObjects("Table3").TableStyle = "TableStyleMedium19"

End Sub

