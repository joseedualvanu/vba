Attribute VB_Name = "Módulo5"
Option Explicit

Sub h2SaldoNegativo()
' Comentario Artículos con saldo negativo
    
' Defino las variables
Dim LastRow As Variant
    
    ' Obtengo la ultima fila
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Columna con comentarios
        Range("S2").Select
    ' Formula
        ActiveCell.FormulaR1C1 = "=IF(RC[-14] < 0,""Artículos con saldo negativo"","""")"
        
    ' Selecciono la formula y la aplico hacia abajo
        Range("S2").Select
        Selection.AutoFill Destination:=Range("S2:S" & LastRow), Type:=xlFillDefault
        Range("A1").Select
End Sub

