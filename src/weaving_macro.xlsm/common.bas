Attribute VB_Name = "common"
Option Explicit

' (row, clm)のセルの値を読み込んで返す。セルの値がない時は、initをセルに入れて返す
Public Function readCellValue(row As Integer, clm As Integer, init As Integer)
    If Cells(row, clm) = "" Then
        Cells(row, clm) = init
    End If
    readCellValue = Cells(row, clm)
End Function

