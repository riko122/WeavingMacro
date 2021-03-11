Attribute VB_Name = "common"
'-----------------------------------------------------------------
' 共通Function、共通Sub配置用モジュール
'   by Riko(https://github.com/riko122/WeavingMacro)
' This software is released under the Mozilla Public License 2.0.
'-----------------------------------------------------------------
Option Explicit

' (row, clm)のセルの値を読み込んで返す。セルの値がない時は、初期値をセルに入れて返す
Public Function readCellValue(row As Integer, clm As Integer, default As Integer)
    If Cells(row, clm) = "" Then
        Cells(row, clm) = default
    End If
    readCellValue = Cells(row, clm)
End Function

' 指定範囲にマス目を書く
Public Sub writeGrid(first_row As Integer, last_row As Integer, _
                     first_clm As Integer, last_clm As Integer)
    Range(Cells(first_row, first_clm), Cells(last_row, last_clm)).Select
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
End Sub

' 指定範囲内で黒マスがある最初の行を得る
Public Function getFirstRow(first_row As Integer, last_row As Integer, _
                            first_clm As Integer, last_clm As Integer) As Integer
    Dim first As Integer
    Dim k As Integer
    Dim l As Integer
    
    first = 0
    For l = first_row To last_row
        '対象行で黒のマスを探す。黒のマスがあればその行は有効
        For k = first_clm To last_clm
            If Cells(l, k).Interior.ColorIndex = 1 Then
                first = l
                Exit For
            End If
        Next
        If first > 0 Then
            Exit For '有効行があればそこが開始行なので終了
        End If
    Next
    getFirstRow = first
End Function

' 指定範囲内で黒マスがある最後の行を得る
Public Function getLastRow(first_row As Integer, last_row As Integer, _
                           first_clm As Integer, last_clm As Integer) As Integer
    Dim last As Integer
    Dim k As Integer
    Dim l As Integer
    
    last = 0
    For l = last_row To first_row Step -1
        '対象行で黒のマスを探す。黒のマスがあればその行は有効
        For k = first_clm To last_clm
            If Cells(l, k).Interior.ColorIndex = 1 Then
                last = l
                Exit For
            End If
        Next
        If last > 0 Then
            Exit For '有効行があればそこがラスト行なので終了
        End If
    Next
    getLastRow = last
End Function

' 指定範囲内で黒マスがある最初の列を得る
Public Function getFirstColumn(first_row As Integer, last_row As Integer, _
                               first_clm As Integer, last_clm As Integer) As Integer
    Dim first As Integer
    Dim i As Integer
    Dim k As Integer
    
    first = 0
    For i = first_clm To last_clm
        '対象列で黒のマスを探す。黒のマスがあればその列は有効
        For k = first_row To last_row
            If Cells(k, i).Interior.ColorIndex = 1 Then
                first = i
                Exit For
            End If
        Next
        If first > 0 Then
            Exit For
        End If
    Next
    getFirstColumn = first
End Function

' 指定範囲内で黒マスがある最後の列を得る
Public Function getLastColumn(first_row As Integer, last_row As Integer, _
                              first_clm As Integer, last_clm As Integer) As Integer
    Dim last As Integer
    Dim i As Integer
    Dim k As Integer
    
    last = 0
    For i = last_clm To first_clm Step -1
        '綜絖の通し方図対象行で、黒のマスを探す。黒のマスがあればその行は有効
        For k = first_row To last_row
            If Cells(k, i).Interior.ColorIndex = 1 Then
                last = i
                Exit For
            End If
        Next
        If last > 0 Then
            Exit For
        End If
    Next
    getLastColumn = last
End Function
