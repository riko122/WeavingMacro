Attribute VB_Name = "FromDrawDown"
'-----------------------------------------------------------------
' 組織図から完全意匠図を作成
'   by Riko(https://github.com/riko122/WeavingMacro)
' This software is released under the Mozilla Public License 2.0.
'-----------------------------------------------------------------
Option Explicit

Const header_line = 7 ' ヘッダー部分の行数

' 全サブルーチンで共通に使用する変数
Dim x0 As Integer '綜絖の通し方図・組織図の基点列
Dim y0 As Integer '綜絖の通し方図・タイアップの基点行
Dim x1 As Integer '綜絖の通し方図・組織図の最終列
Dim y1 As Integer '綜絖の通し方図・タイアップの最終行
Dim x2 As Integer 'タイアップ・踏み方図の基点列
Dim y2 As Integer '組織図・踏み方図の基点行
Dim x3 As Integer 'タイアップ・踏み方図の最終列
Dim y3 As Integer '組織図・踏み方図の最終行

Dim n As Integer 'n枚綜絖を使用。
Dim f As Integer 'f本の踏み木を使用
Dim w As Integer '組織図の幅
Dim h As Integer '組織図の高さ

Dim kind As String ' 踏み木を踏んだら、綜絖が上がるか下がるか。
Dim tie_up_position As String 'タイアップをどの位置にするか

'初期値設定
Private Sub initFromDrawDown()
    f = readCellValue(7, 5, 4)
    n = readCellValue(7, 14, 4)
    w = readCellValue(7, 36, 48)
    h = readCellValue(7, 46, 48)
    
    tie_up_position = Cells(7, 28)
    Select Case tie_up_position
        Case "右上"
            x0 = 1
            x1 = x0 + w - 1
            x2 = x1 + 2
            x3 = x2 + f - 1
            y0 = header_line + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "右下"
            x0 = 1
            x1 = x0 + w - 1
            x2 = x1 + 2
            x3 = x2 + f - 1
            y2 = header_line + 2
            y3 = y2 + h - 1
            y0 = y3 + 2
            y1 = y0 + n - 1
        Case "左上"
            x2 = 1
            x3 = x2 + f - 1
            x0 = x3 + 2
            x1 = x0 + w - 1
            y0 = header_line + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "左下"
            x2 = 1
            x3 = x2 + f - 1
            x0 = x3 + 2
            x1 = x0 + w - 1
            y2 = header_line + 2
            y3 = y2 + h - 1
            y0 = y3 + 2
            y1 = y0 + n - 1
    End Select
End Sub

Public Sub makeCanvas()
    Call initFromDrawDown
    
    ' クリア。ヘッダー以外の行をちょっと多めに削除する。
    Rows(header_line + 1 & ":" & header_line + n + h + 100).Select
    Selection.Delete Shift:=xlUp

    ' 対象範囲のマスの高さをそろえる。
    Rows(header_line + 1 & ":" & header_line + n + h + 5).Select
    Selection.RowHeight = 11
    
    ' 綜絖通し部分のマス目を書く
    Range(Cells(y0, x0), Cells(y1, x1)).Select
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin

    ' タイアップ部分のマス目を書く
    Range(Cells(y0, x2), Cells(y1, x3)).Select
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin

    ' 組織図部分のマス目を書く
    Range(Cells(y2, x0), Cells(y3, x1)).Select
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin

    ' 踏み木部分のマス目を書く
    Range(Cells(y2, x2), Cells(y3, x3)).Select
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin

End Sub

Public Sub make()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim a As Integer
    Dim s As Integer
    Dim status() As String
    Dim found As Boolean
    Dim firstR As Integer
    Dim lastR As Integer
    
    Call initFromDrawDown
    ' 踏み木を踏んだら、綜絖が上がるか下がるかを読み取る。
    kind = Cells(6, 46)  ' ↑か↓
    
    ReDim status(n)
    
    ' 綜絖の通し方図と踏み方図の範囲をクリア
    Range(Cells(y0, x0), Cells(y1, x1)).Interior.ColorIndex = xlNone
    Range(Cells(y2, x2), Cells(y3, x3)).Interior.ColorIndex = xlNone
    
    firstR = firstRowOnDrawDown()
    If firstR = 0 Then
        MsgBox ("組織図が黒く塗られていません")
        Exit Sub
    End If
    lastR = lastRowOnDrawDown()
    
    ' 綜絖の通し方を考える。
    a = 0
    For i = x1 To x0 Step -1
        ' 現在列のパターンを取得する
        status(a) = getCurrentColumnStatus(i)
        ' 空き羽の場合はどこも黒くしない
        If InStr(status(a), "1") = 0 Then
            GoTo Continue
        End If
            
        ' 綜絖通し方図で黒くする行を決める
        found = False
        ' 今までに同じパターンがあれば、そのパターンと同じ行を黒くする
        For j = 0 To a - 1
            If status(j) = status(a) Then
                Cells(y0 + j, i).Interior.ColorIndex = 1
                found = True
                ' aは再利用
                Exit For
            End If
        Next j
        ' 見つからなかった場合は、新しい行なのでy0+aを黒くする
        If found = False Then
            Cells(y0 + a, i).Interior.ColorIndex = 1
            a = a + 1 ' aは次を使う
        End If
Continue:
    Next i
    
    ' Tie-Upが書かれているかどうか。単式か複式かもチェックしないとな。
    If getTieUpStatus = False Then
        MsgBox ("現在のところ単式でタイアップが描かれていないとダメです")
        Exit Sub
    End If
    
    ' 踏み木を考える
    For i = y0 To y1
        For j = x1 To x0 Step -1
            ' 綜絖の通し方のi行目で最初に出てくる黒い列を探す
            If Cells(i, j).Interior.ColorIndex = 1 Then
                ' Tie-upでその行が黒い列を探す
                For k = x2 To x3
                    If Cells(i, k).Interior.ColorIndex = 1 Then
                        Call copyDrawDownToTreadling(firstR, lastR, j, k)
                        'Range(Cells(y2, j), Cells(y3, j)).Copy Cells(y2, k)
                        Exit For
                    End If
                Next k
                Exit For
            End If
        Next j
    Next i
End Sub

' 組織図のfromClm列の状態をもとに、踏み方図のtoClm列の状態を決める
' ↑の場合はそのままコピー、下の場合は白黒反転コピー
Private Sub copyDrawDownToTreadling(firstRow As Integer, lastRow As Integer, fromClm As Integer, toClm As Integer)
    Dim i As Integer
    
    If kind = "↑" Then ' 天秤式など
        Range(Cells(firstRow, fromClm), Cells(lastRow, fromClm)).Copy Cells(y2, toClm)
    Else ' ろくろ式など
        For i = firstRow To lastRow
            If Cells(i, fromClm).Interior.ColorIndex <> 1 Then
                Cells(i, toClm).Interior.ColorIndex = 1
            End If
        Next i
    End If
    
End Sub

Private Function getCurrentColumnStatus(col As Integer) As String
    Dim i As Integer
    Dim status As String
    
    status = ""
    For i = y2 To y3
        If (Cells(i, col).Interior.ColorIndex = 1) Then
            status = status + "1"
        Else
            status = status + "0"
        End If
    Next i
    getCurrentColumnStatus = status
End Function

Private Function getCurrentRowStatus(row As Integer) As String
    Dim i As Integer
    Dim status As String
    
    status = ""
    For i = x0 To x1
        If (Cells(row, i).Interior.ColorIndex = 1) Then
            status = status + "1"
        Else
            status = status + "0"
        End If
    Next i
    getCurrentRowStatus = status

End Function

' Tie-Upが最初から書かれているかどうか。
' 黒いところが一個でもあれば書かれているとみなす
' 単式か複式かもチェックしたい。
Private Function getTieUpStatus() As Boolean
    Dim i, j As Integer
    getTieUpStatus = False
    
    For i = y0 To y1
        For j = x2 To x3
            If (Cells(i, j).Interior.ColorIndex = 1) Then
                getTieUpStatus = True
                Exit Function
            End If
        Next j
    Next i
End Function

' 組織図に黒マスがある最初の行を得る
Private Function firstRowOnDrawDown() As Integer
    Dim first As Integer
    Dim k As Integer
    Dim l As Integer
    
    first = 0
    For l = y2 To y3
        '組織図対象列で黒のマスを探す。黒のマスがあればその行は有効
        For k = x0 To x1
            If Cells(l, k).Interior.ColorIndex = 1 Then
                first = l
                Exit For
            End If
        Next
        If first > 0 Then
            Exit For '有効行があればそこが開始行なので終了
        End If
    Next
    firstRowOnDrawDown = first
End Function

' 組織図に黒マスがある最後の行を得る
Private Function lastRowOnDrawDown() As Integer
    Dim last As Integer
    Dim k As Integer
    Dim l As Integer
    
    last = 0
    For l = y3 To y2 Step -1
        '踏み方図対象列で黒のマスを探す。黒のマスがあればその行は有効
        For k = x0 To x1
            If Cells(l, k).Interior.ColorIndex = 1 Then
                last = l
                Exit For
            End If
        Next
        If last > 0 Then
            Exit For '有効行があればそこがラスト行なので終了
        End If
    Next
    lastRowOnDrawDown = last
End Function
