Attribute VB_Name = "FromDrawDown"
'-----------------------------------------------------------------
' 組織図から完全意匠図を作成
'   by Riko(https://github.com/riko122/WeavingMacro)
' This software is released under the Mozilla Public License 2.0.
'-----------------------------------------------------------------
Option Explicit

Const HEADER_LINE = 7 ' ヘッダー部分の行数

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
Private Sub init()
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
            y0 = HEADER_LINE + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "右下"
            x0 = 1
            x1 = x0 + w - 1
            x2 = x1 + 2
            x3 = x2 + f - 1
            y2 = HEADER_LINE + 2
            y3 = y2 + h - 1
            y0 = y3 + 2
            y1 = y0 + n - 1
        Case "左上"
            x2 = 1
            x3 = x2 + f - 1
            x0 = x3 + 2
            x1 = x0 + w - 1
            y0 = HEADER_LINE + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "左下"
            x2 = 1
            x3 = x2 + f - 1
            x0 = x3 + 2
            x1 = x0 + w - 1
            y2 = HEADER_LINE + 2
            y3 = y2 + h - 1
            y0 = y3 + 2
            y1 = y0 + n - 1
    End Select
End Sub

Public Sub makeCanvas()
    Call init
    
    ' クリア。ヘッダー以外の行をちょっと多めに削除する。
    Rows(HEADER_LINE + 1 & ":" & HEADER_LINE + n + h + 100).Select
    Selection.Delete Shift:=xlUp

    ' 対象範囲のマスの高さをそろえる。
    Rows(HEADER_LINE + 1 & ":" & HEADER_LINE + n + h + 5).Select
    Selection.RowHeight = 11
    
    Call writeGrid(y0, y1, x0, x1) ' 綜絖通し部分のマス目
    Call writeGrid(y0, y1, x2, x3) ' タイアップ部分のマス目
    Call writeGrid(y2, y3, x0, x1) ' 組織図部分のマス目
    Call writeGrid(y2, y3, x2, x3) ' 踏み木部分のマス目
End Sub

Public Sub make()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim a As Integer
    Dim s As Integer
    Dim status() As String
    Dim found As Boolean
    Dim first_row As Integer
    Dim last_row As Integer
    Dim shaft_row() As Integer
    
    Call init
    ' 踏み木を踏んだら、綜絖が上がるか下がるかを読み取る。
    kind = Cells(6, 46)  ' ↑か↓
    
    ReDim status(n)
    ReDim shaft_row(n)
    
    ' 綜絖の通し方図と踏み方図の範囲をクリア
    Range(Cells(y0, x0), Cells(y1, x1)).Interior.ColorIndex = xlNone
    Range(Cells(y2, x2), Cells(y3, x3)).Interior.ColorIndex = xlNone
    
    first_row = getFirstRow(y2, y3, x0, x1)
    If first_row = 0 Then
        MsgBox ("組織図が黒く塗られていません")
        Exit Sub
    End If
    last_row = getLastRow(y2, y3, x0, x1)
    
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
        If a > 0 Then
            For j = 0 To a - 1
                If status(j) = status(a) Then
                    Cells(y0 + j, i).Interior.ColorIndex = 1
                    found = True
                    ' aは再利用
                    Exit For
                End If
            Next j
        End If
        ' 見つからなかった場合(a=0の時も)は、新しい行なのでy0+aを黒くする
        If found = False Then
            If a >= n Then
                MsgBox ("この組織図を実現するには綜絖が足りません")
                Exit Sub
            End If
            Cells(y0 + a, i).Interior.ColorIndex = 1
            a = a + 1 ' aは次を使う
        End If
Continue:
    Next i
    
    If getFirstRow(y0, y1, x2, x3) = 0 Then
        MsgBox ("現在のところタイアップが書かれているものにしか対応していません")
        Exit Sub
    End If
    
    ' 踏み木を考える
    If getMaxShaftPerPedal = 1 Then
        For i = y0 To y1
            For j = x1 To x0 Step -1
                ' 綜絖の通し方のi行目で最初に出てくる黒い列を探す
                If Cells(i, j).Interior.ColorIndex = 1 Then
                    ' Tie-upでその行が黒い列を探す
                    found = False
                    For k = x2 To x3
                        If Cells(i, k).Interior.ColorIndex = 1 Then
                            Call copyDrawDownToTreadling(first_row, last_row, j, k)
                            found = True
                            Exit For
                        End If
                    Next k
                    If found Then
                        Exit For ' その綜絖に該当する踏み方は書いたので終わる
                    Else ' 全部見ても見つからなかった場合
                        MsgBox ("この組織図を実現するにはタイアップが不適切です")
                        Exit Sub
                    End If
                End If
            Next j
        Next i
    Else
        For k = first_row To last_row
            ' 各行について、綜絖何枚めが黒いか（経糸が上か）を読み取る
            ' 全列読まなくても、綜絖の枚数分でいい（あとは同じパターンだから）
            a = 0
            For i = y0 To y1
                For j = x1 To x0 Step -1
                    ' 綜絖の通し方のi行目で最初に出てくる黒い列を探す
                    If Cells(i, j).Interior.ColorIndex = 1 Then
                        shaft_row(a) = getTieupStatus(k, j)
                        a = a + 1
                        Exit For
                    End If
                Next j
            Next i
            ' タイアップで、その行が黒い列を探す
            For j = x2 To x3
                found = True
                For a = 0 To n - 1
                    If Cells(y0 + a, j).Interior.ColorIndex <> shaft_row(a) Then
                        found = False
                        Exit For ' 違ったので次の列を探す
                    End If
                Next a
                ' 違わないままFor aが終わった場合は、jが該当する踏み木
                If found Then
                    Cells(k, j).Interior.ColorIndex = 1
                    Exit For
                End If
            Next j
            If found = False Then ' タイアップ全部見ても見つからない場合
                MsgBox ("この組織図を実現するにはタイアップが不適切です")
                Exit Sub
            End If
        Next k
    End If
End Sub

' 綜絖が通っているclm列が、組織図のrow行目で黒いか白いかで、TieUpの状態を示す
' 例えば4枚綜絖の1と4が黒い行は、↑なら1と4が黒いタイアップの列、
' ↓なら2と3が黒いタイアップの列を探すので、↑か↓かによって返すものが逆。
Private Function getTieupStatus(row As Integer, clm As Integer) As Integer

    If kind = "↑" Then ' 天秤式など。組織図で黒ければタイアップで黒い
        If Cells(row, clm).Interior.ColorIndex = 1 Then
            getTieupStatus = 1
        Else
            getTieupStatus = xlNone
        End If
    Else ' ろくろ式など。組織図で白ければタイアップで黒い
        If Cells(row, clm).Interior.ColorIndex <> 1 Then
            getTieupStatus = 1
        Else
            getTieupStatus = xlNone
        End If
    End If
End Function

' 組織図のfrom_clm列の状態をもとに、踏み方図のto_clm列の状態を決める
' ↑の場合はそのままコピー、↓の場合は白黒反転コピー
Private Sub copyDrawDownToTreadling(first_row As Integer, last_row As Integer, _
                                    from_clm As Integer, to_clm As Integer)
    Dim i As Integer
    
    If kind = "↑" Then ' 天秤式など
        Range(Cells(first_row, from_clm), Cells(last_row, from_clm)).Copy Cells(first_row, to_clm)
    Else ' ろくろ式など
        For i = first_row To last_row
            If Cells(i, from_clm).Interior.ColorIndex <> 1 Then
                Cells(i, to_clm).Interior.ColorIndex = 1
            End If
        Next i
    End If
    
End Sub

Private Function getCurrentColumnStatus(clm As Integer) As String
    Dim i As Integer
    Dim status As String
    
    status = ""
    For i = y2 To y3
        If (Cells(i, clm).Interior.ColorIndex = 1) Then
            status = status + "1"
        Else
            status = status + "0"
        End If
    Next i
    getCurrentColumnStatus = status
End Function

' 一本の踏み木につながっている綜絖枠の最大数
Private Function getMaxShaftPerPedal()
    Dim i As Integer
    Dim j As Integer
    Dim cnt As Integer
    Dim max As Integer
    
    max = 0
    For i = x2 To x3
        cnt = 0
        For j = y0 To y1
            If (Cells(j, i).Interior.ColorIndex = 1) Then
                cnt = cnt + 1
            End If
        Next j
        If cnt > max Then
            max = cnt
        End If
    Next i
    getMaxShaftPerPedal = max
End Function

