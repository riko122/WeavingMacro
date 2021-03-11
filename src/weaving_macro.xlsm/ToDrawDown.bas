Attribute VB_Name = "ToDrawDown"
'-----------------------------------------------------------------
' 綜絖の通し方図・タイアップ・踏み方図から、組織図や配色図を作成
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
Dim x9 As Integer '緯糸色指定列
Dim y9 As Integer '経糸色指定行

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
            x9 = x3 + 2
            y9 = HEADER_LINE + 2
            y0 = y9 + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "右下"
            x0 = 1
            x1 = x0 + w - 1
            x2 = x1 + 2
            x3 = x2 + f - 1
            x9 = x3 + 2
            y2 = HEADER_LINE + 2
            y3 = y2 + h - 1
            y0 = y3 + 2
            y1 = y0 + n - 1
            y9 = y1 + 2
        Case "左上"
            x9 = 1
            x2 = x9 + 2
            x3 = x2 + f - 1
            x0 = x3 + 2
            x1 = x0 + w - 1
            y9 = HEADER_LINE + 2
            y0 = y9 + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "左下"
            x9 = 1
            x2 = x9 + 2
            x3 = x2 + f - 1
            x0 = x3 + 2
            x1 = x0 + w - 1
            y2 = HEADER_LINE + 2
            y3 = y2 + h - 1
            y0 = y3 + 2
            y1 = y0 + n - 1
            y9 = y1 + 2
    End Select
End Sub

' 初期化ボタンクリックで実行。
Public Sub clearToDrawDown()

    Call init

    ' クリア。ヘッダー以外の行をちょっと多めに削除する。
    Rows(HEADER_LINE + 1 & ":" & HEADER_LINE + n + h + 100).Select
    Selection.Delete Shift:=xlUp

    ' 対象範囲のマスの高さをそろえる。
    Rows(HEADER_LINE + 1 & ":" & HEADER_LINE + n + h + 5).Select
    Selection.RowHeight = 11

    Call writeGrid(y0, y1, x0, x1) ' 綜絖通し部分のマス目
    Call writeGrid(y0, y1, x2, x3) ' タイアップ部分のマス目

    ' 経糸色指定行に「経糸の色」と書く
    Range(Cells(y9, x2), Cells(y9, x3)).Select
    If x1 < x2 Then
        ActiveCell.FormulaR1C1 = "経糸の色"
    Else
        ActiveCell.FormulaR1C1 = "経糸の色"
    End If
    With Selection.Font
        .Size = 6
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With

    ' 緯糸色指定列に「緯糸の色」と書く
    Range(Cells(y0, x9), Cells(y1, x9)).Select
    If y1 < y2 Then
        ActiveCell.FormulaR1C1 = "緯糸の色"
    Else
        ActiveCell.FormulaR1C1 = "緯糸の色"
    End If
    With Selection.Font
        .Size = 6
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = xlVertical
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With

    Call writeGrid(y2, y3, x0, x1) ' 組織図部分のマス目
    Call writeGrid(y2, y3, x2, x3) ' 踏み木部分のマス目
End Sub

' 組織図ボタンクリックで実行。
Public Sub black()
    Dim first_clm As Integer
    Dim last_clm As Integer
    Dim first_row As Integer
    Dim last_row As Integer
    Dim i As Integer
    Dim j As Integer
    Dim l As Integer
    Dim k As Integer
    Dim init_row_status() As Boolean
    Dim current_row_status() As Boolean
    
    Call init
    ' 踏み木を踏んだら、綜絖が上がるか下がるかを読み取る。
    kind = Cells(6, 40)  ' ↑か↓
    
    ' 組織図部分を書き直す
    Call writeGrid(y2, y3, x0, x1) ' 組織図部分のマス目
    Range(Cells(y2, x0), Cells(y3, x1)).Interior.ColorIndex = xlNone
    
    ReDim init_row_status(w)
    ReDim current_row_status(w)

    first_clm = getFirstColumn(y0, y1, x0, x1)
    If first_clm = 0 Then
        MsgBox ("綜絖の通し方図が黒く塗られていません")
        Exit Sub
    End If
    last_clm = getLastColumn(y0, y1, x0, x1)
    
    first_row = getFirstRow(y2, y3, x2, x3)
    If first_row = 0 Then
        MsgBox ("踏み方図が黒く塗られていません")
        Exit Sub
    End If
    last_row = getLastRow(y2, y3, x2, x3)
    
    init_row_status = setInitRowStatus()
    
    For l = first_row To last_row
        current_row_status = getCurrentRowStatus(l, init_row_status)
        For i = first_clm To last_clm
            ' 経糸が上のマスは黒く塗る
            If current_row_status(i) = True Then
                Cells(l, i).Interior.ColorIndex = 1
            End If
        Next
    Next
End Sub

' 配色図ボタンクリックで実行。
Public Sub color()
    Dim first_clm As Integer
    Dim last_clm As Integer
    Dim first_row As Integer
    Dim last_row As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim init_row_status() As Boolean
    Dim current_row_status() As Boolean
    Dim before_row_status() As Boolean
    
    Call init
    ' 踏み木を踏んだら、綜絖が上がるか下がるかを読み取る。
    kind = Cells(6, 40)  ' ↑か↓
    
    ' 組織図部分を書き直す(色は配色図は上書きするので特に塗りなおさない）
    Call writeGrid(y2, y3, x0, x1) ' 組織図部分のマス目
    
    ReDim init_row_status(w)
    ReDim current_row_status(w)
    ReDim before_row_status(w)

    first_clm = getFirstColumn(y0, y1, x0, x1)
    If first_clm = 0 Then
        MsgBox ("綜絖の通し方図が黒く塗られていません")
        Exit Sub
    End If
    last_clm = getLastColumn(y0, y1, x0, x1)

    first_row = getFirstRow(y2, y3, x2, x3)
    If first_row = 0 Then
        MsgBox ("踏み方図が黒く塗られていません")
        Exit Sub
    End If
    last_row = getLastRow(y2, y3, x2, x3)
    
    init_row_status = setInitRowStatus()
    For i = x0 To x1
        before_row_status(i) = False
    Next

    For l = first_row To last_row
        current_row_status = getCurrentRowStatus(l, init_row_status)
        For i = first_clm To last_clm
            ' 経糸が上のマスは経糸の色で塗る。そうでなければ緯糸の色で塗る
            If current_row_status(i) = True Then
                Cells(l, i).Interior.color = Cells(y9, i).Interior.color
                ' 前の行も経糸が上なら、上の罫線をなしにする
                If before_row_status(i) = True Then
                    Range(Cells(l, i), Cells(l, i)).Borders(xlEdgeTop).LineStyle = xlNone
                End If
            Else
                Cells(l, i).Interior.color = Cells(l, x9).Interior.color
                ' 左のセルも緯糸が上なら、左の罫線をなしにする
                If i > first_clm And current_row_status(i - 1) = False Then
                    Range(Cells(l, i), Cells(l, i)).Borders(xlEdgeLeft).LineStyle = xlNone
                End If
            End If
        Next
        before_row_status = current_row_status
    Next
End Sub

' 現在の行の各セルについて、経糸が上になっていればTrue, そうでなければFalseを配列に登録する
Private Function getCurrentRowStatus(ByVal row As Integer, ByRef init_row_status() As Boolean) As Boolean()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim current_row_status() As Boolean
    Dim f As Boolean
    
    ' 配列の初期化
    ReDim current_row_status(w)
    current_row_status = init_row_status
   
    ' 踏み方図で黒いマスを探す。
    For j = x2 To x3
        If (Cells(row, j).Interior.ColorIndex = 1) Then
            ' そのマスの列のタイアップで黒のマスを探す
            For k = y0 To y1
                If Cells(k, j).Interior.ColorIndex = 1 Then
                    ' そのマスの行の綜絖の通し方が黒い列の組織点は経糸が上に出るのでTrueにする
                    ' ロクロ式なら、緯糸が上に出るのでFalseにする（そのため初期化で元がTrueにしてある）
                    For i = x0 To x1
                        If Cells(k, i).Interior.ColorIndex = 1 Then
                            If current_row_status(i) = False Then
                                current_row_status(i) = True
                            Else
                                current_row_status(i) = False
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Next
    getCurrentRowStatus = current_row_status
End Function

' 各行の状態初期化。kindに応じて変わる。
Private Function setInitRowStatus() As Boolean()
    Dim init_row_status() As Boolean
    Dim i As Integer
    Dim k As Integer
    
    ReDim init_row_status(w)
    ' 配列の初期化
    For i = x0 To x1
        If kind = "↑" Then ' 天秤式など。false(緯糸が上)で初期化
            init_row_status(i) = False
        Else ' ろくろ式など。基本true(経糸が上)で初期化。但し空羽はfalse。
            init_row_status(i) = False
            For k = y0 To y1
                If Cells(k, i).Interior.ColorIndex = 1 Then
                    init_row_status(i) = True
                    Exit For
                End If
            Next
        End If
    Next
    setInitRowStatus = init_row_status
End Function

