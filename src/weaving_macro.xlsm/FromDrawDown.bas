Attribute VB_Name = "FromDrawDown"
'-----------------------------------------------------------------
' �g�D�}���犮�S�ӏ��}���쐬
'   by Riko(https://github.com/riko122/WeavingMacro)
' This software is released under the Mozilla Public License 2.0.
'-----------------------------------------------------------------
Option Explicit

Const header_line = 7 ' �w�b�_�[�����̍s��

' �S�T�u���[�`���ŋ��ʂɎg�p����ϐ�
Dim x0 As Integer '���L�̒ʂ����}�E�g�D�}�̊�_��
Dim y0 As Integer '���L�̒ʂ����}�E�^�C�A�b�v�̊�_�s
Dim x1 As Integer '���L�̒ʂ����}�E�g�D�}�̍ŏI��
Dim y1 As Integer '���L�̒ʂ����}�E�^�C�A�b�v�̍ŏI�s
Dim x2 As Integer '�^�C�A�b�v�E���ݕ��}�̊�_��
Dim y2 As Integer '�g�D�}�E���ݕ��}�̊�_�s
Dim x3 As Integer '�^�C�A�b�v�E���ݕ��}�̍ŏI��
Dim y3 As Integer '�g�D�}�E���ݕ��}�̍ŏI�s

Dim n As Integer 'n�����L���g�p�B
Dim f As Integer 'f�{�̓��ݖ؂��g�p
Dim w As Integer '�g�D�}�̕�
Dim h As Integer '�g�D�}�̍���

Dim kind As String ' ���ݖ؂𓥂񂾂�A���L���オ�邩�����邩�B
Dim tie_up_position As String '�^�C�A�b�v���ǂ̈ʒu�ɂ��邩

'�����l�ݒ�
Private Sub initFromDrawDown()
    f = readCellValue(7, 5, 4)
    n = readCellValue(7, 14, 4)
    w = readCellValue(7, 36, 48)
    h = readCellValue(7, 46, 48)
    
    tie_up_position = Cells(7, 28)
    Select Case tie_up_position
        Case "�E��"
            x0 = 1
            x1 = x0 + w - 1
            x2 = x1 + 2
            x3 = x2 + f - 1
            y0 = header_line + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "�E��"
            x0 = 1
            x1 = x0 + w - 1
            x2 = x1 + 2
            x3 = x2 + f - 1
            y2 = header_line + 2
            y3 = y2 + h - 1
            y0 = y3 + 2
            y1 = y0 + n - 1
        Case "����"
            x2 = 1
            x3 = x2 + f - 1
            x0 = x3 + 2
            x1 = x0 + w - 1
            y0 = header_line + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "����"
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
    
    ' �N���A�B�w�b�_�[�ȊO�̍s��������Ƒ��߂ɍ폜����B
    Rows(header_line + 1 & ":" & header_line + n + h + 100).Select
    Selection.Delete Shift:=xlUp

    ' �Ώ۔͈͂̃}�X�̍��������낦��B
    Rows(header_line + 1 & ":" & header_line + n + h + 5).Select
    Selection.RowHeight = 11
    
    ' ���L�ʂ������̃}�X�ڂ�����
    Range(Cells(y0, x0), Cells(y1, x1)).Select
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin

    ' �^�C�A�b�v�����̃}�X�ڂ�����
    Range(Cells(y0, x2), Cells(y1, x3)).Select
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin

    ' �g�D�}�����̃}�X�ڂ�����
    Range(Cells(y2, x0), Cells(y3, x1)).Select
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).Weight = xlThin

    ' ���ݖؕ����̃}�X�ڂ�����
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
    ' ���ݖ؂𓥂񂾂�A���L���オ�邩�����邩��ǂݎ��B
    kind = Cells(6, 46)  ' ������
    
    ReDim status(n)
    
    ' ���L�̒ʂ����}�Ɠ��ݕ��}�͈̔͂��N���A
    Range(Cells(y0, x0), Cells(y1, x1)).Interior.ColorIndex = xlNone
    Range(Cells(y2, x2), Cells(y3, x3)).Interior.ColorIndex = xlNone
    
    firstR = firstRowOnDrawDown()
    If firstR = 0 Then
        MsgBox ("�g�D�}�������h���Ă��܂���")
        Exit Sub
    End If
    lastR = lastRowOnDrawDown()
    
    ' ���L�̒ʂ������l����B
    a = 0
    For i = x1 To x0 Step -1
        ' ���ݗ�̃p�^�[�����擾����
        status(a) = getCurrentColumnStatus(i)
        ' �󂫉H�̏ꍇ�͂ǂ����������Ȃ�
        If InStr(status(a), "1") = 0 Then
            GoTo Continue
        End If
            
        ' ���L�ʂ����}�ō�������s�����߂�
        found = False
        ' ���܂łɓ����p�^�[��������΁A���̃p�^�[���Ɠ����s����������
        For j = 0 To a - 1
            If status(j) = status(a) Then
                Cells(y0 + j, i).Interior.ColorIndex = 1
                found = True
                ' a�͍ė��p
                Exit For
            End If
        Next j
        ' ������Ȃ������ꍇ�́A�V�����s�Ȃ̂�y0+a����������
        If found = False Then
            Cells(y0 + a, i).Interior.ColorIndex = 1
            a = a + 1 ' a�͎����g��
        End If
Continue:
    Next i
    
    ' Tie-Up��������Ă��邩�ǂ����B�P�������������`�F�b�N���Ȃ��ƂȁB
    If getTieUpStatus = False Then
        MsgBox ("���݂̂Ƃ���P���Ń^�C�A�b�v���`����Ă��Ȃ��ƃ_���ł�")
        Exit Sub
    End If
    
    ' ���ݖ؂��l����
    For i = y0 To y1
        For j = x1 To x0 Step -1
            ' ���L�̒ʂ�����i�s�ڂōŏ��ɏo�Ă��鍕�����T��
            If Cells(i, j).Interior.ColorIndex = 1 Then
                ' Tie-up�ł��̍s���������T��
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

' �g�D�}��fromClm��̏�Ԃ����ƂɁA���ݕ��}��toClm��̏�Ԃ����߂�
' ���̏ꍇ�͂��̂܂܃R�s�[�A���̏ꍇ�͔������]�R�s�[
Private Sub copyDrawDownToTreadling(firstRow As Integer, lastRow As Integer, fromClm As Integer, toClm As Integer)
    Dim i As Integer
    
    If kind = "��" Then ' �V�����Ȃ�
        Range(Cells(firstRow, fromClm), Cells(lastRow, fromClm)).Copy Cells(y2, toClm)
    Else ' �낭�뎮�Ȃ�
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

' Tie-Up���ŏ����珑����Ă��邩�ǂ����B
' �����Ƃ��낪��ł�����Ώ�����Ă���Ƃ݂Ȃ�
' �P�������������`�F�b�N�������B
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

' �g�D�}�ɍ��}�X������ŏ��̍s�𓾂�
Private Function firstRowOnDrawDown() As Integer
    Dim first As Integer
    Dim k As Integer
    Dim l As Integer
    
    first = 0
    For l = y2 To y3
        '�g�D�}�Ώۗ�ō��̃}�X��T���B���̃}�X������΂��̍s�͗L��
        For k = x0 To x1
            If Cells(l, k).Interior.ColorIndex = 1 Then
                first = l
                Exit For
            End If
        Next
        If first > 0 Then
            Exit For '�L���s������΂������J�n�s�Ȃ̂ŏI��
        End If
    Next
    firstRowOnDrawDown = first
End Function

' �g�D�}�ɍ��}�X������Ō�̍s�𓾂�
Private Function lastRowOnDrawDown() As Integer
    Dim last As Integer
    Dim k As Integer
    Dim l As Integer
    
    last = 0
    For l = y3 To y2 Step -1
        '���ݕ��}�Ώۗ�ō��̃}�X��T���B���̃}�X������΂��̍s�͗L��
        For k = x0 To x1
            If Cells(l, k).Interior.ColorIndex = 1 Then
                last = l
                Exit For
            End If
        Next
        If last > 0 Then
            Exit For '�L���s������΂��������X�g�s�Ȃ̂ŏI��
        End If
    Next
    lastRowOnDrawDown = last
End Function