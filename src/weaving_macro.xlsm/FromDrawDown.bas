Attribute VB_Name = "FromDrawDown"
'-----------------------------------------------------------------
' �g�D�}���犮�S�ӏ��}���쐬
'   by Riko(https://github.com/riko122/WeavingMacro)
' This software is released under the Mozilla Public License 2.0.
'-----------------------------------------------------------------
Option Explicit

Const HEADER_LINE = 7 ' �w�b�_�[�����̍s��

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
Private Sub init()
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
            y0 = HEADER_LINE + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "�E��"
            x0 = 1
            x1 = x0 + w - 1
            x2 = x1 + 2
            x3 = x2 + f - 1
            y2 = HEADER_LINE + 2
            y3 = y2 + h - 1
            y0 = y3 + 2
            y1 = y0 + n - 1
        Case "����"
            x2 = 1
            x3 = x2 + f - 1
            x0 = x3 + 2
            x1 = x0 + w - 1
            y0 = HEADER_LINE + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "����"
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
    
    ' �N���A�B�w�b�_�[�ȊO�̍s��������Ƒ��߂ɍ폜����B
    Rows(HEADER_LINE + 1 & ":" & HEADER_LINE + n + h + 100).Select
    Selection.Delete Shift:=xlUp

    ' �Ώ۔͈͂̃}�X�̍��������낦��B
    Rows(HEADER_LINE + 1 & ":" & HEADER_LINE + n + h + 5).Select
    Selection.RowHeight = 11
    
    Call writeGrid(y0, y1, x0, x1) ' ���L�ʂ������̃}�X��
    Call writeGrid(y0, y1, x2, x3) ' �^�C�A�b�v�����̃}�X��
    Call writeGrid(y2, y3, x0, x1) ' �g�D�}�����̃}�X��
    Call writeGrid(y2, y3, x2, x3) ' ���ݖؕ����̃}�X��
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
    ' ���ݖ؂𓥂񂾂�A���L���オ�邩�����邩��ǂݎ��B
    kind = Cells(6, 46)  ' ������
    
    ReDim status(n)
    ReDim shaft_row(n)
    
    ' ���L�̒ʂ����}�Ɠ��ݕ��}�͈̔͂��N���A
    Range(Cells(y0, x0), Cells(y1, x1)).Interior.ColorIndex = xlNone
    Range(Cells(y2, x2), Cells(y3, x3)).Interior.ColorIndex = xlNone
    
    first_row = getFirstRow(y2, y3, x0, x1)
    If first_row = 0 Then
        MsgBox ("�g�D�}�������h���Ă��܂���")
        Exit Sub
    End If
    last_row = getLastRow(y2, y3, x0, x1)
    
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
        If a > 0 Then
            For j = 0 To a - 1
                If status(j) = status(a) Then
                    Cells(y0 + j, i).Interior.ColorIndex = 1
                    found = True
                    ' a�͍ė��p
                    Exit For
                End If
            Next j
        End If
        ' ������Ȃ������ꍇ(a=0�̎���)�́A�V�����s�Ȃ̂�y0+a����������
        If found = False Then
            If a >= n Then
                MsgBox ("���̑g�D�}����������ɂ͑��L������܂���")
                Exit Sub
            End If
            Cells(y0 + a, i).Interior.ColorIndex = 1
            a = a + 1 ' a�͎����g��
        End If
Continue:
    Next i
    
    If getFirstRow(y0, y1, x2, x3) = 0 Then
        MsgBox ("���݂̂Ƃ���^�C�A�b�v��������Ă�����̂ɂ����Ή����Ă��܂���")
        Exit Sub
    End If
    
    ' ���ݖ؂��l����
    If getMaxShaftPerPedal = 1 Then
        For i = y0 To y1
            For j = x1 To x0 Step -1
                ' ���L�̒ʂ�����i�s�ڂōŏ��ɏo�Ă��鍕�����T��
                If Cells(i, j).Interior.ColorIndex = 1 Then
                    ' Tie-up�ł��̍s���������T��
                    found = False
                    For k = x2 To x3
                        If Cells(i, k).Interior.ColorIndex = 1 Then
                            Call copyDrawDownToTreadling(first_row, last_row, j, k)
                            found = True
                            Exit For
                        End If
                    Next k
                    If found Then
                        Exit For ' ���̑��L�ɊY�����铥�ݕ��͏������̂ŏI���
                    Else ' �S�����Ă�������Ȃ������ꍇ
                        MsgBox ("���̑g�D�}����������ɂ̓^�C�A�b�v���s�K�؂ł�")
                        Exit Sub
                    End If
                End If
            Next j
        Next i
    Else
        For k = first_row To last_row
            ' �e�s�ɂ��āA���L�����߂��������i�o�����ォ�j��ǂݎ��
            ' �S��ǂ܂Ȃ��Ă��A���L�̖������ł����i���Ƃ͓����p�^�[��������j
            a = 0
            For i = y0 To y1
                For j = x1 To x0 Step -1
                    ' ���L�̒ʂ�����i�s�ڂōŏ��ɏo�Ă��鍕�����T��
                    If Cells(i, j).Interior.ColorIndex = 1 Then
                        shaft_row(a) = getTieupStatus(k, j)
                        a = a + 1
                        Exit For
                    End If
                Next j
            Next i
            ' �^�C�A�b�v�ŁA���̍s���������T��
            For j = x2 To x3
                found = True
                For a = 0 To n - 1
                    If Cells(y0 + a, j).Interior.ColorIndex <> shaft_row(a) Then
                        found = False
                        Exit For ' ������̂Ŏ��̗��T��
                    End If
                Next a
                ' ���Ȃ��܂�For a���I������ꍇ�́Aj���Y�����铥�ݖ�
                If found Then
                    Cells(k, j).Interior.ColorIndex = 1
                    Exit For
                End If
            Next j
            If found = False Then ' �^�C�A�b�v�S�����Ă�������Ȃ��ꍇ
                MsgBox ("���̑g�D�}����������ɂ̓^�C�A�b�v���s�K�؂ł�")
                Exit Sub
            End If
        Next k
    End If
End Sub

' ���L���ʂ��Ă���clm�񂪁A�g�D�}��row�s�ڂō������������ŁATieUp�̏�Ԃ�����
' �Ⴆ��4�����L��1��4�������s�́A���Ȃ�1��4�������^�C�A�b�v�̗�A
' ���Ȃ�2��3�������^�C�A�b�v�̗��T���̂ŁA���������ɂ���ĕԂ����̂��t�B
Private Function getTieupStatus(row As Integer, clm As Integer) As Integer

    If kind = "��" Then ' �V�����ȂǁB�g�D�}�ō�����΃^�C�A�b�v�ō���
        If Cells(row, clm).Interior.ColorIndex = 1 Then
            getTieupStatus = 1
        Else
            getTieupStatus = xlNone
        End If
    Else ' �낭�뎮�ȂǁB�g�D�}�Ŕ�����΃^�C�A�b�v�ō���
        If Cells(row, clm).Interior.ColorIndex <> 1 Then
            getTieupStatus = 1
        Else
            getTieupStatus = xlNone
        End If
    End If
End Function

' �g�D�}��from_clm��̏�Ԃ����ƂɁA���ݕ��}��to_clm��̏�Ԃ����߂�
' ���̏ꍇ�͂��̂܂܃R�s�[�A���̏ꍇ�͔������]�R�s�[
Private Sub copyDrawDownToTreadling(first_row As Integer, last_row As Integer, _
                                    from_clm As Integer, to_clm As Integer)
    Dim i As Integer
    
    If kind = "��" Then ' �V�����Ȃ�
        Range(Cells(first_row, from_clm), Cells(last_row, from_clm)).Copy Cells(first_row, to_clm)
    Else ' �낭�뎮�Ȃ�
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

' ��{�̓��ݖ؂ɂȂ����Ă��鑎�L�g�̍ő吔
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

