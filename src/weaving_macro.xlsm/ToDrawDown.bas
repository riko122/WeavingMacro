Attribute VB_Name = "ToDrawDown"
'-----------------------------------------------------------------
' ���L�̒ʂ����}�E�^�C�A�b�v�E���ݕ��}����A�g�D�}��z�F�}���쐬
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
Dim x9 As Integer '�܎��F�w���
Dim y9 As Integer '�o���F�w��s

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
            x9 = x3 + 2
            y9 = HEADER_LINE + 2
            y0 = y9 + 2
            y1 = y0 + n - 1
            y2 = y1 + 2
            y3 = y2 + h - 1
        Case "�E��"
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
        Case "����"
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
        Case "����"
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

' �������{�^���N���b�N�Ŏ��s�B
Public Sub clearToDrawDown()

    Call init

    ' �N���A�B�w�b�_�[�ȊO�̍s��������Ƒ��߂ɍ폜����B
    Rows(HEADER_LINE + 1 & ":" & HEADER_LINE + n + h + 100).Select
    Selection.Delete Shift:=xlUp

    ' �Ώ۔͈͂̃}�X�̍��������낦��B
    Rows(HEADER_LINE + 1 & ":" & HEADER_LINE + n + h + 5).Select
    Selection.RowHeight = 11

    Call writeGrid(y0, y1, x0, x1) ' ���L�ʂ������̃}�X��
    Call writeGrid(y0, y1, x2, x3) ' �^�C�A�b�v�����̃}�X��

    ' �o���F�w��s�Ɂu�o���̐F�v�Ə���
    Range(Cells(y9, x2), Cells(y9, x3)).Select
    If x1 < x2 Then
        ActiveCell.FormulaR1C1 = "�o���̐F"
    Else
        ActiveCell.FormulaR1C1 = "�o���̐F"
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

    ' �܎��F�w���Ɂu�܎��̐F�v�Ə���
    Range(Cells(y0, x9), Cells(y1, x9)).Select
    If y1 < y2 Then
        ActiveCell.FormulaR1C1 = "�܎��̐F"
    Else
        ActiveCell.FormulaR1C1 = "�܎��̐F"
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

    Call writeGrid(y2, y3, x0, x1) ' �g�D�}�����̃}�X��
    Call writeGrid(y2, y3, x2, x3) ' ���ݖؕ����̃}�X��
End Sub

' �g�D�}�{�^���N���b�N�Ŏ��s�B
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
    ' ���ݖ؂𓥂񂾂�A���L���オ�邩�����邩��ǂݎ��B
    kind = Cells(6, 40)  ' ������
    
    ' �g�D�}��������������
    Call writeGrid(y2, y3, x0, x1) ' �g�D�}�����̃}�X��
    Range(Cells(y2, x0), Cells(y3, x1)).Interior.ColorIndex = xlNone
    
    ReDim init_row_status(w)
    ReDim current_row_status(w)

    first_clm = getFirstColumn(y0, y1, x0, x1)
    If first_clm = 0 Then
        MsgBox ("���L�̒ʂ����}�������h���Ă��܂���")
        Exit Sub
    End If
    last_clm = getLastColumn(y0, y1, x0, x1)
    
    first_row = getFirstRow(y2, y3, x2, x3)
    If first_row = 0 Then
        MsgBox ("���ݕ��}�������h���Ă��܂���")
        Exit Sub
    End If
    last_row = getLastRow(y2, y3, x2, x3)
    
    init_row_status = setInitRowStatus()
    
    For l = first_row To last_row
        current_row_status = getCurrentRowStatus(l, init_row_status)
        For i = first_clm To last_clm
            ' �o������̃}�X�͍����h��
            If current_row_status(i) = True Then
                Cells(l, i).Interior.ColorIndex = 1
            End If
        Next
    Next
End Sub

' �z�F�}�{�^���N���b�N�Ŏ��s�B
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
    ' ���ݖ؂𓥂񂾂�A���L���オ�邩�����邩��ǂݎ��B
    kind = Cells(6, 40)  ' ������
    
    ' �g�D�}��������������(�F�͔z�F�}�͏㏑������̂œ��ɓh��Ȃ����Ȃ��j
    Call writeGrid(y2, y3, x0, x1) ' �g�D�}�����̃}�X��
    
    ReDim init_row_status(w)
    ReDim current_row_status(w)
    ReDim before_row_status(w)

    first_clm = getFirstColumn(y0, y1, x0, x1)
    If first_clm = 0 Then
        MsgBox ("���L�̒ʂ����}�������h���Ă��܂���")
        Exit Sub
    End If
    last_clm = getLastColumn(y0, y1, x0, x1)

    first_row = getFirstRow(y2, y3, x2, x3)
    If first_row = 0 Then
        MsgBox ("���ݕ��}�������h���Ă��܂���")
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
            ' �o������̃}�X�͌o���̐F�œh��B�����łȂ���Έ܎��̐F�œh��
            If current_row_status(i) = True Then
                Cells(l, i).Interior.color = Cells(y9, i).Interior.color
                ' �O�̍s���o������Ȃ�A��̌r�����Ȃ��ɂ���
                If before_row_status(i) = True Then
                    Range(Cells(l, i), Cells(l, i)).Borders(xlEdgeTop).LineStyle = xlNone
                End If
            Else
                Cells(l, i).Interior.color = Cells(l, x9).Interior.color
                ' ���̃Z�����܎�����Ȃ�A���̌r�����Ȃ��ɂ���
                If i > first_clm And current_row_status(i - 1) = False Then
                    Range(Cells(l, i), Cells(l, i)).Borders(xlEdgeLeft).LineStyle = xlNone
                End If
            End If
        Next
        before_row_status = current_row_status
    Next
End Sub

' ���݂̍s�̊e�Z���ɂ��āA�o������ɂȂ��Ă����True, �����łȂ����False��z��ɓo�^����
Private Function getCurrentRowStatus(ByVal row As Integer, ByRef init_row_status() As Boolean) As Boolean()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim current_row_status() As Boolean
    Dim f As Boolean
    
    ' �z��̏�����
    ReDim current_row_status(w)
    current_row_status = init_row_status
   
    ' ���ݕ��}�ō����}�X��T���B
    For j = x2 To x3
        If (Cells(row, j).Interior.ColorIndex = 1) Then
            ' ���̃}�X�̗�̃^�C�A�b�v�ō��̃}�X��T��
            For k = y0 To y1
                If Cells(k, j).Interior.ColorIndex = 1 Then
                    ' ���̃}�X�̍s�̑��L�̒ʂ�����������̑g�D�_�͌o������ɏo��̂�True�ɂ���
                    ' ���N�����Ȃ�A�܎�����ɏo��̂�False�ɂ���i���̂��ߏ������Ō���True�ɂ��Ă���j
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

' �e�s�̏�ԏ������Bkind�ɉ����ĕς��B
Private Function setInitRowStatus() As Boolean()
    Dim init_row_status() As Boolean
    Dim i As Integer
    Dim k As Integer
    
    ReDim init_row_status(w)
    ' �z��̏�����
    For i = x0 To x1
        If kind = "��" Then ' �V�����ȂǁBfalse(�܎�����)�ŏ�����
            init_row_status(i) = False
        Else ' �낭�뎮�ȂǁB��{true(�o������)�ŏ������B�A����H��false�B
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

