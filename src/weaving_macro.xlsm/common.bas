Attribute VB_Name = "common"
'-----------------------------------------------------------------
' ����Function�A����Sub�z�u�p���W���[��
'   by Riko(https://github.com/riko122/WeavingMacro)
' This software is released under the Mozilla Public License 2.0.
'-----------------------------------------------------------------
Option Explicit

' (row, clm)�̃Z���̒l��ǂݍ���ŕԂ��B�Z���̒l���Ȃ����́A�����l���Z���ɓ���ĕԂ�
Public Function readCellValue(row As Integer, clm As Integer, default As Integer)
    If Cells(row, clm) = "" Then
        Cells(row, clm) = default
    End If
    readCellValue = Cells(row, clm)
End Function

' �w��͈͂Ƀ}�X�ڂ�����
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

' �w��͈͓��ō��}�X������ŏ��̍s�𓾂�
Public Function getFirstRow(first_row As Integer, last_row As Integer, _
                            first_clm As Integer, last_clm As Integer) As Integer
    Dim first As Integer
    Dim k As Integer
    Dim l As Integer
    
    first = 0
    For l = first_row To last_row
        '�Ώۍs�ō��̃}�X��T���B���̃}�X������΂��̍s�͗L��
        For k = first_clm To last_clm
            If Cells(l, k).Interior.ColorIndex = 1 Then
                first = l
                Exit For
            End If
        Next
        If first > 0 Then
            Exit For '�L���s������΂������J�n�s�Ȃ̂ŏI��
        End If
    Next
    getFirstRow = first
End Function

' �w��͈͓��ō��}�X������Ō�̍s�𓾂�
Public Function getLastRow(first_row As Integer, last_row As Integer, _
                           first_clm As Integer, last_clm As Integer) As Integer
    Dim last As Integer
    Dim k As Integer
    Dim l As Integer
    
    last = 0
    For l = last_row To first_row Step -1
        '�Ώۍs�ō��̃}�X��T���B���̃}�X������΂��̍s�͗L��
        For k = first_clm To last_clm
            If Cells(l, k).Interior.ColorIndex = 1 Then
                last = l
                Exit For
            End If
        Next
        If last > 0 Then
            Exit For '�L���s������΂��������X�g�s�Ȃ̂ŏI��
        End If
    Next
    getLastRow = last
End Function

' �w��͈͓��ō��}�X������ŏ��̗�𓾂�
Public Function getFirstColumn(first_row As Integer, last_row As Integer, _
                               first_clm As Integer, last_clm As Integer) As Integer
    Dim first As Integer
    Dim i As Integer
    Dim k As Integer
    
    first = 0
    For i = first_clm To last_clm
        '�Ώۗ�ō��̃}�X��T���B���̃}�X������΂��̗�͗L��
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

' �w��͈͓��ō��}�X������Ō�̗�𓾂�
Public Function getLastColumn(first_row As Integer, last_row As Integer, _
                              first_clm As Integer, last_clm As Integer) As Integer
    Dim last As Integer
    Dim i As Integer
    Dim k As Integer
    
    last = 0
    For i = last_clm To first_clm Step -1
        '���L�̒ʂ����}�Ώۍs�ŁA���̃}�X��T���B���̃}�X������΂��̍s�͗L��
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
