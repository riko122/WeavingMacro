Attribute VB_Name = "common"
Option Explicit

' (row, clm)�̃Z���̒l��ǂݍ���ŕԂ��B�Z���̒l���Ȃ����́Ainit���Z���ɓ���ĕԂ�
Public Function readCellValue(row As Integer, clm As Integer, init As Integer)
    If Cells(row, clm) = "" Then
        Cells(row, clm) = init
    End If
    readCellValue = Cells(row, clm)
End Function

