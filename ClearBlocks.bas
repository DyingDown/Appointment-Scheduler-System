Attribute VB_Name = "ClearBlocks"
Sub ClearBlocks(ByVal targetDate As Date)
    Dim wsDaily As Worksheet
    Dim startCol As Long
    Dim endCol As Long
    Dim appointmentTimeCol As Long
    Dim targetRange As Range
    Dim currentRow As Long
    Dim col As Long

    ' ��ȡ�������������Ű��"�Ű�_����"��
    Set wsDaily = ThisWorkbook.Sheets("�Ű�_" & Day(targetDate))
    
    ' �ҵ�ԤԼʱ�����ڵ���
    appointmentTimeCol = 0 ' ��ʼ���к�
    For col = 1 To wsDaily.Columns.Count
        If wsDaily.Cells(1, col).Value = "ԤԼʱ��" Then ' ���� "ԤԼʱ��" �ڵ�3��
            appointmentTimeCol = col
            Exit For
        End If
    Next col
    
    If appointmentTimeCol = 0 Then
        MsgBox "�Ҳ���ԤԼʱ���У�", vbExclamation
        Exit Sub
    End If
    
    ' ���������Χ�Ľ����У�ԤԼʱ���� - 2��
    endCol = appointmentTimeCol - 2
    
    ' ����Ŀ�귶Χ����C6��Ŀ������У��з�Χ��6��69
    Set targetRange = wsDaily.Range(wsDaily.Cells(6, 3), wsDaily.Cells(69, endCol))
    
    ' ������ݺͱ�����ɫ
    targetRange.Clear ' ������ݡ�������ɫ�͸�ʽ
End Sub



