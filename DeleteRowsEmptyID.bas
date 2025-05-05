Attribute VB_Name = "DeleteRowsEmptyID"
Sub DeleteRowsWithEmptyFirstColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' ����Ŀ�깤����
    Set ws = ThisWorkbook.Sheets("Orders")
    
    ' �ҵ����һ�У�����������������������ܣ�
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' �����һ�п�ʼ���ϱ���������ɾ���е�������
    For i = lastRow To 1 Step -1
        If Trim(ws.Cells(i, 1).Value) = "" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

