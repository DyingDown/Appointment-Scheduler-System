Attribute VB_Name = "GetNextIndexBySheetName"
Function GetNextIndexForSheet(sheetName As String) As Long
    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Sheets("IndexStorage")

    Dim matchCell As Range
    Set matchCell = storeSheet.Columns("A").Find(What:=sheetName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not matchCell Is Nothing Then
        ' �ҵ���¼����ȡ������
        Dim currentIndex As Long
        currentIndex = storeSheet.Cells(matchCell.row, 2).Value + 1
        storeSheet.Cells(matchCell.row, 2).Value = currentIndex
        GetNextIndexForSheet = currentIndex
    Else
        ' û�м�¼������ 0 ���ʼ������
        MsgBox "δ�ҵ��ù������ Index ��¼��", vbExclamation
        GetNextIndexForSheet = 0
    End If
End Function


