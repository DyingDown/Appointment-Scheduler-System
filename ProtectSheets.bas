Attribute VB_Name = "ProtectSheets"
'-------------------------------
' �� SelectionChange ����д�뵥��������
'-------------------------------
Sub InjectSelectionChange(ws As Worksheet)
    Dim moduleCode As String
    
    moduleCode = _
    "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "     ' ���ѡ���˱�����������ֱ��ѡ�ر��" & vbCrLf & _
    "     If Target.row > 1 Then" & vbCrLf & _
    "        MsgBox ""�벻Ҫ�ֶ��޸����ݣ�ʹ�� DailySheet ҳ�水ť�����������޸ġ�"", vbExclamation" & vbCrLf & _
    "    End If" & vbCrLf & _
    "End Sub"
    
    With ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, moduleCode
    End With
End Sub

'-------------------------------
' ����ע�뵽���Ź�����
'-------------------------------
Sub InjectToAllSheets()
    Dim sheetNames As Variant
    Dim i As Long
    
    ' �������г�����Ҫ���� SelectionChange �߼��Ĺ�������
    sheetNames = Array("Orders", "OrderPayments", "GiftCards", "Logs")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        InjectSelectionChange ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
    Next i
    
    MsgBox "�ѳɹ�ע�� SelectionChange �¼���ָ��������", vbInformation
End Sub

