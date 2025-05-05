Attribute VB_Name = "ShowOrderForm"
Public Sub ShowOrderFormFromShape()
    Dim shapeName As String
    Dim orderId As String
    shapeName = Application.Caller ' ���� "idx_123"
    orderId = ExtractIdFromName(shapeName)
    SelectedOrderId = orderId
    OrderInfo.Show
End Sub


Function ExtractIdFromName(Name As String) As String
    Dim firstUnderscore As Long
    Dim secondUnderscore As Long

    ' �ҵ���һ���»���
    firstUnderscore = InStr(1, Name, "_")
    If firstUnderscore = 0 Then
        ExtractIdFromName = ""  ' û���»��ߣ����ؿ�
        Exit Function
    End If

    ' �ӵ�һ���»��ߺ�ʼ�ҵڶ���
    secondUnderscore = InStr(firstUnderscore + 1, Name, "_")

    If secondUnderscore = 0 Then
        ' û�еڶ����»��ߣ�ȡ�ӵ�һ����������ݵ���β
        ExtractIdFromName = Mid(Name, firstUnderscore + 1)
    Else
        ' �еڶ�����ȡ����֮�������
        ExtractIdFromName = Mid(Name, firstUnderscore + 1, secondUnderscore - firstUnderscore - 1)
    End If
End Function
