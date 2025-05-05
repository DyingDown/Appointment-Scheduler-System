Attribute VB_Name = "ShowOrderForm"
Public Sub ShowOrderFormFromShape()
    Dim shapeName As String
    Dim orderId As String
    shapeName = Application.Caller ' 比如 "idx_123"
    orderId = ExtractIdFromName(shapeName)
    SelectedOrderId = orderId
    OrderInfo.Show
End Sub


Function ExtractIdFromName(Name As String) As String
    Dim firstUnderscore As Long
    Dim secondUnderscore As Long

    ' 找到第一个下划线
    firstUnderscore = InStr(1, Name, "_")
    If firstUnderscore = 0 Then
        ExtractIdFromName = ""  ' 没有下划线，返回空
        Exit Function
    End If

    ' 从第一个下划线后开始找第二个
    secondUnderscore = InStr(firstUnderscore + 1, Name, "_")

    If secondUnderscore = 0 Then
        ' 没有第二个下划线，取从第一个后面的内容到结尾
        ExtractIdFromName = Mid(Name, firstUnderscore + 1)
    Else
        ' 有第二个，取两者之间的内容
        ExtractIdFromName = Mid(Name, firstUnderscore + 1, secondUnderscore - firstUnderscore - 1)
    End If
End Function
