Attribute VB_Name = "UpdateOrderRowStyle"
Sub UpdateOrderRowStyles(ByVal Target As Variant)
    Dim rowRange As Range
    Dim statusCell As Range
    Dim status As String
    
    ' 判断参数类型，如果是 ListRow，就获取它的 Range
    If TypeName(Target) = "ListRow" Then
        Set rowRange = Target.Range
        Set statusCell = rowRange.Columns(10)
    ' 如果是 Range，直接使用
    ElseIf TypeName(Target) = "Range" Then
        Set rowRange = Target
        Set statusCell = rowRange.Cells(1, 10) ' 假设第 10 列就是状态列
    Else
        ' 不是 ListRow 或 Range 类型，抛出错误
        MsgBox "Invalid parameter type"
        Exit Sub
    End If

    ' 获取状态值并进行处理
    status = Trim(statusCell.Value)
    
    If LCase(status) = "booked" Then
        rowRange.Interior.Color = RGB(226, 239, 218)
        rowRange.Font.Color = RGB(83, 120, 53)
    ElseIf LCase(status) = "arrived" Then
        rowRange.Interior.Color = RGB(255, 199, 206)
        rowRange.Font.Color = RGB(156, 0, 6)
    ElseIf LCase(status) = "cancelled" Then
        rowRange.Interior.Color = RGB(250, 250, 250)
        rowRange.Font.Color = RGB(127, 127, 127)
    Else
        rowRange.Interior.Color = RGB(255, 255, 255)
        rowRange.Font.Color = RGB(0, 0, 0)
    End If
End Sub

Sub UpdateAllOrderRowStyles()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Orders").ListObjects("OrdersTable")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        ' 传递每一行到 UpdateOrderRowStyles 方法
        UpdateOrderRowStyles tbl.ListRows(i)
    Next i
End Sub

