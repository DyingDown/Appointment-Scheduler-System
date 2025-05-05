Attribute VB_Name = "AddUpdatePayments"
Public Sub AddOrderPayment(orderId As Long, _
                           paymentMethod As String, amount As Double, _
                           giftCardNo As Long)

    On Error GoTo Catch
    
    If amount = 0 Then ' 如果付钱为零，没必要记录
        Exit Sub
    End If
    
    ' 参数检查
    If orderId <= 0 Then Err.Raise vbObjectError + 4002, , "Order ID 非法" ' ===== TODO: 是否验证orderID存在 =====
    If Trim(paymentMethod) = "" And IsValidPaymentMethod(paymentMethod) Then Err.Raise vbObjectError + 4003, , "支付方式为空或不合法"
    If amount < 0 And paymentMethod <> "Cash" Then Err.Raise vbObjectError + 4004, , "非现金支付金额必须大于 0"
    If paymentMethod = "Gift Card" And giftCardNo = 0 Then
        Err.Raise vbObjectError + 4005, , "礼品卡支付必须提供卡号"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OrderPayments")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("OrderPaymentsTable")
    
    If tbl Is Nothing Then
        MsgBox "未找到 OrderPaymentsTable", vbExclamation
        Exit Sub
    End If

    ' 添加行
    Dim newRow As listRow
    Set newRow = tbl.ListRows.Add

    With newRow.Range
        .Cells(1, 1).Value = GetNextIndexForSheet(ws.Name)
        .Cells(1, 2).Value = orderId
        .Cells(1, 3).Value = paymentMethod
        .Cells(1, 4).Value = amount
        .Cells(1, 5).Value = giftCardNo
        .Cells(1, 6).Value = Now               ' Timestamp
        .Cells(1, 7).Value = False             ' IsDeleted
        .Cells(1, 8).Value = ""                ' DeletedTime
        .Cells(1, 9).Value = Environ("Username") ' 获取当前操作系统登录的用户名
        .Cells(1, 10).Value = ""               ' DeletedBy
    End With
    
    If paymentMethod = "Gift Card" Then
        Call UpdateGiftCardByNumber(giftCardNo, amount)
    End If
    On Error GoTo Catch

    MsgBox "Payment 已添加！", vbInformation
    
    Exit Sub

Catch:
    MsgBox "添加 Payment 失败：" & Err.Description, vbCritical, "错误代码 " & Err.Number
End Sub


