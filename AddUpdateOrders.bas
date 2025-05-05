Attribute VB_Name = "AddUpdateOrders"
Public Sub AddOrder(scheduledDate As Date, scheduledTime As Variant, _
                    service As String, req As String, technician As String, phone As String, _
                    customerName As String, comment As String)
     
    On Error GoTo Catch
    ' 参数验证
    If IsMissing(scheduledDate) Or scheduledDate = 0 Then
        Err.Raise vbObjectError + 1001, "AddOrder", "Scheduled Date 不能为空"
    End If

    If IsMissing(scheduledTime) Or IsEmpty(scheduledTime) Then
        Err.Raise vbObjectError + 1002, "AddOrder", "Scheduled Time 不能为空"
    End If

    If Trim(service) = "" Then
        Err.Raise vbObjectError + 1003, "AddOrder", "Service 不能为空"
    End If

    If Trim(req) = "" Then
        Err.Raise vbObjectError + 1004, "AddOrder", "Technician Requirement 不能为空"
    End If

    If Trim(technician) = "" Then
        Err.Raise vbObjectError + 1005, "AddOrder", "Technician 不能为空"
    End If

    If Trim(phone) = "" Then
        Err.Raise vbObjectError + 1006, "AddOrder", "Phone Number 不能为空"
    End If

    If Trim(customerName) = "" Then
        Err.Raise vbObjectError + 1007, "AddOrder", "Customer Name 不能为空"
    End If
    ' 获取 Orders 表
    Dim wsOrders As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set wsOrders = ThisWorkbook.Sheets("Orders")
    On Error GoTo 0

    If wsOrders Is Nothing Then
        MsgBox "未找到 Orders 工作表！", vbExclamation
        Exit Sub
    End If

    ' 获取 ListObject 表格
    On Error Resume Next
    Set tbl = wsOrders.ListObjects("OrdersTable")
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "未找到 Orders 表格！请确保名称为 OrdersTable。", vbExclamation
        Exit Sub
    End If
    
    ' 添加新行
    Dim newRow As listRow
    Set newRow = tbl.ListRows.Add
    Dim nextIndex As Long
    nextIndex = GetNextIndexForSheet(wsOrders.Name)
    With newRow
        .Range(1, 1).Value = nextIndex ' Index
        .Range(1, 2).Value = scheduledDate            ' Date from A1 of active sheet
        .Range(1, 3).Value = scheduledTime       ' Scheduled Time
        .Range(1, 4).Value = ""                  ' Start Time
        .Range(1, 5).Value = service         ' Project Type
        .Range(1, 6).Value = req       ' Technician Requirement
        .Range(1, 7).Value = technician          ' Technician
        .Range(1, 8).Value = phone               ' Phone Number
        .Range(1, 9).Value = customerName        ' Customer Name
        .Range(1, 10).Value = "Booked"            ' Status
        .Range(1, 11).Value = ""                 ' Payment Time
        .Range(1, 12).Value = ""                 ' Payment Method
        .Range(1, 13).Value = ""                 ' Payment Amount‘
        .Range(1, 14).Value = comment
    End With
    
    MsgBox "订单已提交！", vbInformation
    
    Call UpdateOrderRowStyles(newRow)
    
    Call reDrawBlock(scheduledDate)
    Exit Sub
Catch:
    MsgBox "添加订单失败：" & Err.Description, vbCritical, "错误代码 " & Err.Number
    
End Sub


Sub UpdateOrderRowByIndex(orderId As Long, _
                          Optional newDate As Date, _
                          Optional newScheduledTime As Variant, _
                          Optional newService As String, _
                          Optional newTechReq As String, _
                          Optional newTechnician As String, _
                          Optional newPhone As String, _
                          Optional newCustomer As String, _
                          Optional newStatus As String, _
                          Optional newPaymentAmount As Double, _
                          Optional newComment As String)

    Dim wsOrders As Worksheet
    Dim orderRow As Range
    
    If Not IsValidTechnicianReq(newTechReq) Then
        Err.Raise vbObjectError + 1, , "Invalid TechnicianReq"
    End If
    
    If Not IsValidStatus(newStatus) Then
        Err.Raise vbObjectError + 2, , "Invalid Status"
    End If
    
    If Not IsValidTechnician(newTechnician) Then
        Err.Raise vbObjectError + 3, , "Invalid Technician"
    End If
    
    ' ===== TODO: 验证Service是否合法 ======

    ' 获取工作表
    On Error Resume Next
    Set wsOrders = ThisWorkbook.Sheets("Orders")
    On Error GoTo 0

    If wsOrders Is Nothing Then
        MsgBox "未找到 Orders 工作表！", vbExclamation
        Exit Sub
    End If

    ' 查找 Index 所在行
    Set orderRow = wsOrders.Range("A:A").Find(orderId)

    If orderRow Is Nothing Then
        MsgBox orderId & "号订单不存在", vbExclamation
        Exit Sub
    End If
    
    Dim currentTime As Date: currentTime = Time()
    
    ' 更新不为空的字段
    If Not IsMissing(newDate) And Not IsEmpty(newDate) And IsDate(newDate) Then
        orderRow.Offset(0, 1).Value = newDate
    End If
    
    If Not IsMissing(newScheduledTime) And Not IsEmpty(newScheduledTime) And IsValidTime(newScheduledTime) Then
        orderRow.Offset(0, 2).Value = newScheduledTime
    End If
    
    If Not IsMissing(newStatus) And newStatus = "Arrived" Then
        orderRow.Offset(0, 3).Value = currentTime
    End If
    
    If Not IsMissing(newService) And Len(newService) > 0 Then
        orderRow.Offset(0, 4).Value = newService
    End If
    
    If Not IsMissing(newTechReq) And Len(newTechReq) > 0 Then
        orderRow.Offset(0, 5).Value = newTechReq
    End If
    
    If Not IsMissing(newTechnician) And Len(newTechnician) > 0 Then
        orderRow.Offset(0, 6).Value = newTechnician
    End If
    
    If Not IsMissing(newPhone) And Len(newPhone) > 0 Then
        orderRow.Offset(0, 7).Value = newPhone
    End If
    
    If Not IsMissing(newCustomer) And Len(newCustomer) > 0 Then
        orderRow.Offset(0, 8).Value = newCustomer
    End If
    
    If Not IsMissing(newStatus) And Len(newStatus) > 0 Then
        orderRow.Offset(0, 9).Value = newStatus
    End If
    
    If Not IsMissing(newStatus) And newStatus = "Paid" Then
        orderRow.Offset(0, 10).Value = currentTime
    End If
    
    If Not IsMissing(newPaymentAmount) And IsNumeric(newPaymentAmount) And newPaymentAmount <> 0 Then
        orderRow.Offset(0, 12).Value = newPaymentAmount
    End If
    
    If Not IsMissing(newComment) And Len(newComment) > 0 Then
        orderRow.Offset(0, 13).Value = newComment
    End If
    
    MsgBox "订单信息已保存", vbInformation
    Dim rowRange As Range
    Set rowRange = orderRow.ListObject.ListRows(orderRow.row - orderRow.ListObject.Range.row).Range
    Call UpdateOrderRowStyles(rowRange)
End Sub

