Attribute VB_Name = "SetUpOrderSystem"
' 负责主表单初始化以及整体控制的部分
Sub SetUpOrderSystem()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsOrders As Worksheet, wsServices As Worksheet, wsLists As Worksheet
    Dim wsPayments As Worksheet, wsGiftCards As Worksheet
    Dim wsLogs As Worksheet, wsIndex As Worksheet

    ' 删除旧表
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Orders").Delete
    wb.Sheets("Services").Delete
    wb.Sheets("Lists").Delete
    wb.Sheets("OrderPayments").Delete
    wb.Sheets("GiftCards").Delete
    wb.Sheets("Logs").Delete
    wb.Sheets("IndexStorage").Visible = xlSheetVisible
    wb.Sheets("IndexStorage").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新建表
    Set wsOrders = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsOrders.Name = "Orders"
    Set wsServices = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsServices.Name = "Services"
    Set wsLists = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsLists.Name = "Lists"
    Set wsPayments = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsPayments.Name = "OrderPayments"
    Set wsGiftCards = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsGiftCards.Name = "GiftCards"
    Set wsLogs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsLogs.Name = "Logs"
    Set wsIndex = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsIndex.Name = "IndexStorage"
    
    ' 设置 Services 表头
    SetUpServices wsServices
    ' 设置 Lists 表内容
    SetUpLists wsLists
    ' 设置命名区域
    SetUpNamedRanges wb, wsLists, wsServices
    ' 设置 Orders 表头及数据表
    SetUpOrdersTable wsOrders
    ' 设置 Payments 表格
    SetUpPayments wsPayments
    ' 设置 GiftCards 表格
    SetUpGiftCards wsGiftCards
    ' 设置Logs表格
    SetUpLogs wsLogs
    ' 设置 Index 表格
    SetUpIndexStorage wsIndex

    MsgBox "各类表单初始化完成！", vbInformation
    Call InjectToAllSheets
End Sub


' 负责 Services 工作表的设置
Sub SetUpServices(ws As Worksheet)
    ws.Range("A1:D1").Value = Array("Service Name", "Duration", "Price", "Abbr.")
End Sub

' 负责 Lists 工作表的设置
Sub SetUpLists(ws As Worksheet)
    ws.Range("A1:A4").Value = Application.WorksheetFunction.Transpose(Array("Unspecified", "Male", "Female", "By Name"))
    ws.Range("B1:B5").Value = Application.WorksheetFunction.Transpose(Array("Jason", "David", "Lee", "Steven", "Alisa", "Kevin", "Rose"))
    ws.Range("C1:C4").Value = Application.WorksheetFunction.Transpose(Array("Booked", "Arrived", "Paid", "Cancelled"))
    ws.Range("D1:D6").Value = Application.WorksheetFunction.Transpose(Array("Cash", "POS", "Zelle", "Venmo", "Apple Pay", "Gift Card"))
    ws.Range("E1:E3").Value = Application.WorksheetFunction.Transpose(Array("Active", "Used", "Expired"))
    ws.Range("F1:F5").Value = Application.WorksheetFunction.Transpose(Array("Add Order", "Update Order", "Add Payment", "Delete Payment", "Issue Gift Card"))
End Sub

' 设置动态命名区域
Sub SetUpNamedRanges(wb As Workbook, wsLists As Worksheet, wsServices As Worksheet)
    With wb.names
        .Add Name:="TechnicianReqList", RefersTo:="=Lists!$A$1:$A$" & wsLists.Cells(wsLists.Rows.Count, "A").End(xlUp).row
        .Add Name:="TechnicianList", RefersTo:="=Lists!$B$1:$B$" & wsLists.Cells(wsLists.Rows.Count, "B").End(xlUp).row
        .Add Name:="StatusList", RefersTo:="=Lists!$C$1:$C$" & wsLists.Cells(wsLists.Rows.Count, "C").End(xlUp).row
        .Add Name:="PaymentMethodList", RefersTo:="=Lists!$D$1:$D$" & wsLists.Cells(wsLists.Rows.Count, "D").End(xlUp).row
        .Add Name:="ProjectTypeList", RefersTo:="=Services!$A$2:$A$100"
        .Add Name:="CardStatusList", RefersTo:="=Lists!$E$1:$E$" & wsLists.Cells(wsLists.Rows.Count, "E").End(xlUp).row
        .Add Name:="OperationList", RefersTo:="=Lists!$F$1:$F$" & wsLists.Cells(wsLists.Rows.Count, "F").End(xlUp).row
    End With
End Sub

' 负责 Orders 表格的设置
Sub SetUpOrdersTable(ws As Worksheet)
    Dim headers As Variant
    headers = Array("Index", "Date", "Scheduled Time", "Start Time", "Service", "Technician Requirement", _
                    "Technician", "Phone Number", "Customer Name", "Status", "Payment Time", "Payment Method", "Payment Amount", "Comment")
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers

    ' 把 Orders 表格转换为 ListObject
    Dim tblOrder As ListObject
    Set tblOrder = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:N2"), , xlYes)
    tblOrder.Name = "OrdersTable"
    tblOrder.TableStyle = ""

    ' 格式化列
    With tblOrder.DataBodyRange
        .Columns(2).NumberFormat = "yyyy-mm-dd" ' Date
        .Columns(3).NumberFormat = "hh:mm"      ' Scheduled Time
        .Columns(4).NumberFormat = "hh:mm"      ' Start Time
        .Columns(11).NumberFormat = "hh:mm"     ' Payment Time
        .Columns(8).NumberFormat = "@"          ' Phone
        .Interior.ColorIndex = xlColorIndexNone ' 设置背景为透明
        .Font.Color = RGB(0, 0, 0)              ' 字体颜色黑色
    End With

    ' 设置数据验证
    With ws
        .Range("OrdersTable[Service]").Validation.Add Type:=xlValidateList, Formula1:="=ProjectTypeList"
        .Range("OrdersTable[Technician Requirement]").Validation.Add Type:=xlValidateList, Formula1:="=TechnicianReqList"
        .Range("OrdersTable[Technician]").Validation.Add Type:=xlValidateList, Formula1:="=TechnicianList"
        .Range("OrdersTable[Status]").Validation.Add Type:=xlValidateList, Formula1:="=StatusList"
        .Range("OrdersTable[Payment Method]").Validation.Add Type:=xlValidateList, Formula1:="=PaymentMethodList"
    End With

    ' 冻结首行
    With ws
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With

    ' 表头样式：绿色背景白字
    With tblOrder.HeaderRowRange
        .Interior.Color = RGB(83, 120, 53)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

' 设置 OrderPayments 表格，支持软删除字段和创建/删除人信息
Sub SetUpPayments(ws As Worksheet)
    On Error Resume Next
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "OrderPayments"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' 添加表头
    Dim paymentHeaders As Variant
    paymentHeaders = Array("Payment ID", "Order ID", "Payment Method", "Amount", "Gift Card No", "Timestamp", "IsDeleted", "DeletedTime", "CreatedBy", "DeletedBy")
    ws.Range("A1:J1").Value = paymentHeaders

    ' 转为表格对象
    Dim tblPayments As ListObject
    Set tblPayments = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:J2"), , xlYes)
    tblPayments.Name = "OrderPaymentsTable"
    tblPayments.TableStyle = ""

    ' 格式化列
    With tblPayments.DataBodyRange
        .Columns(4).NumberFormat = "#,##0.00"            ' Amount
        .Columns(6).NumberFormat = "yyyy-mm-dd hh:mm"    ' Timestamp
        .Columns(8).NumberFormat = "yyyy-mm-dd hh:mm"    ' DeletedTime
        .Font.Color = RGB(0, 0, 0)
        .Interior.ColorIndex = xlColorIndexNone
    End With

    ' 数据验证
    With ws
        .Range("OrderPaymentsTable[Payment Method]").Validation.Add Type:=xlValidateList, Formula1:="=PaymentMethodList"
    End With

    ' 表头样式
    With tblPayments.HeaderRowRange
        .Interior.Color = RGB(83, 120, 53)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' 冻结首行
    With ws
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub


' 负责 GiftCards 表格的设置
Sub SetUpGiftCards(ws As Worksheet)
    On Error Resume Next
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "GiftCards"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' 添加表头
    Dim giftHeaders As Variant
    giftHeaders = Array("Gift Card No", "Balance", "Status", "Issued By", "Created Time")
    ws.Range("A1:E1").Value = giftHeaders

    ' 转为表格对象
    Dim tblGiftCards As ListObject
    Set tblGiftCards = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E2"), , xlYes)
    tblGiftCards.Name = "GiftCardsTable"
    tblGiftCards.TableStyle = ""

    ' 格式化列
    With tblGiftCards.DataBodyRange
        .Columns(2).NumberFormat = "#,##0.00"            ' Balance
        .Columns(5).NumberFormat = "yyyy-mm-dd hh:mm"    ' Created Time
        .Font.Color = RGB(0, 0, 0)
        .Interior.ColorIndex = xlColorIndexNone
    End With

    ' 数据验证
    With ws
        .Range("GiftCardsTable[Status]").Validation.Add Type:=xlValidateList, Formula1:="=StatusList"
    End With

    ' 表头样式
    With tblGiftCards.HeaderRowRange
        .Interior.Color = RGB(83, 120, 53)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' 冻结首行
    With ws
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

' 设置日志表，用于记录操作日志
Sub SetUpLogs(ws As Worksheet)
    On Error Resume Next
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "Logs"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' 添加表头
    Dim logHeaders As Variant
    logHeaders = Array("Log ID", "Operation Type", "Target Table", "Target ID", "User", "Timestamp")
    ws.Range("A1:F1").Value = logHeaders

    ' 转为表格对象
    Dim tblLogs As ListObject
    Set tblLogs = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F2"), , xlYes)
    tblLogs.Name = "LogsTable"
    tblLogs.TableStyle = ""

    ' 格式化列
    With tblLogs.DataBodyRange
        .Columns(5).NumberFormat = "yyyy-mm-dd hh:mm:ss" ' Timestamp
        .Font.Color = RGB(0, 0, 0)
        .Interior.ColorIndex = xlColorIndexNone
    End With

    ' 表头样式
    With tblLogs.HeaderRowRange
        .Interior.Color = RGB(83, 120, 53)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' 冻结首行
    With ws
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

Sub SetUpIndexStorage(ws As Worksheet)
    ws.Visible = xlSheetVeryHidden ' 设置为非常隐藏
    ws.Cells(1, 1).Value = "SheetName"
    ws.Cells(1, 2).Value = "CurrentIndex"

    ' 指定需要初始化的工作表列表
    ' Dim sheetNames As Variant
    ' sheetNames = Array("Orders", "OrderPayments", "GiftCards", "Logs") ' 这里添加你需要初始化的工作表名
    
      ' 设置初始值
    ws.Cells(1, 1).Value = "Orders"
    ws.Cells(1, 2).Value = 0 ' Orders 初始 index

    ws.Cells(2, 1).Value = "OrderPayments"
    ws.Cells(2, 2).Value = 0 ' OrderPayments 初始 index

    ws.Cells(3, 1).Value = "GiftCards"
    ws.Cells(3, 2).Value = 0 ' GiftCards 初始 index

    ws.Cells(4, 1).Value = "Logs"
    ws.Cells(4, 2).Value = 0 ' Logs 初始 index
End Sub
