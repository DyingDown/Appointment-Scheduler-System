VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderInfo 
   Caption         =   "OrderInfo"
   ClientHeight    =   8920.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8360.001
   OleObjectBlob   =   "OrderInfo.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "OrderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CashInput_AfterUpdate()
    If Not IsNumeric(CashInput.Value) Then
        MsgBox "请输入有效的数字金额", vbExclamation, "无效输入"
        CashInput.Value = ""
    End If
End Sub

Private Sub CancelBtn_Click()
    Unload Me
End Sub

Private Sub GiftCardNoInput_AfterUpdate()
    Set wsGiftCards = ThisWorkbook.Sheets("GiftCards")
    Dim cardNo As Long: cardNo = Me.GiftCardNoInput
    Set cardRow = wsGiftCards.Range("A:A").Find(What:=cardNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not cardRow Is Nothing Then
        If cardRow.Offset(0, 3) <> "Active" Then
            MsgBox "该礼品卡已失效或余额为 0，无法使用。"
            GiftCardInput.Value = ""
        End If
    Else
        MsgBox "卡号不存在"
        GiftCardInput.Value = ""
    End If
End Sub

Private Sub POSInput_AfterUpdate()
    If Not IsNumeric(POSInput.Value) Then
        MsgBox "请输入有效的数字金额", vbExclamation, "无效输入"
        POSInput.Value = ""
    End If
End Sub

Private Sub GiftCardInput_AfterUpdate()
    If Me.GiftCardNoInput.Value = "" Then
        MsgBox "请先输入礼品卡号”"
        GiftCardInput.Value = ""
        Exit Sub
    End If
    
    If Not IsNumeric(GiftCardInput.Value) Then
        MsgBox "请输入有效的数字金额", vbExclamation, "无效输入"
        GiftCardInput.Value = ""
        Exit Sub
    End If
    
    Set wsGiftCards = ThisWorkbook.Sheets("GiftCards")
    Dim cardNo As Long: cardNo = Me.GiftCardNoInput
    Set cardRow = wsGiftCards.Range("A:A").Find(What:=cardNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not cardRow Is Nothing Then
        If cardRow.Offset(0, 2) < GiftCardInput.Value Then
            MsgBox "该礼品卡余额不足。"
            GiftCardInput.Value = ""
            Exit Sub
        End If
    End If
End Sub

Private Sub OtherAmountInput_AfterUpdate()
    If Not IsNumeric(OtherAmountInput.Value) Then
        MsgBox "请输入有效的数字金额", vbExclamation, "无效输入"
        OtherAmountInput.Value = ""
    End If
End Sub

Private Sub DateInput_AfterUpdate()
    Dim rawInput As String
    Dim parsedDate As Date
    rawInput = Trim(Me.DateInput.Value)

    On Error GoTo InvalidInput

    ' 能直接识别就格式化为 M/D/YYYY
    If IsDate(rawInput) Then
        parsedDate = CDate(rawInput)
        Me.DateInput.Value = Format(parsedDate, "m/d/yyyy")
        Exit Sub
    End If

    ' 尝试处理纯数字形式，如 20250503、05032025
    rawInput = Replace(rawInput, " ", "")
    If Len(rawInput) = 8 Then
        If IsNumeric(rawInput) Then
            Dim y As Integer, m As Integer, d As Integer

            ' yyyyMMdd
            y = Left(rawInput, 4)
            m = Mid(rawInput, 5, 2)
            d = Right(rawInput, 2)
            If IsDate(DateSerial(y, m, d)) Then
                Me.DateInput.Value = Format(DateSerial(y, m, d), "m/d/yyyy")
                Exit Sub
            End If

            ' MMddyyyy
            m = Left(rawInput, 2)
            d = Mid(rawInput, 3, 2)
            y = Right(rawInput, 4)
            If IsDate(DateSerial(y, m, d)) Then
                Me.DateInput.Value = Format(DateSerial(y, m, d), "m/d/yyyy")
                Exit Sub
            End If
        End If
    End If

InvalidInput:
    MsgBox "请输入合法日期，例如 5/3/2025、2025-05-03 或 05032025 等", vbExclamation
    Me.DateInput.Value = ""
End Sub

Private Sub PhoneInput_AfterUpdate()
    Dim phoneNumber As String
    Dim digitsOnly As String

    phoneNumber = Trim(PhoneInput.Value)

    ' 提取纯数字
    digitsOnly = ""
    Dim i As Integer
    For i = 1 To Len(phoneNumber)
        If Mid(phoneNumber, i, 1) Like "#" Then
            digitsOnly = digitsOnly & Mid(phoneNumber, i, 1)
        End If
    Next i

    ' 如果原始输入已经是格式化好的号码，就不动
    If phoneNumber Like "(###) ###-####" Then
        Exit Sub
    End If

    ' 如果纯数字为10位，格式化
    If Len(digitsOnly) = 10 Then
        PhoneInput.Value = FormatPhoneNumber(digitsOnly)
    Else
        MsgBox "请输入一个有效的10位电话号码！", vbCritical, "无效的电话号码"
        PhoneInput.Value = ""
    End If
End Sub

' 格式化电话号码为 (xxx) xxx-xxxx
Private Function FormatPhoneNumber(ByVal phone As String) As String
    ' 将电话号码格式化为 (xxx) xxx-xxxx
    FormatPhoneNumber = "(" & Mid(phone, 1, 3) & ") " & Mid(phone, 4, 3) & "-" & Mid(phone, 7, 4)
End Function

Private Sub SaveButton_Click()
    ' 提取 UserForm 数据
    Dim orderId As Long
    Dim scheduledDate As Date, scheduledTime As Variant
    Dim service As String, req As String, technician As String
    Dim phone As String, Name As String, comment As String
    Dim status As String, amount As String
    Dim cash As Double, POS As Double, giftCard As Double, otherAmount As Double
    Dim otherMethod As String
    Dim cardNo As Long
    
    ' 默认初始化为合理值
    orderId = SelectedOrderId
    cash = 0 ' 数值类型，默认为0
    POS = 0 ' 数值类型，默认为0
    giftCard = 0 ' 数值类型，默认为0
    otherAmount = 0 ' 数值类型，默认为0
    cardNo = 0 ' 数值类型，默认为0

    With Me
        ' 提取并处理用户输入数据
        Name = .NameInput.Value
    
        ' 提取并转换 scheduledDate 和 scheduledTime
        If .DateInput.Value <> "" Then
            scheduledDate = CDate(.DateInput.Value)
        End If
        
        If .TimeInput.Value <> "" Then
            scheduledTime = CDate(.TimeInput.Value)
            Debug.Print scheduledTime
        End If
    
        ' 提取其他文本字段
        req = .ReqSelection.Value
        technician = .TechnicianSelection.Value
        service = .ServiceSelection.Value
        comment = .CommentInput.Value
        status = .StatusSelection.Value
        phone = .PhoneInput.Value
    
        ' 直接转换金额字段，先判断是否为空
        If .AmountInput.Value <> "" Then
            amount = CStr(.AmountInput.Value)
        End If
        
        If .CashInput.Value <> "" Then
            cash = CDbl(.CashInput.Value)
        End If
        
        If .POSInput.Value <> "" Then
            POS = CDbl(.POSInput.Value)
        End If
        
        If .GiftCardInput.Value <> "" Then
            giftCard = CDbl(.GiftCardInput.Value)
        End If
        
        If .OtherAmountInput.Value <> "" Then
            otherAmount = CDbl(.OtherAmountInput.Value)
        End If
    
        otherMethod = .OtherSelection.Value
        
        If .GiftCardNoInput.Value <> "" Then
            cardNo = CLng(.GiftCardNoInput.Value)
        End If
    End With
    
    If status = "Paid" Then
        ' 更新 Orders 表
        Call UpdateOrderRowByIndex(orderId, , , , , , phone, Name, status, , comment)
        On Error GoTo Catch
    Else
        Dim totalPaid As Double: totalPaid = cash + POS + giftCard + otherAmount
        ' 如果各种支付方式的总金额大于应收款，则将支付状态修改为Paid
        If totalPaid >= amount Then
            status = "Paid"
        End If
        Call UpdateOrderRowByIndex(orderId, scheduledDate, scheduledTime, service, req, technician, phone, Name, status, totalPaid, comment)
        
        If cash <> 0 Then
            Call AddOrderPayment(orderId, "Cash", cash, 0)
        End If
        On Error GoTo Catch
        
        If POS <> 0 Then
            Call AddOrderPayment(orderId, "POS", POS, 0)
        End If
        On Error GoTo Catch
        
        If giftCard <> 0 Then
            Call AddOrderPayment(orderId, "Gift Card", giftCard, cardNo)
        End If
        On Error GoTo Catch
        
        If otherMethod <> "" Then
            Call AddOrderPayment(orderId, otherMethod, otherAmount, 0)
        End If
        On Error GoTo Catch
    End If

    ' 保存后关闭 UserForm
    Unload Me
    
    ' 更新图
    Call reDrawBlock(scheduledDate)
    
    Exit Sub
Catch:
    MsgBox "错误编号: " & Err.Number & vbCrLf & "错误描述: " & Err.Description, vbExclamation, "发生错误"
End Sub

Private Sub TimeInput_AfterUpdate()
    Dim rawInput As String
    Dim hourPart As String
    Dim minutePart As String
    Dim temp As String
    
    rawInput = Trim(Me.TimeInput.Value)
    rawInput = Replace(rawInput, " ", "") ' 去掉空格
    rawInput = Replace(rawInput, "：", ":") ' 中文冒号转英文
    rawInput = Replace(rawInput, ".", ":") ' 点转冒号
    
    ' 如果输入中有冒号
    If InStr(rawInput, ":") > 0 Then
        hourPart = Split(rawInput, ":")(0)
        minutePart = Split(rawInput, ":")(1)
    ElseIf Len(rawInput) = 4 Then
        ' 如 0930 或 2359
        hourPart = Left(rawInput, 2)
        minutePart = Right(rawInput, 2)
    ElseIf Len(rawInput) = 3 Then
        ' 如 930 → 9:30
        hourPart = Left(rawInput, 1)
        minutePart = Right(rawInput, 2)
    Else
        MsgBox "请输入有效的时间，如 0930、9:30 或 13:05", vbExclamation
        Me.TimeInput.Value = ""
        Exit Sub
    End If
    
    ' 验证合法性
    If IsNumeric(hourPart) And IsNumeric(minutePart) Then
        If CInt(hourPart) >= 0 And CInt(hourPart) <= 23 And CInt(minutePart) >= 0 And CInt(minutePart) <= 59 Then
            Me.TimeInput.Value = Format(TimeSerial(CInt(hourPart), CInt(minutePart), 0), "hh:mm")
            Exit Sub
        End If
    End If
    
    MsgBox "请输入有效的时间，小时应为 0-23，分钟为 0-59。", vbExclamation
    Me.TimeInput.Value = ""
End Sub


Private Sub UserForm_Initialize()
    ' 读取订单信息
    Dim wsOrders As Worksheet, wsPayments As Worksheet, wsLists As Worksheet, wsServices As Worksheet
    Set wsOrders = ThisWorkbook.Sheets("Orders")
    Set wsPayments = ThisWorkbook.Sheets("OrderPayments")
    Set wsLists = ThisWorkbook.Sheets("Lists")
    Set wsServices = ThisWorkbook.Sheets("Services")
    
    ' 加载订单信息（比如客户名、时间、应收金额等）
    Dim orderRow As Range
    Set orderRow = wsOrders.Range("A:A").Find(SelectedOrderId, , xlValues, xlWhole)
    If Not orderRow Is Nothing Then
        ' 加载字段
        Me.NameInput.Value = orderRow.Offset(0, 8).Value
        Me.PhoneInput.Value = orderRow.Offset(0, 7).Value
        Me.DateInput.Value = orderRow.Offset(0, 1).Value
        Me.TimeInput.Value = Format(orderRow.Offset(0, 2).Value, "hh:mm")
        ' 下拉表单从其他sheet读取
        LoadComboFromRange Me.TechnicianSelection, wsLists.Range("B1:B" & wsLists.Cells(wsLists.Rows.Count, "B").End(xlUp).row)
        LoadComboFromRange Me.StatusSelection, wsLists.Range("C1:C" & wsLists.Cells(wsLists.Rows.Count, "C").End(xlUp).row)
        LoadComboFromRange Me.ServiceSelection, wsServices.Range("A2:A" & wsServices.Cells(wsServices.Rows.Count, "A").End(xlUp).row)
        LoadComboFromRange Me.ReqSelection, wsLists.Range("A1:A" & wsLists.Cells(wsLists.Rows.Count, "A").End(xlUp).row)
        Me.ServiceSelection = orderRow.Offset(0, 4).Value
        Me.TechnicianSelection = orderRow.Offset(0, 6).Value
        Me.StatusSelection = orderRow.Offset(0, 9).Value
        Me.ReqSelection = orderRow.Offset(0, 5).Value
        Me.CommentInput = orderRow.Offset(0, 13).Value
        '根据Services表查找价格
        Dim price As Variant
        Dim i As Long
        For i = 2 To wsServices.Cells(wsServices.Rows.Count, 1).End(xlUp).row
            If wsServices.Cells(i, 1).Value = Me.ServiceSelection.Value Then
                price = wsServices.Cells(i, 3).Value
                Exit For
            End If
        Next i
        If price = 0 Then Exit Sub
        Me.AmountInput = price
    End If
    
    Dim isPaid As Boolean
    isPaid = (Me.StatusSelection.Value = "Paid")

    Call SetControlState(Me.CashInput, Not isPaid)
    Call SetControlState(Me.POSInput, Not isPaid)
    Call SetControlState(Me.GiftCardInput, Not isPaid)
    Call SetControlState(Me.OtherSelection, Not isPaid)
    Call SetControlState(Me.OtherAmountInput, Not isPaid)
    Call SetControlState(Me.GiftCardNoInput, Not isPaid)
    Call SetControlState(Me.AmountInput, Not isPaid)
    Call SetControlState(Me.DateInput, Not isPaid)
    Call SetControlState(Me.TimeInput, Not isPaid)
    Call SetControlState(Me.ServiceSelection, Not isPaid)
    Call SetControlState(Me.ReqSelection, Not isPaid)
    Call SetControlState(Me.TechnicianSelection, Not isPaid)
    
End Sub
Private Sub LoadComboFromRange(combo As MSForms.ComboBox, rng As Range)
    Dim cell As Range
    combo.Clear
    For Each cell In rng
        If Trim(cell.Value) <> "" And Trim(cell.Value) <> "Paid" Then combo.AddItem cell.Value
    Next cell
End Sub

Private Sub SetControlState(ctrl As MSForms.Control, enabled As Boolean)
    ctrl.enabled = enabled
    If enabled Then
        ctrl.BackColor = vbWhite
        ctrl.ControlTipText = ""
    Else
        ctrl.BackColor = RGB(240, 240, 240)
        ctrl.ControlTipText = "付过钱后不可修改此字段"
    End If
End Sub

