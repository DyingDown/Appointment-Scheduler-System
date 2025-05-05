Attribute VB_Name = "AddUpdateGiftCards"
Public Sub AddGiftCard(balance As Double)

    On Error GoTo Catch

    If balance <= 0 Then
        Err.Raise vbObjectError + 2002, "AddGiftCard", "Balance 必须大于 0"
    End If

    ' 获取 GiftCards 工作表和表格
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("GiftCards")
    Set tbl = ws.ListObjects("GiftCardsTable")
    
    If tbl Is Nothing Then
        MsgBox "未找到 GiftCardsTable，请确认表格名称", vbExclamation
        Exit Sub
    End If

    ' 添加新行
    Dim newRow As listRow
    Set newRow = tbl.ListRows.Add
    
    With newRow
        .Range(1, 1).Value = GetNextIndexForSheet(ws.Name)
        .Range(1, 2).Value = balance
        .Range(1, 3).Value = "Active"
        .Range(1, 4).Value = Environ("Username")
        .Range(1, 5).Value = Now            ' Created Time 自动生成
    End With

    MsgBox "Gift Card 添加成功！", vbInformation
    Exit Sub

Catch:
    MsgBox "添加 Gift Card 失败：" & Err.Description, vbCritical, "错误代码 " & Err.Number
End Sub

Public Sub UpdateGiftCardByNumber(giftCardNo As Long, _
                                  paymentAmount As Double)

    On Error GoTo Catch

    If Trim(giftCardNo) = "" Then
        Err.Raise vbObjectError + 3001, "UpdateGiftCardByNumber", "Gift Card No 不能为空"
    End If

    ' 获取工作表和表格
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GiftCards")

    If tbl Is Nothing Then
        MsgBox "未找到 GiftCardsTable", vbExclamation
        Exit Sub
    End If

    Dim cardRow As Range
    Set cardRow = ws.Range("A:A").Find(What:=giftCardNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cardRow Is Nothing Then
        Err.Raise vbObjectError + 1000, , "找不到礼品卡编号: " & giftCardNo
    End If

    Dim balance As Double
    Dim status As String
    Dim createdTime As Date
    
    balance = cardRow.Offset(0, 1).Value ' Balance 在 B 列
    status = Trim(cardRow.Offset(0, 2).Value) ' Status 在 C 列
    createdTime = cardRow.Offset(0, 3).Value ' CreatedTime 在 D 列

    Dim rowIndex As Long
    rowIndex = foundCell.row - tbl.HeaderRowRange.row
    
    ' 检查是否过期
    If Now >= DateAdd("yyyy", 1, createdTime) Then
        cardRow.Offset(0, 2).Value = "Expired"
         Err.Raise vbObjectError + 1001, , "礼品卡已过期"
    End If
    
    ' 检查状态
    If status <> "Active" Then
        Err.Raise vbObjectError + 1002, , "礼品卡状态不合法: " & status
    End If
    
    ' 检查余额
    If balance < paymentAmount Then
        Err.Raise vbObjectError + 1003, , "余额不足，当前余额为 " & balance
    End If
    
    ' Step 4: 更新余额
    balance = balance - paymentAmount
    cardRow.Offset(0, 1).Value = balance

    ' Step 5: 如果余额为 0，标记为 Used
    If balance = 0 Then
        cardRow.Offset(0, 2).Value = "Used"
    End If
    MsgBox "Gift Card 信息已更新！", vbInformation
    Exit Sub

Catch:
    MsgBox "更新 Gift Card 失败：" & Err.Description, vbCritical, "错误代码 " & Err.Number
End Sub
