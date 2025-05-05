Attribute VB_Name = "SubmitForm"
Sub SubmitOrderForm()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' 当前表单页

    

    ' 获取表单输入数据
    
    ' 验证时间是否为空
    Dim scheduledTime As Variant: scheduledTime = ws.Range("scheduledTime").Value
    Dim projectType As String: projectType = ws.Range("projectType").Value
    Dim technicianReq As String: technicianReq = ws.Range("technicianReq").Value
    Dim technician As String: technician = ws.Range("technician").Value
    Dim phone As String: phone = ws.Range("phone").Value
    Dim customerName As String: customerName = ws.Range("customerName").Value
    Dim comment As String: comment = ws.Range("comment").Value
    Dim formDate As Date: formDate = ws.Range("A1").Value

    ' 数据验证
    If Not IsValidTime(scheduledTime) Then
        MsgBox "请输入正确的预约时间！", vbExclamation
    End If

    ' 验证项目名称是否为空
    If IsEmpty(projectType) Or projectType = "" Then
        MsgBox "请填写项目名称！", vbExclamation
        Exit Sub
    End If

    ' 验证电话号码格式（如果有电话号码，则进行验证）
    If phone <> "" Then
        If Not IsPhoneNumberValid(phone) Then
            MsgBox "电话号码格式无效！请输入有效的电话号码。", vbExclamation
            Exit Sub
        End If
    End If
    
    ' 设置 technicianReq 的默认值（如果为空或全是空格）
    If Trim(technicianReq) = "" Then
        technicianReq = "Unspecified"  ' 默认值
    End If

    Call AddOrder(formDate, scheduledTime, projectType, technicianReq, technician, phone, customerName, comment)
    
End Sub

' 验证电话号码格式（例如：美国电话号码）
Public Function IsPhoneNumberValid(phone As String) As Boolean
    ' 简单的电话号码格式验证，允许带有括号、短横线、空格
    Dim phoneRegex As Object
    Set phoneRegex = CreateObject("VBScript.RegExp")
    
    phoneRegex.IgnoreCase = True
    phoneRegex.Global = True
    phoneRegex.Pattern = "^\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}$"  ' 可匹配如 (123) 456-7890 或 123-456-7890 格式
    
    IsPhoneNumberValid = phoneRegex.Test(phone)
End Function

Public Function IsValidTime(inputValue As Variant) As Boolean
    On Error GoTo InvalidInput

    ' 如果是 Date 类型或可以被识别为时间的字符串
    If IsDate(inputValue) Then
        Dim t As Double
        t = CDbl(CDate(inputValue))
        
        ' 检查是否在 0 到 1 之间（即 24 小时内的时间，不含日期部分）
        If t >= 0 And t < 1 Then
            IsValidTime = True
        Else
            IsValidTime = False
        End If
        Exit Function
    End If

InvalidInput:
    IsValidTime = False
End Function

