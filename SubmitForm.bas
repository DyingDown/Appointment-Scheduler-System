Attribute VB_Name = "SubmitForm"
Sub SubmitOrderForm()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' ��ǰ��ҳ

    

    ' ��ȡ����������
    
    ' ��֤ʱ���Ƿ�Ϊ��
    Dim scheduledTime As Variant: scheduledTime = ws.Range("scheduledTime").Value
    Dim projectType As String: projectType = ws.Range("projectType").Value
    Dim technicianReq As String: technicianReq = ws.Range("technicianReq").Value
    Dim technician As String: technician = ws.Range("technician").Value
    Dim phone As String: phone = ws.Range("phone").Value
    Dim customerName As String: customerName = ws.Range("customerName").Value
    Dim comment As String: comment = ws.Range("comment").Value
    Dim formDate As Date: formDate = ws.Range("A1").Value

    ' ������֤
    If Not IsValidTime(scheduledTime) Then
        MsgBox "��������ȷ��ԤԼʱ�䣡", vbExclamation
    End If

    ' ��֤��Ŀ�����Ƿ�Ϊ��
    If IsEmpty(projectType) Or projectType = "" Then
        MsgBox "����д��Ŀ���ƣ�", vbExclamation
        Exit Sub
    End If

    ' ��֤�绰�����ʽ������е绰���룬�������֤��
    If phone <> "" Then
        If Not IsPhoneNumberValid(phone) Then
            MsgBox "�绰�����ʽ��Ч����������Ч�ĵ绰���롣", vbExclamation
            Exit Sub
        End If
    End If
    
    ' ���� technicianReq ��Ĭ��ֵ�����Ϊ�ջ�ȫ�ǿո�
    If Trim(technicianReq) = "" Then
        technicianReq = "Unspecified"  ' Ĭ��ֵ
    End If

    Call AddOrder(formDate, scheduledTime, projectType, technicianReq, technician, phone, customerName, comment)
    
End Sub

' ��֤�绰�����ʽ�����磺�����绰���룩
Public Function IsPhoneNumberValid(phone As String) As Boolean
    ' �򵥵ĵ绰�����ʽ��֤������������š��̺��ߡ��ո�
    Dim phoneRegex As Object
    Set phoneRegex = CreateObject("VBScript.RegExp")
    
    phoneRegex.IgnoreCase = True
    phoneRegex.Global = True
    phoneRegex.Pattern = "^\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}$"  ' ��ƥ���� (123) 456-7890 �� 123-456-7890 ��ʽ
    
    IsPhoneNumberValid = phoneRegex.Test(phone)
End Function

Public Function IsValidTime(inputValue As Variant) As Boolean
    On Error GoTo InvalidInput

    ' ����� Date ���ͻ���Ա�ʶ��Ϊʱ����ַ���
    If IsDate(inputValue) Then
        Dim t As Double
        t = CDbl(CDate(inputValue))
        
        ' ����Ƿ��� 0 �� 1 ֮�䣨�� 24 Сʱ�ڵ�ʱ�䣬�������ڲ��֣�
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

