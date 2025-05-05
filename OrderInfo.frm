VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderInfo 
   Caption         =   "OrderInfo"
   ClientHeight    =   8920.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8360.001
   OleObjectBlob   =   "OrderInfo.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "OrderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CashInput_AfterUpdate()
    If Not IsNumeric(CashInput.Value) Then
        MsgBox "��������Ч�����ֽ��", vbExclamation, "��Ч����"
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
            MsgBox "����Ʒ����ʧЧ�����Ϊ 0���޷�ʹ�á�"
            GiftCardInput.Value = ""
        End If
    Else
        MsgBox "���Ų�����"
        GiftCardInput.Value = ""
    End If
End Sub

Private Sub POSInput_AfterUpdate()
    If Not IsNumeric(POSInput.Value) Then
        MsgBox "��������Ч�����ֽ��", vbExclamation, "��Ч����"
        POSInput.Value = ""
    End If
End Sub

Private Sub GiftCardInput_AfterUpdate()
    If Me.GiftCardNoInput.Value = "" Then
        MsgBox "����������Ʒ���š�"
        GiftCardInput.Value = ""
        Exit Sub
    End If
    
    If Not IsNumeric(GiftCardInput.Value) Then
        MsgBox "��������Ч�����ֽ��", vbExclamation, "��Ч����"
        GiftCardInput.Value = ""
        Exit Sub
    End If
    
    Set wsGiftCards = ThisWorkbook.Sheets("GiftCards")
    Dim cardNo As Long: cardNo = Me.GiftCardNoInput
    Set cardRow = wsGiftCards.Range("A:A").Find(What:=cardNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not cardRow Is Nothing Then
        If cardRow.Offset(0, 2) < GiftCardInput.Value Then
            MsgBox "����Ʒ�����㡣"
            GiftCardInput.Value = ""
            Exit Sub
        End If
    End If
End Sub

Private Sub OtherAmountInput_AfterUpdate()
    If Not IsNumeric(OtherAmountInput.Value) Then
        MsgBox "��������Ч�����ֽ��", vbExclamation, "��Ч����"
        OtherAmountInput.Value = ""
    End If
End Sub

Private Sub DateInput_AfterUpdate()
    Dim rawInput As String
    Dim parsedDate As Date
    rawInput = Trim(Me.DateInput.Value)

    On Error GoTo InvalidInput

    ' ��ֱ��ʶ��͸�ʽ��Ϊ M/D/YYYY
    If IsDate(rawInput) Then
        parsedDate = CDate(rawInput)
        Me.DateInput.Value = Format(parsedDate, "m/d/yyyy")
        Exit Sub
    End If

    ' ���Դ���������ʽ���� 20250503��05032025
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
    MsgBox "������Ϸ����ڣ����� 5/3/2025��2025-05-03 �� 05032025 ��", vbExclamation
    Me.DateInput.Value = ""
End Sub

Private Sub PhoneInput_AfterUpdate()
    Dim phoneNumber As String
    Dim digitsOnly As String

    phoneNumber = Trim(PhoneInput.Value)

    ' ��ȡ������
    digitsOnly = ""
    Dim i As Integer
    For i = 1 To Len(phoneNumber)
        If Mid(phoneNumber, i, 1) Like "#" Then
            digitsOnly = digitsOnly & Mid(phoneNumber, i, 1)
        End If
    Next i

    ' ���ԭʼ�����Ѿ��Ǹ�ʽ���õĺ��룬�Ͳ���
    If phoneNumber Like "(###) ###-####" Then
        Exit Sub
    End If

    ' ���������Ϊ10λ����ʽ��
    If Len(digitsOnly) = 10 Then
        PhoneInput.Value = FormatPhoneNumber(digitsOnly)
    Else
        MsgBox "������һ����Ч��10λ�绰���룡", vbCritical, "��Ч�ĵ绰����"
        PhoneInput.Value = ""
    End If
End Sub

' ��ʽ���绰����Ϊ (xxx) xxx-xxxx
Private Function FormatPhoneNumber(ByVal phone As String) As String
    ' ���绰�����ʽ��Ϊ (xxx) xxx-xxxx
    FormatPhoneNumber = "(" & Mid(phone, 1, 3) & ") " & Mid(phone, 4, 3) & "-" & Mid(phone, 7, 4)
End Function

Private Sub SaveButton_Click()
    ' ��ȡ UserForm ����
    Dim orderId As Long
    Dim scheduledDate As Date, scheduledTime As Variant
    Dim service As String, req As String, technician As String
    Dim phone As String, Name As String, comment As String
    Dim status As String, amount As String
    Dim cash As Double, POS As Double, giftCard As Double, otherAmount As Double
    Dim otherMethod As String
    Dim cardNo As Long
    
    ' Ĭ�ϳ�ʼ��Ϊ����ֵ
    orderId = SelectedOrderId
    cash = 0 ' ��ֵ���ͣ�Ĭ��Ϊ0
    POS = 0 ' ��ֵ���ͣ�Ĭ��Ϊ0
    giftCard = 0 ' ��ֵ���ͣ�Ĭ��Ϊ0
    otherAmount = 0 ' ��ֵ���ͣ�Ĭ��Ϊ0
    cardNo = 0 ' ��ֵ���ͣ�Ĭ��Ϊ0

    With Me
        ' ��ȡ�������û���������
        Name = .NameInput.Value
    
        ' ��ȡ��ת�� scheduledDate �� scheduledTime
        If .DateInput.Value <> "" Then
            scheduledDate = CDate(.DateInput.Value)
        End If
        
        If .TimeInput.Value <> "" Then
            scheduledTime = CDate(.TimeInput.Value)
            Debug.Print scheduledTime
        End If
    
        ' ��ȡ�����ı��ֶ�
        req = .ReqSelection.Value
        technician = .TechnicianSelection.Value
        service = .ServiceSelection.Value
        comment = .CommentInput.Value
        status = .StatusSelection.Value
        phone = .PhoneInput.Value
    
        ' ֱ��ת������ֶΣ����ж��Ƿ�Ϊ��
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
        ' ���� Orders ��
        Call UpdateOrderRowByIndex(orderId, , , , , , phone, Name, status, , comment)
        On Error GoTo Catch
    Else
        Dim totalPaid As Double: totalPaid = cash + POS + giftCard + otherAmount
        ' �������֧����ʽ���ܽ�����Ӧ�տ��֧��״̬�޸�ΪPaid
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

    ' �����ر� UserForm
    Unload Me
    
    ' ����ͼ
    Call reDrawBlock(scheduledDate)
    
    Exit Sub
Catch:
    MsgBox "������: " & Err.Number & vbCrLf & "��������: " & Err.Description, vbExclamation, "��������"
End Sub

Private Sub TimeInput_AfterUpdate()
    Dim rawInput As String
    Dim hourPart As String
    Dim minutePart As String
    Dim temp As String
    
    rawInput = Trim(Me.TimeInput.Value)
    rawInput = Replace(rawInput, " ", "") ' ȥ���ո�
    rawInput = Replace(rawInput, "��", ":") ' ����ð��תӢ��
    rawInput = Replace(rawInput, ".", ":") ' ��תð��
    
    ' �����������ð��
    If InStr(rawInput, ":") > 0 Then
        hourPart = Split(rawInput, ":")(0)
        minutePart = Split(rawInput, ":")(1)
    ElseIf Len(rawInput) = 4 Then
        ' �� 0930 �� 2359
        hourPart = Left(rawInput, 2)
        minutePart = Right(rawInput, 2)
    ElseIf Len(rawInput) = 3 Then
        ' �� 930 �� 9:30
        hourPart = Left(rawInput, 1)
        minutePart = Right(rawInput, 2)
    Else
        MsgBox "��������Ч��ʱ�䣬�� 0930��9:30 �� 13:05", vbExclamation
        Me.TimeInput.Value = ""
        Exit Sub
    End If
    
    ' ��֤�Ϸ���
    If IsNumeric(hourPart) And IsNumeric(minutePart) Then
        If CInt(hourPart) >= 0 And CInt(hourPart) <= 23 And CInt(minutePart) >= 0 And CInt(minutePart) <= 59 Then
            Me.TimeInput.Value = Format(TimeSerial(CInt(hourPart), CInt(minutePart), 0), "hh:mm")
            Exit Sub
        End If
    End If
    
    MsgBox "��������Ч��ʱ�䣬СʱӦΪ 0-23������Ϊ 0-59��", vbExclamation
    Me.TimeInput.Value = ""
End Sub


Private Sub UserForm_Initialize()
    ' ��ȡ������Ϣ
    Dim wsOrders As Worksheet, wsPayments As Worksheet, wsLists As Worksheet, wsServices As Worksheet
    Set wsOrders = ThisWorkbook.Sheets("Orders")
    Set wsPayments = ThisWorkbook.Sheets("OrderPayments")
    Set wsLists = ThisWorkbook.Sheets("Lists")
    Set wsServices = ThisWorkbook.Sheets("Services")
    
    ' ���ض�����Ϣ������ͻ�����ʱ�䡢Ӧ�ս��ȣ�
    Dim orderRow As Range
    Set orderRow = wsOrders.Range("A:A").Find(SelectedOrderId, , xlValues, xlWhole)
    If Not orderRow Is Nothing Then
        ' �����ֶ�
        Me.NameInput.Value = orderRow.Offset(0, 8).Value
        Me.PhoneInput.Value = orderRow.Offset(0, 7).Value
        Me.DateInput.Value = orderRow.Offset(0, 1).Value
        Me.TimeInput.Value = Format(orderRow.Offset(0, 2).Value, "hh:mm")
        ' ������������sheet��ȡ
        LoadComboFromRange Me.TechnicianSelection, wsLists.Range("B1:B" & wsLists.Cells(wsLists.Rows.Count, "B").End(xlUp).row)
        LoadComboFromRange Me.StatusSelection, wsLists.Range("C1:C" & wsLists.Cells(wsLists.Rows.Count, "C").End(xlUp).row)
        LoadComboFromRange Me.ServiceSelection, wsServices.Range("A2:A" & wsServices.Cells(wsServices.Rows.Count, "A").End(xlUp).row)
        LoadComboFromRange Me.ReqSelection, wsLists.Range("A1:A" & wsLists.Cells(wsLists.Rows.Count, "A").End(xlUp).row)
        Me.ServiceSelection = orderRow.Offset(0, 4).Value
        Me.TechnicianSelection = orderRow.Offset(0, 6).Value
        Me.StatusSelection = orderRow.Offset(0, 9).Value
        Me.ReqSelection = orderRow.Offset(0, 5).Value
        Me.CommentInput = orderRow.Offset(0, 13).Value
        '����Services����Ҽ۸�
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
        ctrl.ControlTipText = "����Ǯ�󲻿��޸Ĵ��ֶ�"
    End If
End Sub

