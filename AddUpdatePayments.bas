Attribute VB_Name = "AddUpdatePayments"
Public Sub AddOrderPayment(orderId As Long, _
                           paymentMethod As String, amount As Double, _
                           giftCardNo As Long)

    On Error GoTo Catch
    
    If amount = 0 Then ' �����ǮΪ�㣬û��Ҫ��¼
        Exit Sub
    End If
    
    ' �������
    If orderId <= 0 Then Err.Raise vbObjectError + 4002, , "Order ID �Ƿ�" ' ===== TODO: �Ƿ���֤orderID���� =====
    If Trim(paymentMethod) = "" And IsValidPaymentMethod(paymentMethod) Then Err.Raise vbObjectError + 4003, , "֧����ʽΪ�ջ򲻺Ϸ�"
    If amount < 0 And paymentMethod <> "Cash" Then Err.Raise vbObjectError + 4004, , "���ֽ�֧����������� 0"
    If paymentMethod = "Gift Card" And giftCardNo = 0 Then
        Err.Raise vbObjectError + 4005, , "��Ʒ��֧�������ṩ����"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OrderPayments")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("OrderPaymentsTable")
    
    If tbl Is Nothing Then
        MsgBox "δ�ҵ� OrderPaymentsTable", vbExclamation
        Exit Sub
    End If

    ' �����
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
        .Cells(1, 9).Value = Environ("Username") ' ��ȡ��ǰ����ϵͳ��¼���û���
        .Cells(1, 10).Value = ""               ' DeletedBy
    End With
    
    If paymentMethod = "Gift Card" Then
        Call UpdateGiftCardByNumber(giftCardNo, amount)
    End If
    On Error GoTo Catch

    MsgBox "Payment ����ӣ�", vbInformation
    
    Exit Sub

Catch:
    MsgBox "��� Payment ʧ�ܣ�" & Err.Description, vbCritical, "������� " & Err.Number
End Sub


