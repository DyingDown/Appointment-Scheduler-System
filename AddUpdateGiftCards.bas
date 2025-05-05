Attribute VB_Name = "AddUpdateGiftCards"
Public Sub AddGiftCard(balance As Double)

    On Error GoTo Catch

    If balance <= 0 Then
        Err.Raise vbObjectError + 2002, "AddGiftCard", "Balance ������� 0"
    End If

    ' ��ȡ GiftCards ������ͱ��
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("GiftCards")
    Set tbl = ws.ListObjects("GiftCardsTable")
    
    If tbl Is Nothing Then
        MsgBox "δ�ҵ� GiftCardsTable����ȷ�ϱ������", vbExclamation
        Exit Sub
    End If

    ' �������
    Dim newRow As listRow
    Set newRow = tbl.ListRows.Add
    
    With newRow
        .Range(1, 1).Value = GetNextIndexForSheet(ws.Name)
        .Range(1, 2).Value = balance
        .Range(1, 3).Value = "Active"
        .Range(1, 4).Value = Environ("Username")
        .Range(1, 5).Value = Now            ' Created Time �Զ�����
    End With

    MsgBox "Gift Card ��ӳɹ���", vbInformation
    Exit Sub

Catch:
    MsgBox "��� Gift Card ʧ�ܣ�" & Err.Description, vbCritical, "������� " & Err.Number
End Sub

Public Sub UpdateGiftCardByNumber(giftCardNo As Long, _
                                  paymentAmount As Double)

    On Error GoTo Catch

    If Trim(giftCardNo) = "" Then
        Err.Raise vbObjectError + 3001, "UpdateGiftCardByNumber", "Gift Card No ����Ϊ��"
    End If

    ' ��ȡ������ͱ��
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GiftCards")

    If tbl Is Nothing Then
        MsgBox "δ�ҵ� GiftCardsTable", vbExclamation
        Exit Sub
    End If

    Dim cardRow As Range
    Set cardRow = ws.Range("A:A").Find(What:=giftCardNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cardRow Is Nothing Then
        Err.Raise vbObjectError + 1000, , "�Ҳ�����Ʒ�����: " & giftCardNo
    End If

    Dim balance As Double
    Dim status As String
    Dim createdTime As Date
    
    balance = cardRow.Offset(0, 1).Value ' Balance �� B ��
    status = Trim(cardRow.Offset(0, 2).Value) ' Status �� C ��
    createdTime = cardRow.Offset(0, 3).Value ' CreatedTime �� D ��

    Dim rowIndex As Long
    rowIndex = foundCell.row - tbl.HeaderRowRange.row
    
    ' ����Ƿ����
    If Now >= DateAdd("yyyy", 1, createdTime) Then
        cardRow.Offset(0, 2).Value = "Expired"
         Err.Raise vbObjectError + 1001, , "��Ʒ���ѹ���"
    End If
    
    ' ���״̬
    If status <> "Active" Then
        Err.Raise vbObjectError + 1002, , "��Ʒ��״̬���Ϸ�: " & status
    End If
    
    ' ������
    If balance < paymentAmount Then
        Err.Raise vbObjectError + 1003, , "���㣬��ǰ���Ϊ " & balance
    End If
    
    ' Step 4: �������
    balance = balance - paymentAmount
    cardRow.Offset(0, 1).Value = balance

    ' Step 5: ������Ϊ 0�����Ϊ Used
    If balance = 0 Then
        cardRow.Offset(0, 2).Value = "Used"
    End If
    MsgBox "Gift Card ��Ϣ�Ѹ��£�", vbInformation
    Exit Sub

Catch:
    MsgBox "���� Gift Card ʧ�ܣ�" & Err.Description, vbCritical, "������� " & Err.Number
End Sub
