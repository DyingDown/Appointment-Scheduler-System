Attribute VB_Name = "DrawBlocks"
Public Sub reDrawCanvas()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' ��ǰ��ҳ
    Dim formDate As Date: formDate = ws.Range("A1").Value
    Call reDrawBlock(formDate)
End Sub
Public Sub reDrawBlock(targetDate As Date)
    ' ͳһ������ͼ����������/�޸ģ�
    Dim rowDict As Object: Set rowDict = CreateObject("Scripting.Dictionary")
    
    Debug.Print targetDate
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets("Orders").ListObjects("OrdersTable")
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Sub
    
    Dim tblRow As Range
    Dim rowStatus As String
    Dim rowDate As Variant
    
    For Each tblRow In tbl.DataBodyRange.Rows
        rowStatus = LCase(Trim(tblRow.Cells(1, 10).Value))  ' ��10����Status
        rowDate = tblRow.Cells(1, 2).Value '�ڶ���������
        ' Debug.Print "row added: " & " " & tblRow.row & " " & rowStatus & " " & rowDate
        If rowDate = targetDate And rowStatus <> "cancelled" Then
            If Not rowDict.exists(tblRow.row) Then
                rowDict.Add tblRow.row, True
            End If
        End If
    Next tblRow
    
    Call ClearBlocks(targetDate)
    
    Dim rowNum As Variant
    For Each rowNum In rowDict.Keys
        ' ȷ�� rowNum �ڱ��Χ��
        If rowNum >= tbl.DataBodyRange.row And rowNum < tbl.DataBodyRange.row + tbl.DataBodyRange.Rows.Count Then
            Call DrawBlock(rowNum)
        End If
    Next rowNum

End Sub


Sub ClearBlocks(ByVal targetDate As Date)
    Dim wsDaily As Worksheet
    Dim startCol As Long
    Dim endCol As Long
    Dim appointmentTimeCol As Long
    Dim targetRange As Range
    Dim currentRow As Long
    Dim col As Long
    
    ' ========== TODO: ��������ڣ��ͳ�ʼ��һ�� ================
    ' ��ȡ�������������Ű��"�Ű�_����"��
    Set wsDaily = ThisWorkbook.Sheets("�Ű�_" & Day(targetDate))
    
    Call ClearShapes(wsDaily)
    
    ' �ҵ�ԤԼʱ�����ڵ���
    appointmentTimeCol = 0 ' ��ʼ���к�
    For col = 1 To wsDaily.Columns.Count
        If wsDaily.Cells(1, col).Value = "ԤԼʱ��" Then ' ���� "ԤԼʱ��" �ڵ�3��
            appointmentTimeCol = col
            Exit For
        End If
    Next col
    
    If appointmentTimeCol = 0 Then
        MsgBox "�Ҳ���ԤԼʱ���У�", vbExclamation
        Exit Sub
    End If
    
    ' ���������Χ�Ľ����У�ԤԼʱ���� - 2��
    endCol = appointmentTimeCol - 2
    
    ' ����Ŀ�귶Χ����C6��Ŀ������У��з�Χ��6��69
    Set targetRange = wsDaily.Range(wsDaily.Cells(6, 3), wsDaily.Cells(69, endCol))
    
    ' ������ݺͱ�����ɫ
    targetRange.Clear ' ������ݡ�������ɫ�͸�ʽ
End Sub

Sub DrawBlock1(ByVal row As Long)
    Dim wsOrder As Worksheet: Set wsOrder = ThisWorkbook.Sheets("Orders")
    Dim orderDate As Variant: orderDate = wsOrder.Cells(row, 2).Value
    Dim scheduledTime As Variant: scheduledTime = wsOrder.Cells(row, 3).Value
    If IsEmpty(orderDate) Or IsEmpty(scheduledTime) Then Exit Sub

    Dim sheetName As String: sheetName = "�Ű�_" & Day(orderDate)
    Dim wsDaily As Worksheet
    On Error Resume Next
    Set wsDaily = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If wsDaily Is Nothing Then Exit Sub

    ' ��ȡ Services ����
    Dim projectType As String: projectType = wsOrder.Cells(row, 5).Value
    Dim duration As Long, shortName As String, price As Variant
    Dim wsServices As Worksheet: Set wsServices = ThisWorkbook.Sheets("Services")

    Dim i As Long
    For i = 2 To wsServices.Cells(wsServices.Rows.Count, 1).End(xlUp).row
        If wsServices.Cells(i, 1).Value = projectType Then
            duration = wsServices.Cells(i, 2).Value
            price = wsServices.Cells(i, 3).Value
            shortName = wsServices.Cells(i, 4).Value
            Exit For
        End If
    Next i
    If duration = 0 Then Exit Sub

    Dim blockHeight As Long: blockHeight = (duration + 9) \ 10

    ' ������ʼ��
    Dim scheduledHour As Long: scheduledHour = Hour(scheduledTime)
    Dim scheduledMinute As Long: scheduledMinute = Minute(scheduledTime)
    Dim startRow As Long: startRow = ((scheduledHour - 10) * 60 + scheduledMinute) \ 10 + 6

    ' �Ҽ�ʦ��������
    Dim techName As String: techName = wsOrder.Cells(row, 7).Value
    Dim col1 As Long, col2 As Long, found As Boolean
    For i = 3 To wsDaily.Cells(3, wsDaily.Columns.Count).End(xlToLeft).Column Step 3
        If wsDaily.Cells(3, i).Value = techName Then
            col1 = i: col2 = i + 1
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        MsgBox "δ�ҵ�ƥ���", vbExclamation
        Exit Sub
    End If

    Debug.Print col1 & " " & col2
    ' ����
    Dim j As Long, cellBlock As Range
    Dim status As String: status = wsOrder.Cells(row, 10).Value
    Dim isUnspecified As Boolean: isUnspecified = (wsOrder.Cells(row, 6).Value = "Unspecified")

    For j = 0 To blockHeight - 1
        Set cellBlock = wsDaily.Range(wsDaily.Cells(startRow + j, col1), wsDaily.Cells(startRow + j, col2))
        With cellBlock
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Name = "΢���ź�"
            .Borders.LineStyle = xlNone

            If status = "Arrived" Then
                If j = 0 Then .Interior.Color = RGB(182, 106, 108) Else .Interior.Color = RGB(255, 182, 193)
            ElseIf status = "Paid" Then
                If j = 0 Then .Interior.Color = RGB(58, 56, 56) Else .Interior.Color = RGB(117, 113, 113)
            ElseIf isUnspecified Then
                If j = 0 Then .Interior.Color = RGB(0, 84, 38) Else .Interior.Color = RGB(0, 130, 59)
            Else
                If j = 0 Then .Interior.Color = RGB(128, 96, 0) Else .Interior.Color = RGB(255, 190, 0)
            End If

            If j = 0 Then
                .Cells(1, 1).Value = wsOrder.Cells(row, 9).Value '���ϣ�����
                If status = "Paid" Then
                    .Cells(1, 2).Value = price '���ϣ��۸�
                End If
            ElseIf j = 1 Then
                .Cells(1, 1).Value = wsOrder.Cells(row, 8).Value '�绰
            ElseIf j = blockHeight - 1 Then
                .Cells(1, 2).Value = shortName '��д
            End If
        End With
    Next j
End Sub

Sub DrawBlock(ByVal row As Long)
    Dim wsOrder As Worksheet: Set wsOrder = ThisWorkbook.Sheets("Orders")
    Dim orderDate As Variant: orderDate = wsOrder.Cells(row, 2).Value
    Dim scheduledTime As Variant: scheduledTime = wsOrder.Cells(row, 3).Value
    If IsEmpty(orderDate) Or IsEmpty(scheduledTime) Then Exit Sub

    Dim sheetName As String: sheetName = "�Ű�_" & Day(orderDate)
    Dim wsDaily As Worksheet
    On Error Resume Next
    Set wsDaily = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If wsDaily Is Nothing Then Exit Sub

    ' ��ȡ Services ����
    Dim projectType As String: projectType = wsOrder.Cells(row, 5).Value
    Dim duration As Long, shortName As String, price As Variant
    Dim wsServices As Worksheet: Set wsServices = ThisWorkbook.Sheets("Services")

    Dim i As Long
    For i = 2 To wsServices.Cells(wsServices.Rows.Count, 1).End(xlUp).row
        If wsServices.Cells(i, 1).Value = projectType Then
            duration = wsServices.Cells(i, 2).Value
            price = wsServices.Cells(i, 3).Value
            shortName = wsServices.Cells(i, 4).Value
            Exit For
        End If
    Next i
    If duration = 0 Then Exit Sub

    Dim blockHeight As Long: blockHeight = (duration + 9) \ 10

    ' ������ʼ��
    Dim scheduledHour As Long: scheduledHour = Hour(scheduledTime)
    Dim scheduledMinute As Long: scheduledMinute = Minute(scheduledTime)
    Dim startRow As Long: startRow = ((scheduledHour - 10) * 60 + scheduledMinute) \ 10 + 6
    
    ' Debug.Print "row, startRow, blockHeight = " & row & " " & startRow & " " & blockHeight

    ' �Ҽ�ʦ��������
    Dim techName As String: techName = wsOrder.Cells(row, 7).Value
    Dim orderId As String: orderId = wsOrder.Cells(row, 1).Value
    Dim col1 As Long, col2 As Long, found As Boolean
    For i = 3 To wsDaily.Cells(3, wsDaily.Columns.Count).End(xlToLeft).Column Step 3
        If wsDaily.Cells(3, i).Value = techName Then
            col1 = i: col2 = i + 1
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        MsgBox "δ�ҵ�ƥ���", vbExclamation
        Exit Sub
    End If

    ' ���飨��״��
    Dim status As String: status = wsOrder.Cells(row, 10).Value
    Dim isUnspecified As Boolean: isUnspecified = (wsOrder.Cells(row, 6).Value = "Unspecified")
    Dim shape As shape
    Dim leftPosition As Double, topPosition As Double, width As Double, height As Double
    
    ' ���㵥Ԫ�������λ�ü����
    leftPosition = wsDaily.Cells(startRow, col1).Left
    topPosition = wsDaily.Cells(startRow, col1).Top
    width = wsDaily.Cells(startRow, col2).Left - leftPosition + wsDaily.Cells(startRow, col2).width
    height = blockHeight * wsDaily.Cells(startRow, col1).height

    Dim shapeList As New Collection
    Dim groupedShape As shape
    
    For j = 0 To blockHeight - 1
        If j = 0 Then
        
            HeaderColor = IIf(status = "Arrived", RGB(182, 106, 108), _
                  IIf(status = "Paid", RGB(58, 56, 56), _
                  IIf(isUnspecified, RGB(0, 84, 38), RGB(128, 96, 0))))
            ' ---- ��һ�У��ֱ��������� ----
            Set ShapeLeft = wsDaily.Shapes.AddShape(msoShapeRectangle, _
                wsDaily.Cells(startRow, col1).Left, topPosition, _
                wsDaily.Cells(startRow, col1).width, wsDaily.Cells(startRow, col1).height)
            With ShapeLeft
                .Fill.ForeColor.RGB = HeaderColor
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = HeaderColor
                .TextFrame2.TextRange.Text = wsOrder.Cells(row, 9).Value
                .TextFrame2.TextRange.Font.Name = "΢���ź�"
                .TextFrame2.VerticalAnchor = msoAnchorMiddle
                .TextFrame2.TextRange.Font.Size = 11
                .TextFrame2.TextRange.Font.Bold = True
                .TextFrame2.MarginLeft = 3
                .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
                .Line.Weight = 1
            End With
            ShapeLeft.Name = "idx_" & orderId & "_" & j & "_0"
            shapeList.Add ShapeLeft
    
            Set ShapeRight = wsDaily.Shapes.AddShape(msoShapeRectangle, _
                wsDaily.Cells(startRow, col2).Left, topPosition, _
                wsDaily.Cells(startRow, col2).width, wsDaily.Cells(startRow, col2).height)
            With ShapeRight
                .Fill.ForeColor.RGB = HeaderColor
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = HeaderColor
                .TextFrame2.TextRange.Font.Name = "΢���ź�"
                .TextFrame2.VerticalAnchor = msoAnchorMiddle
                .TextFrame2.TextRange.Font.Size = 11
                .TextFrame2.TextRange.Font.Bold = True
                .TextFrame2.MarginRight = 3
                .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
                .Line.Weight = 1
                If status = "Paid" Then
                    .TextFrame2.TextRange.Text = price
                    wsDaily.Cells(startRow, col2).Value = price
                End If
            End With
            ShapeRight.Name = "idx_" & orderId & "_" & j & "_1"
            shapeList.Add ShapeRight
    
        Else
            ' ---- �����У�һ���� ----
            Set shape = wsDaily.Shapes.AddShape(msoShapeRectangle, _
                leftPosition, topPosition + j * wsDaily.Cells(startRow, col1).height, _
                width, wsDaily.Cells(startRow, col1).height)
            BodyColor = IIf(status = "Arrived", RGB(255, 182, 193), _
                                      IIf(status = "Paid", RGB(117, 113, 113), _
                                      IIf(isUnspecified, RGB(0, 130, 59), RGB(255, 190, 0))))
            With shape
                .Fill.ForeColor.RGB = BodyColor
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = BodyColor
                .TextFrame2.TextRange.Font.Name = "΢���ź�"
                .TextFrame2.VerticalAnchor = msoAnchorMiddle
                .TextFrame2.TextRange.Font.Size = 11
                .TextFrame2.TextRange.Font.Bold = True
                .TextFrame2.MarginLeft = 3
                If j = 1 Then
                    .TextFrame2.TextRange.Text = wsOrder.Cells(row, 8).Value
                ElseIf j = 2 Then
                    .TextFrame2.TextRange.Text = wsOrder.Cells(row, 14).Value
                End If
    
                If j = blockHeight - 1 Then
                    .TextFrame2.TextRange.Text = shortName
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
                    .TextFrame2.MarginRight = 3
                Else
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
                End If
            End With
            shape.Name = "idx_" & orderId & "_" & j
            shapeList.Add shape
        End If
    Next j
    
    Dim shapesArray() As String
    
    ' ����һ����״���������
    ReDim shapesArray(shapeList.Count - 1)
    
    For i = 1 To shapeList.Count
        Set shape = shapeList.Item(i)
        shapesArray(i - 1) = shape.Name ' �洢���Ƶ�����
    Next i
    
    ' ʹ����״���������������״
    Set groupedShape = wsDaily.Shapes.Range(shapesArray).Group
    groupedShape.Name = "idx_" & orderId ' ��ϵ����֣����� order index
    groupedShape.OnAction = "ShowOrderFormFromShape" ' �󶨺�
End Sub

Sub ClearShapes(ws As Worksheet)
    Dim shape As shape
    On Error Resume Next ' ���Դ�����Ϊû����״ʱ�ᱨ��
    For Each shape In ws.Shapes
        If Left(shape.Name, 4) = "idx_" Then ' �ж���״�����Ƿ��� "idx_" ��ͷ
            shape.Delete ' ɾ��������������״
        End If
    Next shape
    On Error GoTo 0 ' �ָ������Ĵ�����
End Sub



