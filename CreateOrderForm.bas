Attribute VB_Name = "CreateOrderForm"
Sub CreateSimpleOrderForm(anchorCell As Range)
    Dim ws As Worksheet: Set ws = anchorCell.Worksheet

    ' ����һ�С����ۡ��ֶ�
    Dim labels As Variant, names As Variant
    labels = Array("ԤԼʱ��", "��Ŀ����", _
                   "��ʦҪ��", "��ʦ", _
                   "�绰����", "�ͻ�����", "����") ' �������
    names = Array("scheduledTime", "projectType", _
                  "technicianReq", "technician", _
                  "phone", "customerName", "comment") ' �������

    ' ���ò��֣�һ������
    Dim positions(0 To 6, 1 To 2)
    Dim i As Integer
    For i = 0 To 6 ' �������� 6 �޸�Ϊ 7
        positions(i, 1) = 0
        positions(i, 2) = i
    Next i

    ' ���ÿ�Ƭ������ʽ
    Dim cardRange As Range
    Set cardRange = ws.Range(anchorCell, anchorCell.Offset(1, 6)) ' �޸�Ϊ 6������һ��
    With cardRange
        .Font.Name = "΢���ź�"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255) ' �������б߿���ɫΪ��ɫ
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' ����������
    For i = 0 To UBound(labels)
        Dim rowOffset As Integer: rowOffset = positions(i, 1)
        Dim colOffset As Integer: colOffset = positions(i, 2)

        ' Label��
        With anchorCell.Offset(rowOffset, colOffset)
            .Value = labels(i)
            .Font.Bold = True
            .Interior.Color = RGB(37, 78, 120)
            .Font.Color = RGB(255, 255, 255)
        End With

        ' �����/������
        With anchorCell.Offset(rowOffset + 1, colOffset)
            .Name = names(i)
            .ClearContents
            .Interior.Color = RGB(155, 194, 230)
            .Font.Color = RGB(31, 56, 100)
            .Validation.Delete

            Select Case names(i)
                Case "scheduledTime"
                    .NumberFormat = "hh:mm"
                Case "projectType"
                    ' ��̬����ProjectType�������б�
                    .Validation.Add Type:=xlValidateList, _
                        Formula1:="=OFFSET(Services!$A$2,0,0,COUNTA(Services!$A:$A)-1,1)"
                Case "technicianReq"
                    ' ��̬����TechnicianReq�������б�
                    .Validation.Add Type:=xlValidateList, Formula1:="=TechnicianReqList"
                Case "technician"
                    ' ��̬����Technician�������б�
                    .Validation.Add Type:=xlValidateList, Formula1:="=TechnicianList"
                Case "phone"
                    .NumberFormat = "@"
                Case "comment"
                    ' �����в���Ҫ�����ʽ
            End Select
        End With
    Next i
    
    ws.Columns(anchorCell.Offset(0, 4).Column).ColumnWidth = 13.25 ' �����п�
    ws.Columns(anchorCell.Offset(0, 2).Column).ColumnWidth = 11
    ws.Columns(anchorCell.Offset(1, 1).Column).ColumnWidth = 14
    anchorCell.Offset(1, 1).HorizontalAlignment = xlLeft

    ' �ύ��ť
    Dim btnSubmit As shape
    Set btnSubmit = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchorCell.Offset(2, 0).Left, anchorCell.Offset(2, 0).Top, 60, 22)
    With btnSubmit
        .Name = "btnSubmitOrder"
        .TextFrame2.TextRange.Text = "�ύ"
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .Fill.ForeColor.RGB = RGB(0, 130, 59)
        .Line.Visible = msoFalse  ' ȥ���߿�
        .OnAction = "SubmitOrderForm"
    End With
    
    ' ��հ�ť
    Dim btnClear As shape
    Set btnClear = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchorCell.Offset(2, 2).Left, anchorCell.Offset(2, 2).Top, 60, 22)
    With btnClear
        .Name = "btnClearOrder"
        .TextFrame2.TextRange.Text = "���"
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .Fill.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Visible = msoFalse  ' ȥ���߿�
        .OnAction = "ClearOrderForm"
    End With
    
    ' ���»�ͼ��ť
    Dim btnDraw As shape
    Set btnDraw = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchorCell.Offset(2, 4).Left, anchorCell.Offset(2, 2).Top, 60, 22)
    With btnDraw
        .Name = "btnReDrawCanvas"
        .TextFrame2.TextRange.Text = "���»�ͼ"
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .Fill.ForeColor.RGB = RGB(255, 192, 0)
        .Line.Visible = msoFalse  ' ȥ���߿�
        .OnAction = "reDrawCanvas"
    End With

    ' �� Worksheet_Change �Զ���ʽ���绰����
    Call AttachPhoneFormatter(ws)

    ' ������������
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsServices As Worksheet: Set wsServices = wb.Sheets("Services")
    
    ' ��̬������������
    With wb.names
        ' ����Services���е���Ŀ����
        .Add Name:="TechnicianReqList", RefersTo:="=Lists!$A$1:$A$" & wsServices.Cells(wsServices.Rows.Count, "A").End(xlUp).row
        .Add Name:="TechnicianList", RefersTo:="=Lists!$B$1:$B$" & wsServices.Cells(wsServices.Rows.Count, "B").End(xlUp).row
    End With

    MsgBox "��������ɣ����������б��Ѿ���̬���£�", vbInformation
End Sub



Sub AttachPhoneFormatter(ws As Worksheet)
    Dim moduleCode As String
    moduleCode = _
    "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
    "    If Not Intersect(Target, Me.Range(""phone"")) Is Nothing Then" & vbCrLf & _
    "        Application.EnableEvents = False" & vbCrLf & _
    "        Dim raw As String: raw = Target.Value" & vbCrLf & _
    "        Dim digitsOnly As String, c As String, i As Integer: digitsOnly = """"" & vbCrLf & _
    "        For i = 1 To Len(raw)" & vbCrLf & _
    "            c = Mid(raw, i, 1)" & vbCrLf & _
    "            If c Like ""#"" Then digitsOnly = digitsOnly & c" & vbCrLf & _
    "        Next i" & vbCrLf & _
    "        If Len(digitsOnly) = 10 Then" & vbCrLf & _
    "            Target.Value = ""("" & Mid(digitsOnly, 1, 3) & "") "" & Mid(digitsOnly, 4, 3) & ""-"" & Mid(digitsOnly, 7, 4)" & vbCrLf & _
    "        End If" & vbCrLf & _
    "        Application.EnableEvents = True" & vbCrLf & _
    "    End If" & vbCrLf & _
    "End Sub"

    ' д���¼�����
    With ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, moduleCode
    End With
End Sub

