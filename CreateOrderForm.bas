Attribute VB_Name = "CreateOrderForm"
Sub CreateSimpleOrderForm(anchorCell As Range)
    Dim ws As Worksheet: Set ws = anchorCell.Worksheet

    ' 增加一列“评论”字段
    Dim labels As Variant, names As Variant
    labels = Array("预约时间", "项目类型", _
                   "技师要求", "技师", _
                   "电话号码", "客户姓名", "评论") ' 添加评论
    names = Array("scheduledTime", "projectType", _
                  "technicianReq", "technician", _
                  "phone", "customerName", "comment") ' 添加评论

    ' 设置布局：一行七列
    Dim positions(0 To 6, 1 To 2)
    Dim i As Integer
    For i = 0 To 6 ' 将列数从 6 修改为 7
        positions(i, 1) = 0
        positions(i, 2) = i
    Next i

    ' 设置卡片背景格式
    Dim cardRange As Range
    Set cardRange = ws.Range(anchorCell, anchorCell.Offset(1, 6)) ' 修改为 6，新增一列
    With cardRange
        .Font.Name = "微软雅黑"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255) ' 设置所有边框颜色为白色
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 创建表单内容
    For i = 0 To UBound(labels)
        Dim rowOffset As Integer: rowOffset = positions(i, 1)
        Dim colOffset As Integer: colOffset = positions(i, 2)

        ' Label行
        With anchorCell.Offset(rowOffset, colOffset)
            .Value = labels(i)
            .Font.Bold = True
            .Interior.Color = RGB(37, 78, 120)
            .Font.Color = RGB(255, 255, 255)
        End With

        ' 输入框/下拉框
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
                    ' 动态更新ProjectType的下拉列表
                    .Validation.Add Type:=xlValidateList, _
                        Formula1:="=OFFSET(Services!$A$2,0,0,COUNTA(Services!$A:$A)-1,1)"
                Case "technicianReq"
                    ' 动态更新TechnicianReq的下拉列表
                    .Validation.Add Type:=xlValidateList, Formula1:="=TechnicianReqList"
                Case "technician"
                    ' 动态更新Technician的下拉列表
                    .Validation.Add Type:=xlValidateList, Formula1:="=TechnicianList"
                Case "phone"
                    .NumberFormat = "@"
                Case "comment"
                    ' 评论列不需要特殊格式
            End Select
        End With
    Next i
    
    ws.Columns(anchorCell.Offset(0, 4).Column).ColumnWidth = 13.25 ' 更新列宽
    ws.Columns(anchorCell.Offset(0, 2).Column).ColumnWidth = 11
    ws.Columns(anchorCell.Offset(1, 1).Column).ColumnWidth = 14
    anchorCell.Offset(1, 1).HorizontalAlignment = xlLeft

    ' 提交按钮
    Dim btnSubmit As shape
    Set btnSubmit = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchorCell.Offset(2, 0).Left, anchorCell.Offset(2, 0).Top, 60, 22)
    With btnSubmit
        .Name = "btnSubmitOrder"
        .TextFrame2.TextRange.Text = "提交"
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .Fill.ForeColor.RGB = RGB(0, 130, 59)
        .Line.Visible = msoFalse  ' 去掉边框
        .OnAction = "SubmitOrderForm"
    End With
    
    ' 清空按钮
    Dim btnClear As shape
    Set btnClear = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchorCell.Offset(2, 2).Left, anchorCell.Offset(2, 2).Top, 60, 22)
    With btnClear
        .Name = "btnClearOrder"
        .TextFrame2.TextRange.Text = "清空"
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .Fill.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Visible = msoFalse  ' 去掉边框
        .OnAction = "ClearOrderForm"
    End With
    
    ' 重新画图按钮
    Dim btnDraw As shape
    Set btnDraw = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchorCell.Offset(2, 4).Left, anchorCell.Offset(2, 2).Top, 60, 22)
    With btnDraw
        .Name = "btnReDrawCanvas"
        .TextFrame2.TextRange.Text = "重新绘图"
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .Fill.ForeColor.RGB = RGB(255, 192, 0)
        .Line.Visible = msoFalse  ' 去掉边框
        .OnAction = "reDrawCanvas"
    End With

    ' 绑定 Worksheet_Change 自动格式化电话号码
    Call AttachPhoneFormatter(ws)

    ' 创建命名区域
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsServices As Worksheet: Set wsServices = wb.Sheets("Services")
    
    ' 动态更新命名区域
    With wb.names
        ' 更新Services表中的项目类型
        .Add Name:="TechnicianReqList", RefersTo:="=Lists!$A$1:$A$" & wsServices.Cells(wsServices.Rows.Count, "A").End(xlUp).row
        .Add Name:="TechnicianList", RefersTo:="=Lists!$B$1:$B$" & wsServices.Cells(wsServices.Rows.Count, "B").End(xlUp).row
    End With

    MsgBox "表单创建完成，并且下拉列表已经动态更新！", vbInformation
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

    ' 写入事件代码
    With ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, moduleCode
    End With
End Sub

