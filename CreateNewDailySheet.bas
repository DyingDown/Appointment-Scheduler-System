Attribute VB_Name = "CreateNewDailySheet"
Sub GenerateNewScheduleSheet(ByVal staffList As Variant)

    Dim sheetName As String
    sheetName = "排班_" & Day(Date)
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' 如果已存在，删除并重新创建；否则新建
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' 禁止删除确认提示
        ws.Delete ' 删除已存在的工作表
        Application.DisplayAlerts = True ' 恢复删除提示
    End If
    
    ' 新建工作表
    Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = sheetName

    Dim staffCount As Integer
    staffCount = UBound(staffList) - LBound(staffList) + 1

    Dim totalCols As Integer
    totalCols = 2 + 3 * staffCount

    ' 设置字体为微软雅黑加粗
    With ws.Cells.Font
        .Name = "微软雅黑"
        .Bold = True
    End With

    ' 设置前 x 列列宽为 7
    Dim col As Integer
    For col = 1 To totalCols
        ws.Columns(col).ColumnWidth = 7
    Next col

    ' 日期标题（A1:C1）
    With ws.Range("A1:C1")
        .Merge
        .Value = Format(Date, "yyyy-mm-dd")
        .Font.Size = 20
        .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 32
    End With

    ' 星期几（D1）
    With ws.Range("D1")
        .Formula = "=TEXT(A1,""dddd"")"
        .Font.Size = 11
        .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With

    ' 营业额工资余额（倒数第5列）
    ws.Cells(1, totalCols - 4).Formula = "=""营业额:""&C2&""      工资:""&D2&""      余额:""&C2-D2"
    ws.Cells(1, totalCols - 4).Font.Size = 11
    ws.Cells(1, totalCols - 4).Font.Color = RGB(31, 56, 100)
    ws.Cells(1, totalCols - 4).HorizontalAlignment = xlLeft
    ws.Cells(1, totalCols - 4).VerticalAlignment = xlBottom ' 底部对齐

    ' 第二行：隐藏工作小时统计
    ws.Rows(2).RowHeight = 25.8
    ws.Range("C2:D2").Font.Color = RGB(255, 255, 255) ' 白色字体隐藏
    
    Dim sumFormula As String
    Dim colLetter As String
    
    sumFormula = "SUM("
    
    For i = 0 To staffCount - 1
        colLetter = Split(ws.Cells(4, 4 + i * 3).Address(True, False), "$")(0)
        If i > 0 Then sumFormula = sumFormula & ","
        sumFormula = sumFormula & colLetter & "4"
    Next i
    
    sumFormula = sumFormula & ")"
    
    ws.Range("D2").Formula = "=" & sumFormula
    ws.Range("C2").Formula = "=" & sumFormula & "*2"

    ' 第三行：Time和人名
    ws.Range("A3").Value = "Time"
    ws.Range("A3").Font.Color = RGB(31, 56, 100)
    ws.Range("A3").HorizontalAlignment = xlRight
    ws.Rows(3).RowHeight = 22.7
    ws.Rows(3).Font.Size = 14

    Dim k As Integer, colIndex As Integer
    colIndex = 3
    For k = 0 To staffCount - 1
        With ws.Range(ws.Cells(3, colIndex), ws.Cells(3, colIndex + 1))
            .Merge
            .Value = staffList(k)
            .Font.Color = RGB(31, 56, 100)
            .HorizontalAlignment = xlCenter
        End With
        colIndex = colIndex + 3
    Next k

    ' 第四行：当日工作小时统计
    ws.Rows(4).RowHeight = 12.7
    ws.Rows(4).Font.Color = RGB(31, 56, 100)
    colIndex = 4
    For i = 0 To staffCount - 1
        ws.Cells(4, colIndex).Formula = "=SUM(" & Cells(6, colIndex).Address & ":" & Cells(69, colIndex).Address & ")*0.5"
        colIndex = colIndex + 3
    Next i

    ' 第五行下边双线边框
    With ws.Range(ws.Cells(4, 1), ws.Cells(4, totalCols)).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Color = RGB(31, 56, 100)
        .Weight = xlThick
    End With

    ' 第五行行高
    ws.Rows(5).RowHeight = 16.5

        ' 第六行到第69行（10:00~20:30，每10分钟）
    Dim r As Long
    Dim t As Date: t = TimeValue("10:00")
    For r = 6 To 69
        ws.Rows(r).RowHeight = 16
        ws.Cells(r, 1).Value = Format(t, "hh:mm")
        ws.Cells(r, 1).Font.Size = 10
        ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
        If Minute(t) = 0 Or Minute(t) = 30 Then
            ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
        End If
        t = DateAdd("n", 10, t)
    Next r

    ' ==== 画虚线部分开始 ====
    Dim colEnd As String
    colEnd = Split(ws.Cells(1, 2 + staffCount * 3).Address(True, False), "$")(0)

    ' 删除旧的虚线
    Dim shp As shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 5) = "Line_" Then
            shp.Delete
        End If
    Next shp

    Dim topPos As Double, leftPos As Double, rightPos As Double

    For r = 6 To 69 Step 3
        topPos = ws.Cells(r, "B").Top
        leftPos = ws.Cells(r, "B").Left
        rightPos = ws.Cells(r, colEnd).Left + ws.Cells(r, colEnd).width

        Set shp = ws.Shapes.AddLine( _
            BeginX:=leftPos, BeginY:=topPos, _
            EndX:=rightPos, EndY:=topPos)

        With shp.Line
            .ForeColor.RGB = RGB(216, 216, 216)
            .Weight = 0.25
            .DashStyle = msoLineDash
        End With

        shp.Name = "Line_" & r
        shp.Placement = xlMove
    Next r
    ' ==== 画虚线部分结束 ====

    ' 冻结前四行
    ws.Activate
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True

    ' 表单从第1行、第(totalCol + 1)列开始放置
    Dim anchorCell As Range
    Set anchorCell = ws.Cells(1, totalCols + 2)

    ' 创建表单
    Call CreateSimpleOrderForm(anchorCell)
End Sub


