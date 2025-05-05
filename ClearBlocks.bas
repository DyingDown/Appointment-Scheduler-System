Attribute VB_Name = "ClearBlocks"
Sub ClearBlocks(ByVal targetDate As Date)
    Dim wsDaily As Worksheet
    Dim startCol As Long
    Dim endCol As Long
    Dim appointmentTimeCol As Long
    Dim targetRange As Range
    Dim currentRow As Long
    Dim col As Long

    ' 获取工作表，假设是排班表（"排班_日期"）
    Set wsDaily = ThisWorkbook.Sheets("排班_" & Day(targetDate))
    
    ' 找到预约时间所在的列
    appointmentTimeCol = 0 ' 初始化列号
    For col = 1 To wsDaily.Columns.Count
        If wsDaily.Cells(1, col).Value = "预约时间" Then ' 假设 "预约时间" 在第3行
            appointmentTimeCol = col
            Exit For
        End If
    Next col
    
    If appointmentTimeCol = 0 Then
        MsgBox "找不到预约时间列！", vbExclamation
        Exit Sub
    End If
    
    ' 计算清除范围的结束列（预约时间列 - 2）
    endCol = appointmentTimeCol - 2
    
    ' 构建目标范围，从C6到目标结束列，行范围是6到69
    Set targetRange = wsDaily.Range(wsDaily.Cells(6, 3), wsDaily.Cells(69, endCol))
    
    ' 清除内容和背景颜色
    targetRange.Clear ' 清除内容、背景颜色和格式
End Sub



