Attribute VB_Name = "SetUpTechnicianTable"
' 用于弹窗前记录目标单元格
Public targetCell As Range

Public Sub SetUpTechnicianTables()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsTechnicians As Worksheet, wsLeaveLog As Worksheet

    ' 删除旧表
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Technicians").Delete
    wb.Sheets("LeaveLog").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' 新建表
    Set wsTechnicians = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsTechnicians.Name = "Technicians"
    Set wsLeaveLog = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsLeaveLog.Name = "LeaveLog"

    ' 设置 Technicians 表头
    With wsTechnicians
        .Range("A1:D1").Value = Array("Name", "Weekly Rest Days", "Status", "Remarks")
        .Range("A2:A6").Value = Application.WorksheetFunction.Transpose(Array("Jason", "David", "Lee", "Amy", "Cindy"))
        .Range("B2:B6").Value = Application.WorksheetFunction.Transpose(Array("Mon,Wed,Fri", "Tue,Thu", "Thu,Sat", "Mon,Fri", "Sun,Sat"))
        .Range("C2:C6").Value = Application.WorksheetFunction.Transpose(Array("On Duty", "On Leave", "On Leave", "On Duty", "On Duty"))
        .Range("D2:D6").Value = Application.WorksheetFunction.Transpose(Array("", "Cold", "Family Emergency", "", ""))
    End With

    ' 设置 LeaveLog 表头和内容
    With wsLeaveLog
        .Range("A1:D1").Value = Array("Name", "Start Date", "End Date", "Leave Reason")
        .Range("A2:A4").Value = Application.WorksheetFunction.Transpose(Array("Jason", "Lee", "Lee"))
        .Range("B2:B4").Value = Application.WorksheetFunction.Transpose(Array("2025-04-25", "2025-04-26", "2025-05-01"))
        .Range("C2:C4").Value = Application.WorksheetFunction.Transpose(Array("2025-04-25", "2025-04-26", "2025-05-01"))
        .Range("D2:D4").Value = Application.WorksheetFunction.Transpose(Array("Personal", "Family Emergency", "Family Emergency"))
        .Columns("B:C").NumberFormat = "yyyy-mm-dd"
        .Range("D2:D4").Validation.Add Type:=xlValidateList, Formula1:="Personal,Family Emergency,Illness"
    End With

    ' Technicians 表数据验证
    With wsTechnicians
        .Range("B2:B6").Validation.Delete
        .Range("C2:C6").Validation.Delete

        With .Range("C2:C6").Validation
            .Add Type:=xlValidateList, Formula1:="On Duty,On Leave"
            .IgnoreBlank = True
            .ShowInput = True
            .ShowError = True
        End With
    End With

    ' 设置表头样式
    With wsTechnicians.Range("A1:D1")
        .Font.Color = RGB(255, 255, 255) ' 背景颜色
        .Interior.Color = RGB(128, 96, 0) ' 字体颜色
        .Font.Bold = True
    End With

    With wsLeaveLog.Range("A1:D1")
        .Font.Color = RGB(255, 255, 255) ' 背景颜色
        .Interior.Color = RGB(128, 96, 0) ' 字体颜色
        .Font.Bold = True
    End With

    ' 设置全局目标变量为空
    Set targetCell = Nothing

    MsgBox "Technicians 和 LeaveLog 表格已创建，双击休息日单元格可弹出多选窗口！", vbInformation
End Sub

' 用于弹窗前记录目标单元格
Public Sub LaunchPickerForCell(ByVal Target As Range)
    ' 在这里初始化 targetCell 变量
    Set targetCell = Target
    ShowWeekPicker
End Sub

Public Sub ShowWeekPicker()
    ' 确保 targetCell 已正确赋值
    If Not targetCell Is Nothing Then
        WeekPicker.SetTargetCell targetCell
        WeekPicker.Show
    End If
End Sub

