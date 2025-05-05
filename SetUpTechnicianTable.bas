Attribute VB_Name = "SetUpTechnicianTable"
' ���ڵ���ǰ��¼Ŀ�굥Ԫ��
Public targetCell As Range

Public Sub SetUpTechnicianTables()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsTechnicians As Worksheet, wsLeaveLog As Worksheet

    ' ɾ���ɱ�
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Technicians").Delete
    wb.Sheets("LeaveLog").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' �½���
    Set wsTechnicians = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsTechnicians.Name = "Technicians"
    Set wsLeaveLog = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): wsLeaveLog.Name = "LeaveLog"

    ' ���� Technicians ��ͷ
    With wsTechnicians
        .Range("A1:D1").Value = Array("Name", "Weekly Rest Days", "Status", "Remarks")
        .Range("A2:A6").Value = Application.WorksheetFunction.Transpose(Array("Jason", "David", "Lee", "Amy", "Cindy"))
        .Range("B2:B6").Value = Application.WorksheetFunction.Transpose(Array("Mon,Wed,Fri", "Tue,Thu", "Thu,Sat", "Mon,Fri", "Sun,Sat"))
        .Range("C2:C6").Value = Application.WorksheetFunction.Transpose(Array("On Duty", "On Leave", "On Leave", "On Duty", "On Duty"))
        .Range("D2:D6").Value = Application.WorksheetFunction.Transpose(Array("", "Cold", "Family Emergency", "", ""))
    End With

    ' ���� LeaveLog ��ͷ������
    With wsLeaveLog
        .Range("A1:D1").Value = Array("Name", "Start Date", "End Date", "Leave Reason")
        .Range("A2:A4").Value = Application.WorksheetFunction.Transpose(Array("Jason", "Lee", "Lee"))
        .Range("B2:B4").Value = Application.WorksheetFunction.Transpose(Array("2025-04-25", "2025-04-26", "2025-05-01"))
        .Range("C2:C4").Value = Application.WorksheetFunction.Transpose(Array("2025-04-25", "2025-04-26", "2025-05-01"))
        .Range("D2:D4").Value = Application.WorksheetFunction.Transpose(Array("Personal", "Family Emergency", "Family Emergency"))
        .Columns("B:C").NumberFormat = "yyyy-mm-dd"
        .Range("D2:D4").Validation.Add Type:=xlValidateList, Formula1:="Personal,Family Emergency,Illness"
    End With

    ' Technicians ��������֤
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

    ' ���ñ�ͷ��ʽ
    With wsTechnicians.Range("A1:D1")
        .Font.Color = RGB(255, 255, 255) ' ������ɫ
        .Interior.Color = RGB(128, 96, 0) ' ������ɫ
        .Font.Bold = True
    End With

    With wsLeaveLog.Range("A1:D1")
        .Font.Color = RGB(255, 255, 255) ' ������ɫ
        .Interior.Color = RGB(128, 96, 0) ' ������ɫ
        .Font.Bold = True
    End With

    ' ����ȫ��Ŀ�����Ϊ��
    Set targetCell = Nothing

    MsgBox "Technicians �� LeaveLog ����Ѵ�����˫����Ϣ�յ�Ԫ��ɵ�����ѡ���ڣ�", vbInformation
End Sub

' ���ڵ���ǰ��¼Ŀ�굥Ԫ��
Public Sub LaunchPickerForCell(ByVal Target As Range)
    ' �������ʼ�� targetCell ����
    Set targetCell = Target
    ShowWeekPicker
End Sub

Public Sub ShowWeekPicker()
    ' ȷ�� targetCell ����ȷ��ֵ
    If Not targetCell Is Nothing Then
        WeekPicker.SetTargetCell targetCell
        WeekPicker.Show
    End If
End Sub

