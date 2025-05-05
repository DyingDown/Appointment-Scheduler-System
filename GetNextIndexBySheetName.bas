Attribute VB_Name = "GetNextIndexBySheetName"
Function GetNextIndexForSheet(sheetName As String) As Long
    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Sheets("IndexStorage")

    Dim matchCell As Range
    Set matchCell = storeSheet.Columns("A").Find(What:=sheetName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not matchCell Is Nothing Then
        ' 找到记录，获取并递增
        Dim currentIndex As Long
        currentIndex = storeSheet.Cells(matchCell.row, 2).Value + 1
        storeSheet.Cells(matchCell.row, 2).Value = currentIndex
        GetNextIndexForSheet = currentIndex
    Else
        ' 没有记录，返回 0 或初始化处理
        MsgBox "未找到该工作表的 Index 记录！", vbExclamation
        GetNextIndexForSheet = 0
    End If
End Function


