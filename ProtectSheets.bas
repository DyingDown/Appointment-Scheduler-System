Attribute VB_Name = "ProtectSheets"
'-------------------------------
' 将 SelectionChange 代码写入单个工作表
'-------------------------------
Sub InjectSelectionChange(ws As Worksheet)
    Dim moduleCode As String
    
    moduleCode = _
    "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "     ' 如果选中了表格下面的区域，直接选回表格" & vbCrLf & _
    "     If Target.row > 1 Then" & vbCrLf & _
    "        MsgBox ""请不要手动修改数据，使用 DailySheet 页面按钮进行新增或修改。"", vbExclamation" & vbCrLf & _
    "    End If" & vbCrLf & _
    "End Sub"
    
    With ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, moduleCode
    End With
End Sub

'-------------------------------
' 批量注入到多张工作表
'-------------------------------
Sub InjectToAllSheets()
    Dim sheetNames As Variant
    Dim i As Long
    
    ' 在这里列出你想要加上 SelectionChange 逻辑的工作表名
    sheetNames = Array("Orders", "OrderPayments", "GiftCards", "Logs")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        InjectSelectionChange ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
    Next i
    
    MsgBox "已成功注入 SelectionChange 事件到指定工作表！", vbInformation
End Sub

