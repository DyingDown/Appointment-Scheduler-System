Attribute VB_Name = "DeleteRowsEmptyID"
Sub DeleteRowsWithEmptyFirstColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' 设置目标工作表
    Set ws = ThisWorkbook.Sheets("Orders")
    
    ' 找到最后一行（避免整个表都遍历，提高性能）
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' 从最后一行开始往上遍历，避免删除行导致跳行
    For i = lastRow To 1 Step -1
        If Trim(ws.Cells(i, 1).Value) = "" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

