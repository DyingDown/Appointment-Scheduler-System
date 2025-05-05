Attribute VB_Name = "ResetOrderForm"
Sub ClearOrderForm()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' 当前表单页

     ' 清空各个输入框的内容
    ws.Range("scheduledTime").Value = ""
    ws.Range("projectType").Value = ""
    ws.Range("technicianReq").Value = ""
    ws.Range("technician").Value = ""
    ws.Range("phone").Value = ""
    ws.Range("customerName").Value = ""
    ws.Range("comment").Value = ""
    
    MsgBox "表单已清空！", vbInformation
End Sub


