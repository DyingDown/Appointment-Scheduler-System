Attribute VB_Name = "ResetOrderForm"
Sub ClearOrderForm()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' ��ǰ��ҳ

     ' ��ո�������������
    ws.Range("scheduledTime").Value = ""
    ws.Range("projectType").Value = ""
    ws.Range("technicianReq").Value = ""
    ws.Range("technician").Value = ""
    ws.Range("phone").Value = ""
    ws.Range("customerName").Value = ""
    ws.Range("comment").Value = ""
    
    MsgBox "������գ�", vbInformation
End Sub


