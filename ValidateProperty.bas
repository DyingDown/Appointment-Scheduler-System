Attribute VB_Name = "ValidateProperty"
Function IsValidTechnicianReq(valueToCheck As String) As Boolean
    IsValidTechnicianReq = IsValidValueFromNamedRange(valueToCheck, "TechnicianReqList")
End Function

Function IsValidTechnician(valueToCheck As String) As Boolean
    IsValidTechnician = IsValidValueFromNamedRange(valueToCheck, "TechnicianList")
End Function

Function IsValidStatus(valueToCheck As String) As Boolean
    IsValidStatus = IsValidValueFromNamedRange(valueToCheck, "StatusList")
End Function

Function IsValidPaymentMethod(valueToCheck As String) As Boolean
    IsValidPaymentMethod = IsValidValueFromNamedRange(valueToCheck, "PaymentMethodList")
End Function

Function IsValidProjectType(valueToCheck As String) As Boolean
    IsValidProjectType = IsValidValueFromNamedRange(valueToCheck, "ProjectTypeList")
End Function

Function IsValidCardStatus(valueToCheck As String) As Boolean
    IsValidCardStatus = IsValidValueFromNamedRange(valueToCheck, "CardStatusList")
End Function

Function IsValidOperation(valueToCheck As String) As Boolean
    IsValidOperation = IsValidValueFromNamedRange(valueToCheck, "OperationList")
End Function



' 通用校验方法，根据命名区域和待校验值判断是否有效
Function IsValidValueFromNamedRange(valueToCheck As String, rangeName As String) As Boolean
    Dim rng As Range
    Dim cell As Range
    
    On Error Resume Next
    Set rng = Range(rangeName)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "找不到命名区域：" & rangeName, vbCritical
        IsValidValueFromNamedRange = False
        Exit Function
    End If
    
    For Each cell In rng.Cells
        If StrComp(Trim(cell.Value), Trim(valueToCheck), vbTextCompare) = 0 Then
            IsValidValueFromNamedRange = True
            Exit Function
        End If
    Next cell
    
    IsValidValueFromNamedRange = False
End Function

