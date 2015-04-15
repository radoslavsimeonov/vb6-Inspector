Attribute VB_Name = "swSharedFolders"
Option Explicit

Public Function EnumSharedFolders() As SoftwareSharedFolder()
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim tmpArr() As SoftwareSharedFolder
    Dim idx      As Integer

    idx = 0
    ReDim tmpArr(idx)
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Share", , 48)

    For Each objItem In colItems
        ReDim Preserve tmpArr(idx)

        With tmpArr(idx)
            .ShareName = GetProp(objItem.Name)
            .Description = GetProp(objItem.Caption)
            .FolderPath = GetProp(objItem.Path)
            .MaximumAllowed = GetProp(objItem.MaximumAllowed)
            '.ShareType = GetProp(objItem.Type)
        End With

        idx = idx + 1
    Next

    Set objWMIService = Nothing
    Set colItems = Nothing
    
    EnumSharedFolders = tmpArr
End Function

Public Function ShareDelete(sShareName As String, ByRef sMessage As String) As Boolean
On Error Resume Next

    Dim objWMIService, objShare, objOutParams
    
    ShareDelete = False
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set objShare = objWMIService.Get("Win32_Share.Name='" & sShareName & "'")
    
    Set objOutParams = _
        objWMIService.ExecMethod("Win32_Share.Name='" & sShareName & "'", "Delete")
    
    ShareDelete = (objOutParams.ReturnValue = 0)
    
    sMessage = MethodMessage(objOutParams.ReturnValue)
    
    Set objWMIService = Nothing
    Set objShare = Nothing
    Set objOutParams = Nothing
    
End Function

Public Function ShareCreate(sShareName As String, _
                            sSharePath As String, _
                            iMaximumAllowed As Integer, _
                            sDescription As String, _
                            ByRef sMessage As String) As Boolean
On Error Resume Next

    Const FILE_SHARE = 0

    Dim objWMIService, objShare, objInParam, objOutParams
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set objShare = objWMIService.Get("Win32_Share")
    
    Set objInParam = objShare.Methods_("Create"). _
        inParameters.SpawnInstance_()
    
    objInParam.Properties_.Item("Path") = sSharePath
    objInParam.Properties_.Item("Name") = sShareName
    objInParam.Properties_.Item("Type") = FILE_SHARE
    objInParam.Properties_.Item("Description") = sDescription
    objInParam.Properties_.Item("MaximumAllowed") = iMaximumAllowed
    
    Set objOutParams = objWMIService.ExecMethod("Win32_Share", "Create", objInParam)
    
    ShareCreate = (objOutParams.ReturnValue = 0)
    
    sMessage = MethodMessage(objOutParams.ReturnValue)
    
    Set objWMIService = Nothing
    Set objShare = Nothing
    Set objInParam = Nothing
    Set objOutParams = Nothing
    
End Function

Public Function ShareModify(sShareName As String, _
                            iMaximumAllowed As Integer, _
                            sDecription As String, _
                            ByRef sMessage As String) As Boolean

On Error Resume Next

    Dim objWMIService, objShare, errReturn

    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    Set objShare = objWMIService.ExecQuery _
        ("Select * from Win32_Share Where Name = '" & sShareName & "'")
    
    For Each objShare In objShare
        errReturn = objShare.SetShareInfo(iMaximumAllowed, sDecription)
    Next
    
    ShareModify = (errReturn = 0)
    
    sMessage = MethodMessage(CInt(errReturn))
    
    Set objWMIService = Nothing
    Set objShare = Nothing
    
End Function

Public Function GetShareType(sType As String) As String
    
    Select Case Abs(val(sType))
        Case "0": GetShareType = "Disk Drive"
        Case "1": GetShareType = "Print Queue"
        Case "2": GetShareType = "Device"
        Case "3": GetShareType = "IPC"
        Case "2147483645": GetShareType = "IPC Admin"
        Case "2147483648": GetShareType = "Disk Drive Admin"
        Case "2147483649": GetShareType = "Print Queue Admin"
        Case "2147483650": GetShareType = "Device Admin"
        Case "2147483651": GetShareType = "IPC Admin"
    End Select

End Function

Public Function MethodMessage(iMsg As Integer) As String
    
    Select Case iMsg
        Case 0: MethodMessage = "Success"
        Case 2: MethodMessage = "Access Denied"
        Case 8: MethodMessage = "Unknown Failure"
        Case 9: MethodMessage = "Invalid Name"
        Case 10: MethodMessage = "Invalid Level"
        Case 21: MethodMessage = "Invalid Parameter"
        Case 22: MethodMessage = "Duplicate Share"
        Case 23: MethodMessage = "Redirected Path"
        Case 24: MethodMessage = "Unknown Device or Directory"
        Case 25: MethodMessage = "Net Name Not Found"
    End Select
    
End Function
