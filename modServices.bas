Attribute VB_Name = "swServices"
Option Explicit

Public Function EnumServices() As SoftwareService()
On Error Resume Next

    Dim arrTemp() As SoftwareService
    Dim i         As Long
    Dim objWMIService, colServices, objService              As Object

    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2")
    Set colServices = objWMIService.ExecQuery("SELECT * FROM Win32_Service")
    Set objWMIService = Nothing
    ReDim arrTemp(0 To colServices.count)

    For Each objService In colServices

        With objService
            arrTemp(i).AcceptStop = GetProp(.AcceptStop)
            arrTemp(i).Caption = GetProp(GetProp(.Caption))
            arrTemp(i).Description = GetProp(.Description)
            arrTemp(i).Name = GetProp(.Name)
            arrTemp(i).PathName = GetProp(.PathName)
            arrTemp(i).StartMode = GetProp(.StartMode)
            arrTemp(i).state = GetProp(.state)
            If Not IsNull(.PathName) Then _
                arrTemp(i).Manufacturer = GetFileVendor2(.PathName)
        End With

        i = i + 1
    Next

    Set colServices = Nothing
    Set objService = Nothing
    EnumServices = arrTemp
End Function

Public Function ServiceMethods(action As String, SrvName As String) As Boolean
On Error Resume Next

    Dim strCondition As String
    Dim Counter      As Integer
    Dim objWMIService, objShare, objOutParams            As Object

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set objShare = objWMIService.Get("Win32_Service.Name='" & SrvName & "'")

    Select Case action
        Case "StopService", "StartService"
            Set objOutParams = objWMIService.ExecMethod("Win32_Service.Name='" & SrvName & "'", action)

            If action = "StopService" Then strCondition = "Stopped" Else strCondition = "Running"
            frmServices.ProgressBar1.Visible = True

            Do
                Set objShare = objWMIService.Get("Win32_Service.Name='" & SrvName & "'")
                frmServices.ProgressBar1.Value = Counter
                Counter = Counter + 1

                If Counter = frmServices.ProgressBar1.max - 10 Then
                    MsgBox "Операцията не приключи успешно, опитайте отново", vbCritical
                    Exit Do
                End If

                DoEvents
            Loop While objShare.state <> strCondition

            frmServices.ProgressBar1.Visible = False
            frmServices.ProgressBar1.Value = 1
        Case Else

            Dim objInParam

            Set objInParam = objShare.Methods_("ChangeStartMode").inParameters.SpawnInstance_()
            objInParam.Properties_.Item("StartMode") = action
            Set objOutParams = objWMIService.ExecMethod("Win32_Service.Name='" & SrvName & "'", "ChangeStartMode", objInParam)
    End Select

    Set objWMIService = Nothing
    Set objShare = Nothing
    Set objOutParams = Nothing
End Function
