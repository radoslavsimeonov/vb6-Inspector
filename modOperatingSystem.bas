Attribute VB_Name = "swOperatingSystem"
Option Explicit

Public Function GetOperatingSystem() As SoftwareOS
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim tOS As SoftwareOS

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_OperatingSystem", , 48)
    
    For Each objItem In colItems
        With tOS
            .Architecture = IIf(Is64bit, "x64", "x86")
            .BuildNumber = GetProp(objItem.BuildNumber)
            .Caption = GetProp(objItem.Caption)
            .CodeSet = GetProp(objItem.CodeSet)
            .CountryCode = GetProp(objItem.CountryCode)
            .CSDVersion = GetProp(objItem.CSDVersion)
            .CSName = GetProp(objItem.CSName)
            .Domain = GetDomainName
            .CurrentTimeZone = GetProp(objItem.CurrentTimeZone)
            .InstallDate = WMIDateStringToDate(GetProp(objItem.InstallDate))
            .Locale = GetProp(objItem.Locale)
            .Organization = GetProp(objItem.Organization)
            .OSLanugage = GetProp(objItem.OSLanguage)
            .Primary = GetProp(objItem.Primary)
            .ProductType = GetProp(objItem.ProductType)
            .RegisteredUser = GetProp(objItem.RegisteredUser)
            .SPMajorVersion = GetProp(objItem.ServicePackMajorVersion)
            .SPMinorVersion = GetProp(objItem.ServicePackMinorVersion)
            .SystemDirectory = GetProp(objItem.SystemDirectory)
            .SystemDrive = GetProp(objItem.SystemDrive)
            .Version = GetProp(objItem.Version)
            .WindowsDirecorty = GetProp(objItem.WindowsDirectory)
            .ActivationStatus = IsWindowsActivated(.Version)
        End With
    Next
    
    Set objWMIService = Nothing
    Set colItems = Nothing
    
    GetOperatingSystem = tOS
    
End Function

Public Function GetProductType(sType As String) As String
    
    Select Case sType
    
        Case "1": GetProductType = "Workstation"
        Case "2": GetProductType = "Domain controller"
        Case "3": GetProductType = "Server"
    
    End Select
    
End Function

Public Function GetWinKey() As String
On Error Resume Next

    Dim rpk
    Dim oReg
    Dim szPossibleChars, dwAccumulator, szProductKey
    Dim i, j As Integer
    Dim strRegKey As String
    
    strRegKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!//./root/default:StdRegProv")
    oReg.GetBinaryValue HKEY_LOCAL_MACHINE, strRegKey, "DigitalProductId", rpk

    If IsNull(rpk) Then Exit Function

    Const rpkOffset = 52: i = 28
    szPossibleChars = "BCDFGHJKMPQRTVWXY2346789"
    
    Do 'Rep1
    dwAccumulator = 0: j = 14
      Do
      dwAccumulator = dwAccumulator * 256
      dwAccumulator = rpk(j + rpkOffset) + dwAccumulator
      rpk(j + rpkOffset) = (dwAccumulator \ 24) And 255
      dwAccumulator = dwAccumulator Mod 24
      j = j - 1
      Loop While j >= 0
    i = i - 1: szProductKey = Mid$(szPossibleChars, dwAccumulator + 1, 1) & szProductKey
      If (((29 - i) Mod 6) = 0) And (i <> -1) Then
      i = i - 1: szProductKey = "-" & szProductKey
      End If
    Loop While i >= 0 'Goto Rep1
    
    GetWinKey = szProductKey

End Function

Private Function IsWindowsActivated(ByVal sVersion As String) As Integer
On Error Resume Next

    Dim oWMI, colItems, objItem
    Dim Result As Integer
            
    Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")
      
    Select Case sVersion
    
        Case Is < "5.2"
            Set colItems = oWMI.ExecQuery( _
                "SELECT IsNotificationOn,ActivationRequired FROM Win32_WindowsProductActivation")
    
            For Each objItem In colItems
                With objItem
                    IsWindowsActivated = _
                        IIf(GetProp(objItem.ActivationRequired) = 0, 1, 0)
                    
                    If GetProp(objItem.IsNotificationOn) <> 0 And _
                        GetProp(objItem.ActivationRequired) = 1 Then _
                            IsWindowsActivated = 5
                End With
            Next
            
        Case Is > "5.1"
    
            Set colItems = oWMI.ExecQuery( _
                "SELECT Description, LicenseStatus, GracePeriodRemaining FROM SoftwareLicensingProduct WHERE PartialProductKey <> null")
    
            For Each objItem In colItems

                With objItem
                    IsWindowsActivated = GetProp(objItem.LicenseStatus)
                End With

            Next

    End Select

End Function



Public Function GetOSVersion() As String
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim tOS As String

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_OperatingSystem", , 48)
    
    For Each objItem In colItems
        GetOSVersion = GetProp(objItem.Version)
    Next
    
    Set objWMIService = Nothing
    Set colItems = Nothing
    
End Function

