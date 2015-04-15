Attribute VB_Name = "modHelpers"
Option Explicit

Private Declare Function GetLogicalDriveStrings _
                Lib "KERNEL32" _
                Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                 ByVal lpBuffer As String) As Long

Private Declare Function GetDriveType _
                Lib "KERNEL32" _
                Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Declare Function MakeSureDirectoryPathExists _
               Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Public Declare Function IsNTAdmin _
               Lib "advpack.dll" (ByVal dwReserved As Long, _
                                  ByRef lpdwReserved As Long) As Long

Private Declare Function GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                      nSize As Long) As Long

Private Declare Function GetProcAddress _
                Lib "KERNEL32" (ByVal hModule As Long, _
                                ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle _
                Lib "KERNEL32" _
                Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function IsWow64Process _
                Lib "KERNEL32" (ByVal hProc As Long, _
                                bWow64Process As Boolean) As Long

Private Declare Function GetCurrentProcess Lib "KERNEL32" () As Long

Private Declare Function OpenProcess _
                Lib "KERNEL32.dll" (ByVal dwDesiredAccess As Long, _
                                    ByVal bInheritHandle As Long, _
                                    ByVal dwProcessId As Long) As Long

Private Declare Function WaitForSingleObject _
                Lib "KERNEL32.dll" (ByVal hHandle As Long, _
                                    ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "KERNEL32.dll" (ByVal hObject As Long) As Long

' Special folder location
Const MAX_PATH = 260
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Const INFINITE          As Long = &HFFFFFFFF
Private Const SYNCHRONIZE       As Long = &H100000

Public Const REG_EX_IP_ADDRESS  As String = "((2[0-4]\d|25[0-5]|[01]?\d\d?)\.){3}(2[0-4]\d|25[0-5]|[01]?\d\d?)"
Public Const REG_EX_IP_ADDRESS2 As String = "^(\d|[1-9]\d|1\d\d|2([0-4]\d|5[0-5]))\.(\d|[1-9]\d|1\d\d|2([0-4]\d|5[0-5]))\.(\d|[1-9]\d|1\d\d|2([0-4]\d|5[0-5]))\.(\d|[1-9]\d|1\d\d|2([0-4]\d|5[0-5]))$"

Public Function ShellWait(ByVal FileName As String, ByVal windowstyle As VbAppWinStyle)

    On Error Resume Next

    Dim processId     As Long
    Dim ProcessHandle As Long

    processId = Shell(FileName, windowstyle)
    ProcessHandle = OpenProcess(SYNCHRONIZE, 0&, processId)
    WaitForSingleObject ProcessHandle, INFINITE
    CloseHandle ProcessHandle
End Function

Public Function DNSResolver(remote As String)

    Dim obj As clsNetTools

    Set obj = New clsNetTools

    If (RegExIsMatch(remote, REG_EX_IP_ADDRESS, vbNullString)) Then
        DNSResolver = obj.AddressToName(remote)
    Else
        DNSResolver = obj.NameToAddress(remote)
    End If

End Function
    
Public Function Is64bit() As Boolean

    Dim handle As Long, bolFunc As Boolean

    bolFunc = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc
End Function

Public Function GetProp(Value As Variant)
    On Error GoTo Err

    If IsNull(Value) Then

        Select Case VarType(Value)

            Case vbString
                GetProp = ""

            Case vbBoolean
                GetProp = False
        End Select

    Else
        GetProp = Value
    End If
    
    Exit Function
Err:
    GetProp = ""
End Function

Public Function CheckGUID(Value As String) As Boolean

    Const PatternGUID = "{########-####-####-####-############}"
    CheckGUID = UCase$(Value) Like Replace(PatternGUID, "#", "[0-9,A-F]")
End Function

Public Function RegExIsMatch(strString, strPattern, RetVal As String) As Boolean

    Dim RegEx        As RegExp
    Dim RegExMatches As MatchCollection
    Dim RegExEatch   As match

    RegExIsMatch = True
    RetVal = ""
    Set RegEx = New RegExp
    RegEx.IgnoreCase = True
    RegEx.Global = True
    RegEx.Pattern = strPattern

    If Not RegEx.Test(strString) Then
        RegExIsMatch = False
        Exit Function
    End If

    Set RegExMatches = RegEx.Execute(strString)
    RetVal = RegExMatches(0)
End Function

Public Function GetHWID(hardwareId As String) As HWID

    Dim tmpHWID As HWID
    Dim aType() As String

    If Len(Trim$(hardwareId)) = 0 Then Exit Function
    hardwareId = UCase$(Replace(hardwareId, vbNullChar, ""))
    aType = Split(hardwareId, "\")

    With tmpHWID
        .Type = aType(0)

        Select Case .Type

            Case "DISPLAY", "MONITOR", "ACPI"
                .VEN = aType(1)

                If InStr(.VEN, "*") > 0 Then
                    .VEN = Mid$(.VEN, InStr(.VEN, "*") + 1)
                ElseIf InStr(UCase$(.VEN), "DEV_") > 0 Then
                    .VEN = "PNP" & Mid$(.VEN, InStr(.VEN, "DEV_") + 4, 4)
                End If

                .VEN = Trim$(.VEN)

            Case Else

                If InStr(UCase$(hardwareId), "VEN_") > 0 Then
                    .VEN = Mid$(hardwareId, InStr(hardwareId, "VEN_") + 4, 4)
                ElseIf InStr(UCase$(hardwareId), "VID_") > 0 Then
                    .VEN = Mid$(hardwareId, InStr(hardwareId, "VID_") + 4, 4)
                ElseIf InStr(UCase$(hardwareId), "VID") > 0 Then
                    .VEN = Mid$(hardwareId, InStr(hardwareId, "VID") + 3, 4)
                End If

                If InStr(UCase$(hardwareId), "DEV_") > 0 Then
                    .dev = Mid$(hardwareId, InStr(hardwareId, "DEV_") + 4, 4)
                ElseIf InStr(UCase$(hardwareId), "PID_") > 0 Then
                    .dev = Mid$(hardwareId, InStr(hardwareId, "PID_") + 4, 4)
                ElseIf InStr(UCase$(hardwareId), "PID") > 0 Then
                    .dev = Mid$(hardwareId, InStr(hardwareId, "PID") + 3, 4)
                End If

                If InStr(UCase$(hardwareId), "REV_") > 0 Then
                    .REV = Mid$(hardwareId, InStr(hardwareId, "REV_") + 4, 2)
                End If

                If InStr(UCase$(hardwareId), "SUBSYS_") > 0 Then
                    .SUBSYS = Mid$(hardwareId, InStr(hardwareId, "SUBSYS_") + 7, 8)
                    .SubSys1 = Left$(.SUBSYS, 4)
                    .SubSys2 = Right$(.SUBSYS, 4)
                End If

        End Select

    End With

    GetHWID = tmpHWID
End Function

Public Function FExists(OrigFile As String) As Boolean

    Dim FS

    Set FS = CreateObject("Scripting.FileSystemObject")
    FExists = FS.FileExists(OrigFile)
End Function

Public Function DirExists(OrigFile As String)

    Dim FS

    Set FS = CreateObject("Scripting.FileSystemObject")
    DirExists = FS.FolderExists(OrigFile)
End Function

Public Function NotEmpty(sValue As String) As Boolean
    NotEmpty = False
    
    If Len(Trim$(sValue)) > 0 Then NotEmpty = True
End Function

Public Function ValidateTxt(ctrl As Control) As Boolean
    On Error Resume Next
    
    Dim ctrlText As String
    
    ValidateTxt = False
    
    ctrlText = ctrl.Text
    
    If NotEmpty(ctrlText) Then
        ValidateTxt = True
    Else
        MsgBox "Please enter " & LCase$(ctrl.Caption), vbExclamation, "Error"
        ctrl.SetFocus
    End If
    
End Function

Function HasElements(Arr) As Boolean
    On Error Resume Next
    
    HasElements = False
    HasElements = (UBound(Arr) > -1)
End Function

Public Function CurrentUser() As String
    Dim strBuff As String * 255
    Dim x       As Long
    CurrentUser = ""
    x = GetUserName(strBuff, Len(strBuff) - 1)

    If x > 0 Then
        x = InStr(strBuff, vbNullChar)

        If x > 0 Then
            CurrentUser = Left$(strBuff, x - 1)
        Else
            CurrentUser = Left$(strBuff, x)
        End If
    End If

End Function

'Round bytes in KB, MB, GB...
Public Function FormatBytes(ByVal num_bytes As Double, Optional iDig As Integer = -1) As String
    Const ONE_KB As Double = 1024
    Const ONE_MB As Double = ONE_KB * 1024
    Const ONE_GB As Double = ONE_MB * 1024
    Const ONE_TB As Double = ONE_GB * 1024
    Const ONE_YB As Double = ONE_TB * 1024
    
    Dim Value    As Double
    Dim txt      As String

    ' See how big the value is.
    If num_bytes <= 999 Then

        ' Format in bytes.
        FormatBytes = Format$(num_bytes, "0") & " bytes"
    ElseIf num_bytes <= ONE_KB * 999 Then

        ' Format in KB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_KB, iDig) & " " & "KB"
    ElseIf num_bytes <= ONE_MB * 999 Then

        ' Format in MB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_MB, iDig) & " " & "MB"
    ElseIf num_bytes <= ONE_GB * 999 Then

        ' Format in GB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_GB, iDig) & " " & "GB"
    ElseIf num_bytes <= ONE_TB * 999 Then

        ' Format in TB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_TB, iDig) & " " & "TB"
    Else

        ' Format in YB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_YB, iDig) & " " & "YB"
    End If

End Function

' Return the value formatted to include at most three
' non-zero digits and at most two digits after the
' decimal point. Examples:
'         1
'       123
'        12.3
'         1.23
'         0.12
Private Function ThreeNonZeroDigits(ByVal Value As Double, Optional iDig As Integer) As String

    If Value >= 100 Or iDig = 0 Then
        ' No digits after the decimal.
        ThreeNonZeroDigits = Format$(CInt(Value))
    ElseIf Value >= 10 Or iDig = 1 Then
        ' One digit after the decimal.
        ThreeNonZeroDigits = Format$(Value, "0.0")
    Else
        ' Two digits after the decimal.
        ThreeNonZeroDigits = Format$(Value, "0.00")
    End If

End Function

Public Function WMIDateStringToDate(dtmWMIDate) As String
    
    If Not IsNull(dtmWMIDate) Then
    
        WMIDateStringToDate = CDate(Mid$(dtmWMIDate, 5, 2) & "/" & Mid$(dtmWMIDate, 7, 2) & "/" & Left$(dtmWMIDate, 4) & " " & Mid$(dtmWMIDate, 9, 2) & ":" & Mid$(dtmWMIDate, 11, 2) & ":" & Mid$(dtmWMIDate, 13, 2))
    End If

End Function

Public Function GetDrives()
    
    Dim strSave    As String
    Dim strDrives  As String
    Dim strDrive() As String
    Dim ret        As Long
    Dim i          As Integer
    
    strSave = String$(255, Chr$(0)) 'Get all the drives
    ret = GetLogicalDriveStrings(255, strSave) 'Extract the drives from the buffer and print them on the form

    For i = 1 To 100

        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
        strDrives = strDrives & Left$(strSave, InStr(1, strSave, Chr$(0)) - 1) & ";"
        strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
    Next i
    
    strDrive = Split(strDrives, ";")
    
    For i = 0 To UBound(strDrive)

        If GetDriveType(strDrive(i)) = 3 And LCase$(strDrive(i)) <> "c:\" Then
            frmPolicy!cmbDrives.AddItem strDrive(i)
        End If

    Next i
    
    frmPolicy!cmbDrives.ListIndex = frmPolicy!cmbDrives.ListCount - 1
    
End Function

Public Function GetLinkTarget(sPath As String) As String
    On Error Resume Next

    Dim wshShell As Object
    Dim wshLink  As Object
    
    Set wshShell = CreateObject("WScript.Shell")
    Set wshLink = wshShell.CreateShortcut(sPath)
    
    GetLinkTarget = wshLink.TargetPath & " " & wshLink.Arguments
    GetLinkTarget = Trim$(GetLinkTarget)
    
    Set wshShell = Nothing
    Set wshLink = Nothing
    
End Function

Public Function GetFileExt(sFile As String) As String

    Dim i As Integer
    
    i = InStrRev(sFile, ".")

    If i Then
        GetFileExt = Mid$(sFile, i)
    Else
        GetFileExt = sFile
    End If

End Function

Public Function GetSpecialfolder(CSIDL As Long) As String
    
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Dim Path As String

    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = ERROR_SUCCESS Then

        Path = Space$(512)

        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)

        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Public Function GetFilePathWithOutParams(ByVal sPath As String) As String
    Dim strPath() As String
    Dim strPath2() As String
    Dim lngIndex As Long
    Dim lngIndex2 As Long
    Dim tPath As String
    
    If sPath = vbNullString Then Exit Function
    
    sPath = Replace$(sPath, "\\", "\")
    
    strPath() = Split(sPath, "\")  'Put the Parts of our path into an array
    lngIndex = UBound(strPath)
    tPath = strPath(lngIndex)  'Get the File Name from our array
    strPath2 = Split(tPath, " ")
    strPath(lngIndex) = strPath2(0)
    
    GetFilePathWithOutParams = Trim$(Join(strPath, "\"))
End Function

Public Function IsFormLoaded(ByVal frmName As String) As Boolean
    
    Dim frm As Form
    Dim FormCount As Integer
    
    IsFormLoaded = False
    
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = FormCount + 1
        End If
    Next
    
    If FormCount > 0 Then IsFormLoaded = True
End Function
