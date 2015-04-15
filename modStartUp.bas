Attribute VB_Name = "swStartUp"
Option Explicit


Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
    cFilePath As String
End Type

'File search
Private Declare Function FindFirstFile _
                Lib "KERNEL32" _
                Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                        lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile _
                Lib "KERNEL32" _
                Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                       lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes _
                Lib "KERNEL32" _
                Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "KERNEL32" (ByVal hFindFile As Long) As Long


Private Type AutorunInfo
    regKey          As Long
    KeyPath         As String
    KeyName         As String
    FriendlyName    As String
End Type

Private AutoRuns()  As AutorunInfo

Private ARs()       As SoftwareStartCommand
Private iAR         As Long

Private Const NOT_DISPLAIED = "desktop.ini sidebar ctfmon.exe"

Private Function FindStartupFiles(sPath As String, sFreindlyName As String)

    Dim FileName   As String ' Walking filename variable...
    Dim i          As Long ' For-loop counter...
    Dim hSearch    As Long ' Search Handle
    Dim WFD        As WIN32_FIND_DATA
    Dim Cont       As Long
    Dim cnt        As Long
        
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"

    hSearch = FindFirstFile(sPath & "*", WFD)
    Cont = True

    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = StripNulls(WFD.cFileName)

            If (FileName <> ".") And (FileName <> "..") Then
               
                If Left$(FileName, 2) <> "~$" Then
                    If InStr(NOT_DISPLAIED, LCase$(FileName)) = 0 Then
                        cnt = UBound(SW_START_COMMANDS) + 1
                        
                        ReDim Preserve ARs(iAR)
                        
                        With ARs(iAR)
                    
                            
                            If LCase$(GetFileExt(FileName)) = ".lnk" Then
                                .Command = GetLinkTarget(sPath & FileName)
                            Else
                                .Command = FileName
                            End If
                            
                            .CommandNameShort = FileName
                            .Command = Replace$(.Command, Chr$(34), "")
                            .Location = sPath
                            .UserRange = sFreindlyName
                            .Vendor = GetFileVendor2(.Command)
                            
                            If LCase$(GetFileExt(FileName)) = ".lnk" Then
                                .CommandName = FileName
                            Else
                                .CommandName = GetFileVendor2(.Command, FI_PRODUCTNAME)
                            End If
                            
                            '.CommandName = GetFileVendor2(.Command, FI_PRODUCTNAME)
                            'If .CommandName = vbNullString Then .CommandName = Replace$(FileName, ".lnk", "")
    
                            iAR = iAR + 1
                        
                        End With
                    End If
                End If
                                
            End If

            Cont = FindNextFile(hSearch, WFD) ' Get next file

            DoEvents
        Wend
        Cont = FindClose(hSearch)
    End If
        
End Function

Private Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr$(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Private Sub AddRegEntries(hKey As Long, sKeyPath As String, sKeyName As String, Optional sFriendlyName As String = vbNullString)
    Dim cnt As Integer
    Dim tAR As AutorunInfo
    
    cnt = UBound(AutoRuns)
    
    With tAR
        .regKey = hKey
        .KeyPath = sKeyPath
        .KeyName = sKeyName
        .FriendlyName = sFriendlyName
    End With
    
    AutoRuns(cnt) = tAR
    
    ReDim Preserve AutoRuns(cnt + 1)
End Sub

Private Sub FillAtuoRunsInfo()

    ReDim AutoRuns(0)
    
        AddRegEntries HKEY_LOCAL_MACHINE, _
                      "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", _
                      "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", _
                      "All users"

        AddRegEntries HKEY_LOCAL_MACHINE, _
                      "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce", _
                      "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce", _
                      "All users - Run once"

        AddRegEntries HKEY_LOCAL_MACHINE, _
                      "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnceEx", _
                      "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnceEx", _
                      "All users - Run once ex"

        AddRegEntries HKEY_CURRENT_USER, _
                      "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", _
                      "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", _
                      "Current user"

        AddRegEntries HKEY_CURRENT_USER, _
                      "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce", _
                      "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce", _
                      "Current user - Run once"

        AddRegEntries HKEY_LOCAL_MACHINE, _
                      "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Polices\\Explorer\\Run", _
                      "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Polices\\Explorer\\Run", _
                      "All users - Policy"

        AddRegEntries HKEY_CURRENT_USER, _
                      "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Polices\\Explorer\\Run", _
                      "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Polices\\Explorer\\Run", _
                      "Current user - Policy"

    ReDim Preserve AutoRuns(UBound(AutoRuns) - 1)
    
End Sub

Public Function EnumStartUpCommands() As SoftwareStartCommand()
    
    Dim i As Integer
    
    Call FillAtuoRunsInfo
    
    iAR = 0
    ReDim ARs(iAR)
    
    
    For i = 0 To UBound(AutoRuns)
        EnumStartUps AutoRuns(i)
        If Is64bit Then EnumStartUps AutoRuns(i), KEY_READ64
    Next
    
    
    
    FindStartupFiles GetSpecialfolder(CSIDL_COMMON_STARTUP), "Common Startup"
    FindStartupFiles GetSpecialfolder(CSIDL_STARTUP), "Startup"
    EnumStartUpCommands = ARs

End Function

Private Sub EnumStartUps(ari As AutorunInfo, Optional ByVal OSBits As Long = KEY_READ32)

    Dim KeyName     As String   ' receives name of each subkey
    Dim keylen      As Long     ' length of keyname
    Dim classname   As String   ' receives class of each subkey
    Dim classlen    As Long     ' length of classname
    Dim lastwrite   As FILETIME ' receives last-write-to time, but we ignore it here
    Dim hKey        As Long     ' handle to the HKEY_LOCAL_MACHINE\Software key
    Dim RetVal      As Long     ' function's return value
    Dim Index       As Long     ' counter variable for index
    
    Dim tmpArr()    As SoftwareStartCommand
    Dim cnt         As Long

    If ari.regKey = HKEY_CURRENT_USER And OSBits = KEY_READ64 Then Exit Sub

    RetVal = RegOpenKeyEx(ari.regKey, ari.KeyPath, 0, OSBits, hKey)

    If RetVal <> ERROR_SUCCESS Then
        Exit Sub
    End If

    Index = 0  ' initial index value

    While RetVal = 0
        KeyName = Space$(255)
        classname = Space$(255)
        keylen = 255
        classlen = 255

        RetVal = RegEnumValue(hKey, _
                                 Index, _
                                 ByVal KeyName, _
                                 keylen, _
                                 0&, _
                                 REG_DWORD, _
                                 ByVal classname, _
                                 classlen)

        If RetVal = ERROR_SUCCESS Then
            If keylen > 0 Then
                If InStr(NOT_DISPLAIED, LCase$(KeyName)) = 0 And _
                   InStr(NOT_DISPLAIED, LCase$(classname)) = 0 Then
                    
                    ReDim Preserve ARs(iAR)
                    
                    With ARs(iAR)
                        '.CommandName = Left$(KeyName, keylen)
                        .Architecture = OSBits
                        .Command = Replace$(Left$(classname, classlen), Chr$(34), "")
                        .Location = Replace$(ari.KeyName, "\\", "\")
                        .UserRange = ari.FriendlyName
                        .CommandNameShort = Left$(KeyName, keylen)
                        .CommandName = GetFileVendor2(.Command, FI_FILEDESCRIPTION)
                        If Trim$(.CommandName) = vbNullString Then _
                            .CommandName = GetFileVendor2(.Command, FI_PRODUCTNAME)
                        .Vendor = GetFileVendor2(.Command)
                    End With
                    
                    iAR = iAR + 1
                End If
            End If
        End If

        Index = Index + 1

    Wend
    RetVal = RegCloseKey(hKey)
    
    SW_START_COMMANDS = ARs
    Exit Sub
Err:
    MsgBox "Error " & Err.Description
End Sub

Public Sub DeleteStartupItem(tCom As SoftwareStartCommand)
    'On Error Resume Next
    
    Dim objFSO, objWMIService, objReg
    Dim strStartupName, strStartupLocation As String
    Dim strRoot               As String
    Dim strStartupUser        As String
    Dim strStartupFileUser    As String
    Dim arrStartupUser()      As String
    Dim arrCheckFile()        As String
    Dim strCheckFile          As String
    Dim strNetworkServicePath As String
    Dim intRegType            As Integer
    Dim booFile               As Boolean
    
    With tCom
        strStartupName = .CommandNameShort ' objItem.Caption
        strStartupUser = .UserRange 'objItem.User
        strStartupLocation = .Location ' objItem.Location
    End With
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    booFile = True
    
    Select Case tCom.Location

        Case "Common Startup"
            strStartupLocation = GetSpecialfolder(CSIDL_COMMON_STARTUP)
            
        Case "Startup"
            strStartupLocation = GetSpecialfolder(CSIDL_STARTUP)
    End Select
    
    If InStr(strStartupLocation, "HKU") > 0 Then
        strStartupLocation = Replace(strStartupLocation, "HKU\", "")
        booFile = False
        intRegType = 1
    ElseIf InStr(strStartupLocation, "HKLM") > 0 Then
        strStartupLocation = Replace(strStartupLocation, "HKLM\", "")
        booFile = False
        intRegType = 2
    ElseIf InStr(strStartupLocation, "HKCU") > 0 Then
        strStartupLocation = Replace(strStartupLocation, "HKCU\", "")
        booFile = False
        intRegType = 3
    End If
    
    If booFile = True Then
        If FExists(strStartupLocation & "\" & strStartupName) Then
            objFSO.DeleteFile (strStartupLocation & "\" & strStartupName)
        End If
    Else
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
        
        Select Case intRegType

            Case 1
                objReg.DeleteValue HKEY_USERS, strStartupLocation, strStartupName

            Case 2
                DeleteKeys HKEY_LOCAL_MACHINE, strStartupLocation, strStartupName, REG_SZ, tCom.Architecture
                'objReg.DeleteValue HKEY_LOCAL_MACHINE, strStartupLocation, strStartupName

            Case 3
                objReg.DeleteValue HKEY_CURRENT_USER, strStartupLocation, strStartupName
        End Select

    End If

End Sub
    
