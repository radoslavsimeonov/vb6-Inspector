Attribute VB_Name = "swApplications"
Option Explicit

Private apps() As SoftwareApp
Private idx    As Integer

Private Const UNINSTALL_PATH As String = "Software\Microsoft\Windows\CurrentVersion\Uninstall\"

Public Function EnumApplications() As SoftwareApp()
    idx = 0
    ReDim apps(idx)
    FindApplications
    FindApplications , HKEY_CURRENT_USER

    If Is64bit Then
        FindApplications x64
        FindApplications x64, HKEY_CURRENT_USER
    End If
    EnumApplications = apps
End Function

Private Sub FindApplications(Optional ByVal OSBits As RegReadWOW64Constants = x86, Optional reghkey As Long = HKEY_LOCAL_MACHINE)

    Dim KeyName     As String   ' receives name of each subkey
    Dim keylen      As Long     ' length of keyname
    Dim classname   As String   ' receives class of each subkey
    Dim classlen    As Long     ' length of classname
    Dim lastwrite   As FILETIME ' receives last-write-to time, but we ignore it here
    Dim hKey        As Long     ' handle to the HKEY_LOCAL_MACHINE\Software key
    Dim RetVal      As Long     ' function's return value
    Dim Index       As Long     ' counter variable for index
    Dim tmpAppName  As String
    Dim tmpIsUpdate As Boolean
    
    If OSBits = x64 And reghkey = HKEY_CURRENT_USER Then Exit Sub
    
    RetVal = RegOpenKeyEx(reghkey, "Software\Microsoft\Windows\CurrentVersion\Uninstall", 0, KEY_ENUMERATE_SUB_KEYS Or OSBits, hKey)

    If RetVal <> 0 Then
        Exit Sub
    End If

    Index = 0  ' initial index value

    While RetVal = 0
        KeyName = Space$(255)
        classname = Space$(255)
        keylen = 255
        classlen = 255
        RetVal = RegEnumKeyEx(hKey, Index, KeyName, keylen, ByVal 0, classname, classlen, lastwrite)

        If RetVal = 0 Then
            KeyName = Left$(KeyName, keylen)
            classname = Left$(classname, classlen)
            
            tmpAppName = QueryValue(reghkey, UNINSTALL_PATH & KeyName, "DisplayName", OSBits)
            
            tmpIsUpdate = Not ((QueryValue(reghkey, UNINSTALL_PATH & KeyName, "SystemComponent", OSBits) <> "1") _
                          And (InStr(LCase$(tmpAppName), "update for") = 0) _
                          And (InStr(LCase$(tmpAppName), "update rollup") = 0) _
                          And (InStr(LCase$(tmpAppName), "hotfix") = 0) _
                          And (IsKB(tmpAppName)))

            If Len(Trim$(tmpAppName)) > 0 And tmpIsUpdate = False Then
                ReDim Preserve apps(idx)
                apps(idx).AppName = tmpAppName
                apps(idx).Version = QueryValue(reghkey, UNINSTALL_PATH & KeyName, "DisplayVersion", OSBits)
                apps(idx).UninstallString = QueryValue(reghkey, UNINSTALL_PATH & KeyName, "UninstallString", OSBits)
                apps(idx).ModifyString = QueryValue(reghkey, UNINSTALL_PATH & KeyName, "ModifyPath", OSBits)
                apps(idx).AppBits = IIf(OSBits = x64, "x64", "x86")
                apps(idx).IsUpdate = tmpIsUpdate
                apps(idx).Publisher = QueryValue(reghkey, UNINSTALL_PATH & KeyName, "Publisher", OSBits)
                apps(idx).InstalledOn = ReverseDate(QueryValue(reghkey, UNINSTALL_PATH & KeyName, "InstallDate", OSBits))
                apps(idx).InstallLocation = QueryValue(reghkey, UNINSTALL_PATH & KeyName, "InstallLocation")
                idx = idx + 1
            End If
        End If

        Index = Index + 1

    Wend
    RetVal = RegCloseKey(hKey)
    Exit Sub
Err:
    MsgBox "Error " & Err.Description
End Sub

Private Function IsKB(ByVal Name As String) As Boolean

    Dim regExpr As RegExp
    Dim match   As match

    Set regExpr = New RegExp
    regExpr.IgnoreCase = True
    regExpr.Global = True
    regExpr.Pattern = "KB[0-9]{6}"
    IsKB = Not regExpr.Test(Name)
End Function

Private Function ReverseDate(sDate As String) As String

    Dim sDay, sMonth, sYear As String

    If Trim$(sDate) = "" Then Exit Function
    sDay = Right$(sDate, 2)
    sMonth = Mid$(sDate, 5, 2)
    sYear = Left$(sDate, 4)
    ReverseDate = sDay & "." & sMonth & "." & sYear
End Function
