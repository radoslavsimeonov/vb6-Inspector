Attribute VB_Name = "swLocalUsers"
Option Explicit


Private Const NERR_SUCCESS                            As Long = 0&
Private Const OPENUSERBROWSER_INCLUDE_SYSTEM          As Long = &H10000
Private Const OPENUSERBROWSER_SINGLE_SELECTION        As Long = &H1000&
Private Const OPENUSERBROWSER_NO_LOCAL_DOMAIN         As Long = &H100&
Private Const OPENUSERBROWSER_INCLUDE_CREATOR_OWNER   As Long = &H80&
Private Const OPENUSERBROWSER_INCLUDE_EVERYONE        As Long = &H40&
Private Const OPENUSERBROWSER_INCLUDE_INTERACTIVE     As Long = &H20&
Private Const OPENUSERBROWSER_INCLUDE_NETWORK         As Long = &H10&
Private Const OPENUSERBROWSER_INCLUDE_USERS           As Long = &H8&
Private Const OPENUSERBROWSER_INCLUDE_USER_BUTTONS    As Long = &H4&
Private Const OPENUSERBROWSER_INCLUDE_GROUPS          As Long = &H2&
Private Const OPENUSERBROWSER_INCLUDE_ALIASES         As Long = &H1&
Private Const OPENUSERBROWSER_FLAGS                   As Long = OPENUSERBROWSER_INCLUDE_USERS Or _
                                                                OPENUSERBROWSER_INCLUDE_USER_BUTTONS Or _
                                                                OPENUSERBROWSER_INCLUDE_EVERYONE Or _
                                                                OPENUSERBROWSER_INCLUDE_INTERACTIVE Or _
                                                                OPENUSERBROWSER_INCLUDE_NETWORK Or _
                                                                OPENUSERBROWSER_INCLUDE_ALIASES
Private Type OPENUSERBROWSER_STRUCT
   cbSize        As Long
   fCancelled    As Long
   Unknown       As Long
   hwndParent    As Long
   szTitle       As Long
   szDomainName  As Long
   dwFlags       As Long
   dwHelpID      As Long
   szHelpFile    As Long
End Type

Private Type ENUMUSERBROWSER_STRUCT
   SidType        As Long
   Sid1           As Long
   Sid2           As Long
   szFullName     As Long
   szUserName     As Long
   szDisplayName  As Long
   szDomainName   As Long
   szDescription  As Long
   sBuffer        As String * 1000
End Type

Private Declare Function OpenUserBrowser Lib "netui2.dll" _
  (lpOpenUserBrowser As Any) As Long
   
Private Declare Function EnumUserBrowserSelection Lib "netui2.dll" _
  (ByVal hBrowser As Long, _
   ByRef lpEnumUserBrowser As Any, _
   ByRef cbSize As Long) As Long
   
Private Declare Function CloseUserBrowser Lib "netui2.dll" _
   (ByVal hBrowser As Long) As Long
   
Private Declare Function lstrlenW Lib "KERNEL32" _
   (ByVal lpString As Long) As Long
   
Private Declare Sub CopyMemory Lib "KERNEL32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

   
Const ADS_UF_SCRIPT = &H1
Const ADS_UF_ACCOUNTDISABLE = &H2
Const ADS_UF_HOMEDIR_REQUIRED = &H8
Const ADS_UF_LOCKOUT = &H10
Const ADS_UF_PASSWD_NOTREQD = &H20
Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H80
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
Const ADS_UF_SMARTCARD_REQUIRED = &H40000
Const ADS_UF_PASSWORD_EXPIRED = &H800000
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_DELETE = 4

Public Function EnumAccounts() As SoftwareLocalUser()
On Error Resume Next
    
    Dim objComputer, objUser, objGroup, flag, objValues
    Dim arrTmp()    As SoftwareLocalUser
    Dim idx         As Long
    Dim gList       As String
    
    Set objComputer = GetObject("WinNT://.")
    objComputer.Filter = Array("User")
    For Each objUser In objComputer
        
        ReDim Preserve arrTmp(idx)

        arrTmp(idx).Group = "USER"
        For Each objGroup In objUser.Groups
            If LCase$(objGroup.Name) = "administrators" Then arrTmp(idx).Group = "ADMIN"
            If gList <> "" Then gList = gList & ","
            gList = gList & objGroup.Name
        Next
        
        Set objValues = GetObject("WinNT://./" & objUser.Name)
        
        arrTmp(idx).Groups = gList
        gList = ""
    
        flag = objValues.Get("UserFlags")
        
        If flag And ADS_UF_PASSWORD_EXPIRED Then
            arrTmp(idx).PasswordExpired = 1
        End If

        If flag And ADS_UF_DONT_EXPIRE_PASSWD Then
            arrTmp(idx).PasswordNeverExpires = 1
        End If

        If flag And ADS_UF_PASSWD_CANT_CHANGE Then
            arrTmp(idx).CannotChangePassword = 1
        End If

        If flag And ADS_UF_ACCOUNTDISABLE Then
            arrTmp(idx).AccountDisabled = 1
        End If

        If flag And ADS_UF_LOCKOUT Then
            arrTmp(idx).AccountLocked = 1
        End If

        If flag And ADS_UF_ACCOUNTDISABLE Then
            arrTmp(idx).AccountDisabled = 1
        End If
            
        If objUser.Get("LastLogin") <> "" Then
            arrTmp(idx).LastLogin = objUser.Get("LastLogin")
        End If
        
        arrTmp(idx).Index = idx
        arrTmp(idx).Name = objValues.Name
        arrTmp(idx).FullName = objValues.FullName
        arrTmp(idx).Description = objValues.Description
        arrTmp(idx).MaxPasswordLen = objValues.MinPasswordLength
        
        idx = idx + 1
    Next
    
    EnumAccounts = arrTmp
End Function

Public Function DeleteAccount(sUser As String) As Boolean
On Error Resume Next
    
    Dim objComputer
    
    DeleteAccount = True
    Set objComputer = GetObject("WinNT://.")
    objComputer.Delete "user", sUser
    If (Err.Number <> 0) Then
        DeleteAccount = False
    End If
End Function

Public Function SaveAccount(sUser As SoftwareLocalUser) As Boolean
On Error GoTo Err

    Dim objDomain, colAccounts, objUser, objUserFlags, objPasswordExpirationFlag, User, objGroup
    Dim AccountExist As Boolean
    Dim i As Integer
    Dim sGroups() As String
    Dim sComputer As String
    Dim gList As String
    
    SaveAccount = True
    
    sComputer = CreateObject("WScript.Network").ComputerName
    
    AccountExist = False
    Set objDomain = GetObject("WinNT://" & sComputer)
    objDomain.Filter = Array("user")
    For Each User In objDomain
        If LCase$(User.Name) = LCase$(sUser.Name) Then
            AccountExist = True
        End If
    Next
        
    If AccountExist = False Then
        Set objUser = objDomain.Create("user", sUser.Name)
        objUser.SetPassword sUser.Password
        objUser.SetInfo
    End If
    
    Set objUser = GetObject("WinNT://" & sComputer & "/" & sUser.Name & "")
    
    objUser.FullName = sUser.FullName & ""
    objUser.Description = sUser.Description & ""
    objUser.AccountDisabled = (sUser.AccountDisabled = 1)
    objUser.Put "PasswordExpired", CLng(sUser.PasswordExpired)
    
    objUserFlags = objUser.Get("UserFlags")
    
    objPasswordExpirationFlag = objUserFlags
    
    If sUser.PasswordNeverExpires = 1 Then
        objUserFlags = objUserFlags Or ADS_UF_DONT_EXPIRE_PASSWD
    Else
        objUserFlags = objUserFlags And (Not ADS_UF_DONT_EXPIRE_PASSWD)
    End If
    
    If sUser.CannotChangePassword = 1 Then
        objUserFlags = _
            objUserFlags Or ADS_UF_PASSWD_CANT_CHANGE
    Else
        objUserFlags = _
            objUserFlags And (Not ADS_UF_PASSWD_CANT_CHANGE)
    End If
           
    objUser.Put "userFlags", objUserFlags
    
    If Len(sUser.Password) > 0 Then
        objUser.SetPassword sUser.Password
    End If
            
    For Each objGroup In objUser.Groups
        Set objGroup = GetObject("WinNT://" & sComputer & "/" & objGroup.Name)
        objGroup.Remove ("WinNT://" & sComputer & "/" & sUser.Name)
    Next
    objUser.SetInfo
    
    sGroups = Split(sUser.Groups, ",")
    For i = 0 To UBound(sGroups)
        Set objGroup = objDomain.GetObject("group", sGroups(i))
        objGroup.Add ("WinNT://" & sComputer & "/" & sUser.Name)
    Next i
    Exit Function

Err:
    Debug.Print "Error while saving: " & CLng(Err.Number) & vbCrLf & Err.Description
    If Err.Number = -2147022651 Then MsgBox Err.Description: SaveAccount = False: Exit Function
    Resume Next
End Function

Public Function GetBrowserNames(ByVal hParent As Long, _
                                 ByVal sTitle As String, _
                                 sBuff As String) As Boolean

    Dim hBrowser   As Long
    Dim browser    As OPENUSERBROWSER_STRUCT
    Dim enumb      As ENUMUSERBROWSER_STRUCT
    Dim sComputer  As String
   
  
  sComputer = "\\" & CreateObject("WScript.Network").ComputerName
  
   With browser
      .cbSize = Len(browser)
      .fCancelled = 0
      .Unknown = 0
      .hwndParent = hParent
      .szTitle = StrPtr(sTitle)
      .szDomainName = StrPtr(sComputer)
      .dwFlags = OPENUSERBROWSER_INCLUDE_ALIASES Or OPENUSERBROWSER_SINGLE_SELECTION
   End With
       
End Function

Private Function GetPointerToByteStringW(ByVal dwData As Long) As String
  
   Dim tmp() As Byte
   Dim tmplen As Long
   
   If dwData <> 0 Then
   
      tmplen = lstrlenW(dwData) * 2
      
      If tmplen <> 0 Then
      
         ReDim tmp(0 To (tmplen - 1)) As Byte
         CopyMemory tmp(0), ByVal dwData, tmplen
         GetPointerToByteStringW = tmp
         
     End If
     
   End If
End Function

Public Function ComputerRename(sNewName As String) As Boolean
Dim objWMIService, objShare, objInParam, objOutParams

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set objShare = objWMIService.Get("Win32_ComputerSystem.Name='" & COMPUTER_NAME & "'")
    Set objInParam = objShare.Methods_("Rename"). _
        inParameters.SpawnInstance_()
       
    objInParam.Properties_.Item("Name") = sNewName

    Set objOutParams = objWMIService.ExecMethod("Win32_ComputerSystem.Name='" & COMPUTER_NAME & "'", "Rename", objInParam)
    
    Debug.Print "Out Parameters: "
    Debug.Print "ReturnValue: " & objOutParams.ReturnValue
End Function

