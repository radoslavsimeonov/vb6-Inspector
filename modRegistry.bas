Attribute VB_Name = "modRegistry"
Option Explicit

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

   

Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_READ = &H20000
Public Const STANDARD_RIGHTS_WRITE = &H20000
Public Const STANDARD_RIGHTS_EXECUTE = &H20000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000



Public Const REG_SZ        As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY    As Long = 3
Public Const REG_DWORD     As Long = 4
Public Const REG_MULTI_SZ  As Long = 7

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_SUCCESS = 0&
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Public Const BACKUPRESTORE = &H4

Public Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_WOW64_64KEY As Long = &H100&     '32 bit app to access 64 bit hive
Private Const KEY_WOW64_32KEY As Long = &H200&     '64 bit app to access 32 bit hive
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ32 = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_READ64 = ((KEY_WOW64_64KEY Or STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Const KEY_OPTIONS = (BACKUPRESTORE)
Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegQueryValueEx _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          lpData As Any, _
                                          lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          lpData As Long, _
                                          lpcbData As Long) As Long

Private Declare Function RegQueryValueExString _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          ByVal lpData As String, _
                                          lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          ByVal lpData As Long, _
                                          lpcbData As Long) As Long


Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegOpenKeyEx _
               Lib "advapi32.dll" _
               Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                      ByVal lpSubKey As String, _
                                      ByVal ulOptions As Long, _
                                      ByVal samDesired As Long, _
                                      phkResult As Long) As Long

Public Declare Function RegEnumKeyEx _
               Lib "advapi32.dll" _
               Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                                      ByVal dwIndex As Long, _
                                      ByVal lpName As String, _
                                      lpcbName As Long, _
                                      lpReserved As Long, _
                                      ByVal lpClass As String, _
                                      lpcbClass As Long, _
                                      lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegSetValueExString _
                Lib "advapi32.dll" _
                Alias "RegSetValueExA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal reserved As Long, _
                                        ByVal dwType As Long, _
                                        ByVal lpValue As String, _
                                        ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong _
                Lib "advapi32.dll" _
                Alias "RegSetValueExA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal reserved As Long, _
                                        ByVal dwType As Long, _
                                        lpValue As Long, _
                                        ByVal cbData As Long) As Long

Private Declare Function RegDeleteValue _
                Lib "advapi32.dll" _
                Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                         ByVal lpValueName As String) As Long
                                         
Private Declare Function RegDeleteKeyEx _
                Lib "advapi32.dll" _
                Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal samDesired As Long, _
                                       ByVal reserved As Long) As Long

Private Declare Function RegCreateKeyEx _
                Lib "advapi32.dll" _
                Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                         ByVal lpSubKey As String, _
                                         ByVal reserved As Long, _
                                         ByVal lpClass As String, _
                                         ByVal dwOptions As Long, _
                                         ByVal samDesired As Long, _
                                         ByVal lpSecurityAttributes As Long, _
                                         phkResult As Long, _
                                         lpdwDisposition As Long) As Long

Public Declare Function RegEnumValue _
                Lib "advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpValueName As String, _
                                       lpcbValueName As Long, _
                                       ByVal lpReserved As Long, _
                                       lpType As Long, _
                                       lpData As Any, _
                                       lpcbData As Long) As Long

Public Enum RegReadWOW64Constants
    x64 = KEY_WOW64_64KEY
    x86 = KEY_WOW64_32KEY
End Enum

Public Function QueryValue(ByVal hKey As Long, _
                           ByVal sKeyName As String, _
                           ByVal sValueName As Variant, _
                           Optional ByVal WOW64 As RegReadWOW64Constants = x64) As Variant

    Dim lRetval As Long
    Dim vValue  As Variant

    lRetval = RegOpenKeyEx(hKey, sKeyName, 0, KEY_QUERY_VALUE Or WOW64, hKey)
    lRetval = QueryValueEx(hKey, sValueName, vValue)
    QueryValue = vValue
    RegCloseKey (hKey)
End Function

Public Function QueryValueEx(ByVal lhKey As Long, _
                             ByVal szValueName As String, _
                             vValue As Variant) As Long

    Dim cch        As Long
    Dim lrc        As Long
    Dim lenght     As Long
    Dim lType      As Long
    Dim lValue     As Long
    Dim sValue     As String
    Dim aValue()   As Byte
    Dim arrValue() As String
    Dim resString  As String

    On Error GoTo QueryValueExError

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)

    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        Case REG_SZ, REG_EXPAND_SZ:
            sValue = String$(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)

            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        
        Case REG_BINARY:
            lenght = 2048
            ReDim aValue(0 To lenght - 1) As Byte
            lrc = RegQueryValueEx(lhKey, szValueName, 0&, lType, aValue(0), lenght)
            vValue = aValue
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)

            If lrc = ERROR_NONE Then vValue = lValue
        Case REG_MULTI_SZ
            lenght = 1024
            ReDim aValue(0 To lenght - 1) As Byte
            lrc = RegQueryValueEx(lhKey, szValueName, 0&, lType, aValue(0), lenght)
            resString = Space$(lenght - 2)
            CopyMemory ByVal resString, aValue(0), lenght - 2

            If Left$(resString, 1) = vbNullChar And Len(resString) > 1 Then resString = Mid$(resString, 2)
            vValue = Split(resString, vbNullChar) ' Replace$(resString, vbNullChar, ",")
        Case Else
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:

    Resume QueryValueExExit

End Function

Public Function SetKeyValue(lhKey As Long, _
                            sKeyName As String, _
                            sValueName As String, _
                            vValueSetting As Variant, _
                            lValueType As Long) As Long

    Dim lRetval As Long
    Dim hKey    As Long

    lRetval = RegOpenKeyEx(lhKey, sKeyName, 0, KEY_SET_VALUE, hKey)
    lRetval = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
    SetKeyValue = lRetval
End Function
   
Public Function SetValueEx(ByVal hKey As Long, _
                           sValueName As String, _
                           lType As Long, _
                           vValue As Variant) As Long

    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ, REG_MULTI_SZ
            sValue = vValue & vbNullChar
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        Case REG_EXPAND_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
    End Select

End Function

Public Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)

    Dim hNewKey As Long
    Dim lRetval As Long

    lRetval = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetval)
    RegCloseKey (hNewKey)
End Sub

Public Sub DeleteKeys(hType As Long, ByVal subKey As String, ByVal Key As String, lResult As Long, Optional ByVal WOW64 As RegReadWOW64Constants = x64)

    Dim hKey As Long

    If RegOpenKeyEx(hType, subKey, 0, KEY_ALL_ACCESS Or WOW64, hKey) = ERROR_NONE Then
        lResult = RegDeleteValue(hKey, Key)
        RegCloseKey hKey
    End If

End Sub
