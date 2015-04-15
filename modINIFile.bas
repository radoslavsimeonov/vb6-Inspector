Attribute VB_Name = "modINIFile"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As String, _
   ByVal lpDefault As String, _
   ByVal lpReturnedString As String, _
   ByVal nSize As Long, _
   ByVal lpFileName As String) As Long
   
Public Declare Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As String, _
   ByVal lpValue As String, _
   ByVal lpFileName As String) As Long
   
Public Function sGetINI(sINIFile As String, sSection As String, sKey _
                As String, sDefault As String) As String
    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, _
              255, sINIFile)
    sGetINI = Left$(sTemp, nLength)
End Function

Public Sub writeINI(sINIFile As String, sSection As String, sKey _
           As String, sValue As String)
    Dim n As Integer
    Dim sTemp As String
    sTemp = sValue
    'Replace any CR/LF characters with spaces
    For n = 1 To Len(sValue)
        If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf _
        Then Mid$(sValue, n) = " "
    Next n
    n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
End Sub

