Attribute VB_Name = "mErr"
Option Explicit

Private Const NERR_BASE = 2100
Private Const MAX_NERR = (NERR_BASE + 899)
Private Const DLL_NETMSG = "C:\WINDOWS\SYSTEM32\NETMSG.DLL"

Private Const LOAD_LIBRARY_AS_DATAFILE = &H2

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const LANG_USER_DEFAULT = &H400&

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                                    (ByVal dwFlags As Long, _
                                    lpSource As Any, _
                                    ByVal dwMessageId As Long, _
                                    ByVal dwLanguageId As Long, _
                                    ByVal lpBuffer As String, _
                                    ByVal nSize As Long, _
                                    Arguments As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
 
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" ( _
                                    ByVal lpLibFileName As String, _
                                    ByVal hFile As Long, _
                                    ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Function fncGetErrorString(Optional ByVal lngError As Long = 0) As String
    Dim strMsg As String * 256
    Dim lngRet As Long
    Dim lngFlags As Long
    Dim hModule As Long
    
    lngFlags = FORMAT_MESSAGE_FROM_SYSTEM Or _
               FORMAT_MESSAGE_IGNORE_INSERTS
    'lngFlags = FORMAT_MESSAGE_FROM_SYSTEM Or _
               FORMAT_MESSAGE_IGNORE_INSERTS Or _
               FORMAT_MESSAGE_MAX_WIDTH_MASK
    
    If lngError = 0 Then
        lngError = GetLastError()
    End If
    
    If lngError >= NERR_BASE And lngError <= MAX_NERR Then
        hModule = LoadLibraryEx(DLL_NETMSG, ByVal 0, LOAD_LIBRARY_AS_DATAFILE)
        If hModule <> 0 Then
            lngFlags = lngFlags Or FORMAT_MESSAGE_FROM_HMODULE
        End If
    End If
                
    lngRet = FormatMessage(lngFlags, _
                           ByVal hModule, _
                           lngError, _
                           LANG_USER_DEFAULT, _
                           ByVal strMsg, _
                           256, _
                           0)
                           
    If lngRet Then
        fncGetErrorString = Left$(strMsg, InStr(strMsg, Chr$(0)) - 1)
    End If

    If hModule <> 0 Then FreeLibrary (hModule)
End Function



