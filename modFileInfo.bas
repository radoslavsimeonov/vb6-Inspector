Attribute VB_Name = "modFileInfo"
Option Explicit

Public Const FI_COMMENTS As String = "Comments"
Public Const FI_INTERNALNAME As String = "InternalName"
Public Const FI_PRODUCTNAME As String = "ProductName"
Public Const FI_COMPANYNAME As String = "CompanyName"
Public Const FI_LEGALCOPYRIGHT As String = "LegalCopyright"
Public Const FI_PRODUCTVERSION As String = "ProductVersion"
Public Const FI_FILEDESCRIPTION As String = "FileDescription"
Public Const FI_LEGALTRADEMARKS As String = "LegalTrademarks"
Public Const FI_PRIVATEBUILD As String = "PrivateBuild"
Public Const FI_FILEVERSION As String = "FileVersion"
Public Const FI_ORIGINALFILENAME As String = "OriginalFilename"
Public Const FI_SPECIALBUILD As String = "SpecialBuild"

Private Declare Function GetFileVersionInfo _
                Lib "Version.dll" _
                Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, _
                                             ByVal dwhandle As Long, _
                                             ByVal dwlen As Long, _
                                             lpData As Any) As Long

Private Declare Function GetFileVersionInfoSize _
                Lib "Version.dll" _
                Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
                                                 lpdwHandle As Long) As Long

Private Declare Function VerQueryValue _
                Lib "Version.dll" _
                Alias "VerQueryValueA" (pBlock As Any, _
                                        ByVal lpSubBlock As String, _
                                        lplpBuffer As Any, _
                                        puLen As Long) As Long

Private Declare Sub MoveMemory _
                Lib "KERNEL32" _
                Alias "RtlMoveMemory" (dest As Any, _
                                       ByVal Source As Long, _
                                       ByVal Length As Long)

Private Declare Function lstrcpy _
                Lib "KERNEL32" _
                Alias "lstrcpyA" (ByVal lpString1 As String, _
                                  ByVal lpString2 As Long) As Long
                                  
Public Function GetFileVendor2(ByVal sFilePath As String, Optional strDetail As String = FI_COMPANYNAME) As String
On Error Resume Next

    Dim lBufferLen          As Long, lDummy As Long
    Dim sBuffer()           As Byte
    Dim lVerPointer         As Long
    Dim lRet                As Long
    Dim Lang_Charset_String As String
    Dim HexNumber           As Long
    Dim i                   As Integer
    Dim strTemp             As String

    GetFileVendor2 = ""
    sFilePath = Replace(sFilePath, Chr$(34), "")
    sFilePath = LCase$(sFilePath)

    If Is64bit And InStr(LCase$(sFilePath), "\system32\") > 0 Then
        sFilePath = Replace$(LCase$(sFilePath), "\system32\", "\Sysnative\")
    End If

    If InStr(LCase$(sFilePath), "svchost") > InStr(LCase$(sFilePath), "svchost.exe") Then sFilePath = Replace(sFilePath, "svchost", "svchost.exe")
    
    sFilePath = Left$(sFilePath, InStr(sFilePath, ".exe") + 3)
    sFilePath = Replace(sFilePath, "\", "\\")
    
    lBufferLen = GetFileVersionInfoSize(sFilePath, lDummy)

    If lBufferLen < 1 Then
        Exit Function
    End If

    ReDim sBuffer(lBufferLen)
    
    lRet = GetFileVersionInfo(sFilePath, 0&, lBufferLen, sBuffer(0))

    If lRet = 0 Then
        Exit Function
    End If

    lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)

    If lRet = 0 Then
        Exit Function
    End If

    Dim byteBuffer(255) As Byte

    MoveMemory byteBuffer(0), lVerPointer, lBufferLen
    HexNumber = byteBuffer(2) + byteBuffer(3) * &H100 + byteBuffer(0) * &H10000 + byteBuffer(1) * &H1000000
    Lang_Charset_String = Hex$(HexNumber)

    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop

    Dim Buffer As String

    Buffer = String$(255, 0)
    strTemp = "\StringFileInfo\" & Lang_Charset_String & "\" & strDetail
    lRet = VerQueryValue(sBuffer(0), strTemp, lVerPointer, lBufferLen)

    If lRet = 0 Then
        Exit Function
    End If

    lstrcpy Buffer, lVerPointer
    Buffer = Mid$(Buffer, 1, InStr(Buffer, vbNullChar) - 1)

    If i = 0 Then GetFileVendor2 = Trim$(Buffer) Else GetFileVendor2 = ""
End Function

