Attribute VB_Name = "mUtility"
'¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡
'
'             vWFNg¤ÊÖ[eBeB W[
'
'¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡
Option Explicit
'¡¡¡¡ APIé¾ ¡¡¡¡
Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function StrLen Lib "KERNEL32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Declare Function PtrToStr Lib "KERNEL32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Declare Function StrToPtr Lib "KERNEL32" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long

'------------------------
' AhXæè¶ñæ¾
'------------------------
' lngPoint  : AhX
'
' ßèl    : ¶ñ
Public Function GetPtrToStrA(lngPoint As Long) As String
    Dim byteBuffer(500) As Byte
    Call PtrToStr(byteBuffer(0), lngPoint)
    GetPtrToStrA = Left$(byteBuffer, StrLen(lngPoint))
End Function

'------------------------
' ¶ñæèAhXæ¾
'------------------------
' str    : ¶ñ
'
' ßèl : AhX
'Public Function GetStrToPtrA(str As String) As Long
'    Dim byteBuffer() As Byte
'    byteBuffer() = str & vbNullChar
'    byteBuffer() = StrConv(byteBuffer(), vbFromUnicode)
'    GetStrToPtrA = StrPtr(byteBuffer())
'End Function

