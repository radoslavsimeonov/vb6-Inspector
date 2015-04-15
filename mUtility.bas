Attribute VB_Name = "mUtility"
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'
'             プロジェクト共通関数ユーティリティ モジュール
'
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Option Explicit
'■■■■ API宣言 ■■■■
Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function StrLen Lib "KERNEL32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Declare Function PtrToStr Lib "KERNEL32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Declare Function StrToPtr Lib "KERNEL32" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long

'------------------------
' アドレスより文字列取得
'------------------------
' lngPoint  : アドレス
'
' 戻り値    : 文字列
Public Function GetPtrToStrA(lngPoint As Long) As String
    Dim byteBuffer(500) As Byte
    Call PtrToStr(byteBuffer(0), lngPoint)
    GetPtrToStrA = Left$(byteBuffer, StrLen(lngPoint))
End Function

'------------------------
' 文字列よりアドレス取得
'------------------------
' str    : 文字列
'
' 戻り値 : アドレス
'Public Function GetStrToPtrA(str As String) As Long
'    Dim byteBuffer() As Byte
'    byteBuffer() = str & vbNullChar
'    byteBuffer() = StrConv(byteBuffer(), vbFromUnicode)
'    GetStrToPtrA = StrPtr(byteBuffer())
'End Function

