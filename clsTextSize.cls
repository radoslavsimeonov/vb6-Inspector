VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function CreateFontIndirect _
                Lib "gdi32" _
                Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function MulDiv _
                Lib "kernel32" (ByVal nNumber As Long, _
                                ByVal nNumerator As Long, _
                                ByVal nDenominator As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC _
                Lib "user32.dll" (ByVal hwnd As Long, _
                                  ByVal hDC As Long) As Long

Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nIndex As Long) As Long

Private Const LOGPIXELSY = 90

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawText _
                Lib "user32" _
                Alias "DrawTextA" (ByVal hDC As Long, _
                                   ByVal lpStr As String, _
                                   ByVal nCount As Long, _
                                   lpRect As RECT, _
                                   ByVal wFormat As Long) As Long

Private Const DT_CALCRECT = &H400
Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long

Private m_Font As StdFont

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Function TextWidth(ByVal sText As String) As Long

    Dim r As RECT

    r = GetFontSize(sText)
    TextWidth = r.Right
End Function

Public Function TextHeight(ByVal sText As String) As Long

    Dim r As RECT

    r = GetFontSize(sText)
    TextHeight = r.Bottom
End Function

Private Sub OLEFontToLogFont(fntThis As StdFont, ByVal hDC As Long, tLF As LOGFONT)

    Dim sFont As String
    Dim iChar As Integer
    Dim b()   As Byte

    With tLF
        sFont = fntThis.Name
        b = StrConv(sFont, vbFromUnicode)

        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = b(iChar - 1)
        Next iChar

        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic

        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If

        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        .lfCharSet = fntThis.Charset
    End With

End Sub

Private Sub Class_Initialize()
    Set m_Font = New StdFont
    m_Font.Name = "MS Sans Serif"
    m_Font.Size = 8
End Sub

Private Function GetFontSize(ByVal sText As String) As RECT

    Dim hDC  As Long
    Dim tLF  As LOGFONT
    Dim hFnt As Long, hFntOld As Long
    Dim tR   As RECT

    hDC = GetDC(0)
    OLEFontToLogFont m_Font, hDC, tLF
    hFnt = CreateFontIndirect(tLF)
    hFntOld = SelectObject(hDC, hFnt)
    DrawText hDC, sText, -1, tR, DT_CALCRECT
    SelectObject hDC, hFntOld
    DeleteObject hFnt
    Call ReleaseDC(0, hDC)
    GetFontSize = tR
End Function

Private Sub Class_Terminate()
On Error Resume Next

    Set m_Font = Nothing
End Sub
