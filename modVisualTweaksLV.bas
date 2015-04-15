Attribute VB_Name = "modVisualTweaksLV"
Option Explicit

Private Enum WinNotifications
    NM_FIRST = -0& ' (0U- 0U) ' // generic to all controls
    NM_LAST = -99& ' (0U- 99U)
    NM_OUTOFMEMORY = (NM_FIRST - 1)
    NM_CLICK = (NM_FIRST - 2)
    NM_DBLCLK = (NM_FIRST - 3)
    NM_RETURN = (NM_FIRST - 4)
    NM_RCLICK = (NM_FIRST - 5)
    NM_RDBLCLK = (NM_FIRST - 6)
    NM_SETFOCUS = (NM_FIRST - 7)
    NM_KILLFOCUS = (NM_FIRST - 8)
    NM_CUSTOMDRAW = (NM_FIRST - 12)
    NM_HOVER = (NM_FIRST - 13)
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type NMHDR
    hwndFrom As Long       ' Window handle of control sending message
    idFrom As Long         ' Identifier of control sending message
    code As Long           ' Specifies the notification code
End Type

Private Type NMCUSTOMDRAWINFO
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As RECT
    dwItemSpec As Long
    iItemState As Long
    lItemLParam As Long
End Type

Private Type NMLVCUSTOMDRAW
    nmcmd As NMCUSTOMDRAWINFO
    clrText As Long
    clrTextBk As Long
End Type

Private Const WM_NOTIFY& = &H4E
Private Const WM_DRAWITEM = &H2B

Private Const CDDS_PREPAINT& = &H1
Private Const CDDS_POSTPAINT& = &H2
Private Const CDDS_PREERASE& = &H3
Private Const CDDS_POSTERASE& = &H4
Private Const CDDS_ITEM& = &H10000
Private Const CDDS_ITEMPREPAINT& = CDDS_ITEM Or CDDS_PREPAINT
Private Const CDDS_ITEMPOSTPAINT& = CDDS_ITEM Or CDDS_POSTPAINT
Private Const CDDS_ITEMPREERASE& = CDDS_ITEM Or CDDS_PREERASE
Private Const CDDS_ITEMPOSTERASE& = CDDS_ITEM Or CDDS_POSTERASE
Private Const CDDS_SUBITEM& = &H20000
Private Const CDRF_DODEFAULT& = &H0
Private Const CDRF_NEWFONT& = &H2
Private Const CDRF_SKIPDEFAULT& = &H4
Private Const CDRF_NOTIFYPOSTPAINT& = &H10
Private Const CDRF_NOTIFYITEMDRAW& = &H20
Private Const CDRF_NOTIFYSUBITEMDRAW = &H20     ' flags are the same, we candistinguish by context
Private Const CDRF_NOTIFYPOSTERASE& = &H40
Private Const CDRF_NOTIFYITEMERASE& = &H80

Private Const LVM_FIRST = &H1000&               '      // ListView messages
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54 ' << Note the diff
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55 ' << One is Set,Ohter is Get

Private Const SETEXTENDEDSTYLE = &H1000& + 54
Private Const FULLROWSELECT = &H20
Private Const LVS_EX_HEADERDRAGDROP = &H10

Private Declare Sub CopyMemory _
                Lib "KERNEL32" _
                Alias "RtlMoveMemory" (pDest As Any, _
                                       pSource As Any, _
                                       ByVal dwLength As Long)

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const WM_GETFONT = &H31
Declare Function DrawText _
        Lib "user32" _
        Alias "DrawTextA" (ByVal hDC As Long, _
                           ByVal lpStr As String, _
                           ByVal nCount As Long, _
                           lpRect As RECT, _
                           ByVal wFormat As Long) As Long
Declare Function SetTextAlign _
        Lib "gdi32" (ByVal hDC As Long, _
                     ByVal wFlags As Long) As Long
Declare Function SetTextColor _
        Lib "gdi32" (ByVal hDC As Long, _
                     ByVal crColor As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function InvalidateRect _
        Lib "user32" (ByVal hWnd As Long, _
                      lpRect As RECT, _
                      ByVal bErase As Long) As Long
Declare Function BitBlt _
        Lib "gdi32" (ByVal hDestDC As Long, _
                     ByVal x As Long, _
                     ByVal Y As Long, _
                     ByVal nWidth As Long, _
                     ByVal nHeight As Long, _
                     ByVal hSrcDC As Long, _
                     ByVal xSrc As Long, _
                     ByVal ySrc As Long, _
                     ByVal dwRop As Long) As Long

Private Const GW_OWNER = 4
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_WNDPROC = (-4)
Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long
Declare Function SetWindowLong _
        Lib "user32" _
        Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                ByVal nIndex As Long, _
                                ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc _
        Lib "user32" _
        Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                 ByVal hWnd As Long, _
                                 ByVal Msg As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long

Const WM_PAINT = &HF
Const WM_ERASEBKGND = &H14

Private Const SET_COLUMN_WIDTH As Long = 4126
Private Const AUTOSIZE_USEHEADER As Long = -2

Public glHdrTextClr   As Long     ' The text color of the Header Btns.
Public glHdrBkClr     As Long       ' The background color of Header btns.

Private origLVwinProc As Long
Private m_hooked_lv   As Long

Public Sub HookToLV(hWnd, b As Boolean)

    If b Then
        origLVwinProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf LVSubcls_WProc4Hdr)
        m_hooked_lv = hWnd
        glHdrBkClr = vbYellow   ' I have set some jarring colors by default
        glHdrTextClr = vbRed    ' so that you will change them to your lining!:-)
    Else
        Call SetWindowLong(m_hooked_lv, GWL_WNDPROC, origLVwinProc)
    End If

End Sub

Public Function LVSubcls_WProc4Hdr(ByVal hWnd As Long, _
                                   ByVal Msg As Long, _
                                   ByVal wParam As Long, _
                                   ByVal lParam As Long) As Long

    Dim tMessage         As NMHDR
    Dim lCode            As Long
    Dim tLVRedrawMessage As NMLVCUSTOMDRAW

    Select Case Msg
        Case WM_NOTIFY
            CopyMemory tMessage, ByVal lParam, Len(tMessage)
            lCode = tMessage.code

            Select Case lCode
                Case NM_CUSTOMDRAW
                    CopyMemory tLVRedrawMessage, ByVal lParam, Len(tLVRedrawMessage)

                    If tLVRedrawMessage.nmcmd.dwDrawStage = CDDS_PREPAINT Then
                        LVSubcls_WProc4Hdr = CDRF_NOTIFYITEMDRAW
                        Exit Function
                    End If

                    If tLVRedrawMessage.nmcmd.dwDrawStage = CDDS_ITEMPREPAINT Then
                        SetTextColor tLVRedrawMessage.nmcmd.hDC, glHdrTextClr
                        SetBkColor tLVRedrawMessage.nmcmd.hDC, glHdrBkClr
                        LVSubcls_WProc4Hdr = CDRF_DODEFAULT
                        Exit Function
                    End If

                    If tLVRedrawMessage.nmcmd.dwDrawStage = CDDS_ITEMPOSTPAINT Then
                        LVSubcls_WProc4Hdr = CDRF_DODEFAULT
                        Exit Function
                    End If

            End Select

            LVSubcls_WProc4Hdr = CallWindowProc(origLVwinProc, hWnd, Msg, wParam, lParam)
        Case Else
            LVSubcls_WProc4Hdr = CallWindowProc(origLVwinProc, hWnd, Msg, wParam, lParam)
    End Select

End Function

Public Sub AutoSizeListViewColumns(ByVal TargetListView As ListView)

    Const SET_COLUMN_WIDTH As Long = 4126
    Const AUTOSIZE_USEHEADER As Long = -2

    Dim lngColumn As Long

    For lngColumn = 0 To (TargetListView.ColumnHeaders.count - 1)

        Call SendMessage(TargetListView.hWnd, _
            SET_COLUMN_WIDTH, _
            lngColumn, _
            AUTOSIZE_USEHEADER)

    Next lngColumn

End Sub
