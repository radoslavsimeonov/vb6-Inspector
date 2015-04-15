Attribute VB_Name = "modVisualTweaks"
Option Explicit

Private Const GWL_STYLE        As Long = (-16)
Private Const GWL_EXSTYLE      As Long = (-20)
Private Const WS_EX_APPWINDOW  As Long = &H40000
Private Const TVS_HASLINES     As Long = 2
Private Const TV_FIRST         As Long = &H1100
Private Const TVM_SETBKCOLOR   As Long = (TV_FIRST + 29)
Private Const TVM_SEcTextCOLOR As Long = (TV_FIRST + 30)
Private Const TVS_CHECKBOXES = &H100
Private Const TVS_TRACKSELECT = &H200
Private Const WS_EX_LAYERED     As Long = &H80000
Private Const WS_BORDER         As Long = &H800000
Private Const WS_CAPTION        As Long = &HC00000
Private Const LWA_COLORKEY      As Long = &H1
Private Const LWA_ALPHA         As Long = &H2
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WM_USER = &H400
Private Const CCM_FIRST       As Long = &H2000&
Private Const CCM_SETBKCOLOR  As Long = (CCM_FIRST + 1)
Private Const PBM_SETBKCOLOR  As Long = CCM_SETBKCOLOR
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
Private Const EC_USEFONTINFO = &HFFFF&
Private Const EM_SETMARGINS = &HD3&
Private Const EM_GETMARGINS = &HD4&
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160
Private Declare Function FindWindowEx _
                Lib "user32" _
                Alias "FindWindowExA" (ByVal hwndParent As Long, _
                                       ByVal hwndChildAfter As Long, _
                                       ByVal lpszClass As String, _
                                       ByVal lpszWindow As String) As Long

Private Declare Function SendMessageLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Public Declare Function GetWindowLong _
               Lib "user32" _
               Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                       ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong _
               Lib "user32" _
               Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal crKey As Long, _
                              ByVal bAlpha As Byte, _
                              ByVal dwFlags As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Declare Function SetCursorPos _
                Lib "user32" (ByVal x As Long, _
                              ByVal Y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect _
                Lib "user32" (ByVal hWnd As Long, _
                              lpRect As RECT) As Long

Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont _
                Lib "gdi32.dll" _
                Alias "CreateFontA" (ByVal nHeight As Long, _
                                     ByVal nWidth As Long, _
                                     ByVal nEscapement As Long, _
                                     ByVal nOrientation As Long, _
                                     ByVal fnWeight As Long, _
                                     ByVal fdwItalic As Long, _
                                     ByVal fdwUnderline As Long, _
                                     ByVal fdwStrikeOut As Long, _
                                     ByVal fdwCharSet As Long, _
                                     ByVal fdwOutputPrecision As Long, _
                                     ByVal fdwClipPrecision As Long, _
                                     ByVal fdwQuality As Long, _
                                     ByVal fdwPitchAndFamily As Long, _
                                     ByVal lpszFace As String) As Long

Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nIndex As Long) As Long

Private Declare Function MulDiv _
                Lib "KERNEL32" (ByVal nNumber As Long, _
                                ByVal nNumerator As Long, _
                                ByVal nDenominator As Long) As Long

Private Const LOGPIXELSY = 90                    'For GetDeviceCaps - returns the height of a logical pixel
Private Const ANSI_CHARSET = 0                   'Use the default Character set
Private Const CLIP_LH_ANGLES = 16                ' Needed for tilted fonts.
Private Const OUT_TT_PRECIS = 4                  'Tell it to use True Types when Possible
Private Const PROOF_QUALITY = 2                  'Make it as clean as possible.
Private Const DEFAULT_PITCH = 0                  'We want the font to take whatever pitch it defaults to
Private Const FF_DONTCARE = 0                    'Use whatever fontface it is.
Private Const DEFAULT_CHARSET = 1

Public Enum FontWeight
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_ULTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_REGULAR = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_DEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_ULTRABOLD = 800
    FW_HEAVY = 900
    FW_BLACK = 900
End Enum

Private mGradient As New clsGradient

Public Sub SetTVBackColor(pobjTV As TreeView, plngBackColor As Long)

    Dim lngTVHwnd As Long
    Dim lngStyle  As Long
    Dim objTVNode As Node

    lngTVHwnd = pobjTV.hWnd
    Call SendMessage(lngTVHwnd, TVM_SETBKCOLOR, 0, ByVal plngBackColor)
    Call SendMessage(lngTVHwnd, TVM_SEcTextCOLOR, 0, ByVal RGB(255, 255, 255))
    Call SetTreeViewAttrib(pobjTV, TVS_TRACKSELECT)
End Sub

Private Sub SetTreeViewAttrib(C As TreeView, ByVal Attrib As Long)

    Const GWL_STYLE As Long = -16

    Dim rStyle      As Long

    rStyle = GetWindowLong(C.hWnd, GWL_STYLE)
    rStyle = rStyle Or Attrib
    Call SetWindowLong(C.hWnd, GWL_STYLE, rStyle)
End Sub

Public Function MakeWindowedControlTransparent(ctlControl As Control) As Long

    Dim Result As Long

    ctlControl.Visible = False
    Result = SetWindowLong(ctlControl.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    ctlControl.Visible = True    ' Use the visible property as a quick VB way of forcing a repaint with the new style
    MakeWindowedControlTransparent = Result
End Function

Public Sub RemoveBorder(ctl As Control)
    ctl.Visible = False
    SetWindowLong ctl.hWnd, GWL_STYLE, (GetWindowLong(ctl.hWnd, GWL_STYLE) And Not WS_BORDER)
    ctl.Visible = True
End Sub

Public Sub SetProgressBarColors(hWnd, cBar As Long, cBG As Long)
    Call SendMessage(hWnd, PBM_SETBKCOLOR, 0&, ByVal cBG)
    Call SendMessage(hWnd, PBM_SETBARCOLOR, 0&, ByVal cBar)
End Sub

Public Sub RemoveFromTaskbar(hWnd, bln As Boolean)

    Select Case bln
        Case True
            SetWindowLong hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_APPWINDOW)
        Case False
            SetWindowLong hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
    End Select

End Sub

Public Sub MinMaxWindow()

    Dim rec   As RECT
    Dim point As POINTAPI

    If MDIMain.picPlaceHolder.Width > 1300 Then
        MDIMain.picPlaceHolder.Width = 1000
        MDIMain.tvwFeatures.Width = 560
        MDIMain.picSlider.Left = MDIMain.picPlaceHolder.Width - MDIMain.picSlider.Width '710
        MDIMain.lblSliderDown.Caption = "> > > > >"
        MDIMain.lblSliderUp.Caption = "> > > > >"
    Else
        MDIMain.picPlaceHolder.Width = 3465
        MDIMain.tvwFeatures.Width = 2895
        MDIMain.picSlider.Left = MDIMain.picPlaceHolder.Width - MDIMain.picSlider.Width '3220
        MDIMain.lblSliderDown.Caption = "< < < < <"
        MDIMain.lblSliderUp.Caption = "< < < < <"
    End If

    GetCursorPos point
    GetWindowRect MDIMain.picSlider.hWnd, rec
    SetCursorPos rec.Left + MDIMain.picSlider.Width / 770 + 10, point.Y
End Sub

Public Sub AddControlToCombo(ByRef ctrl As Control, ByRef cboThis As ComboBox)

    Dim lHWnd   As Long
    Dim lMargin As Long

    lHWnd = FindWindowEx(cboThis.hWnd, 0, "EDIT", vbNullString)

    If (lHWnd <> 0) Then
        lMargin = ctrl.Width \ Screen.TwipsPerPixelX + 2
        SendMessageLong lHWnd, EM_SETMARGINS, EC_LEFTMARGIN, lMargin
    End If

    ctrl.BackColor = cboThis.BackColor
    ctrl.Move cboThis.Left + 3 * Screen.TwipsPerPixelX, cboThis.Top + 2 * Screen.TwipsPerPixelY, ctrl.Width, cboThis.Height - 4 * Screen.TwipsPerPixelY
    ctrl.ZOrder
End Sub

Public Sub CenterCombo(ByRef cboThis As ComboBox, Margin As Long)

    Dim lHWnd As Long

    lHWnd = FindWindowEx(cboThis.hWnd, 0, "EDIT", vbNullString)

    If (lHWnd <> 0) Then
        Margin = Margin \ Screen.TwipsPerPixelX + 2
        SendMessageLong lHWnd, EM_SETMARGINS, EC_LEFTMARGIN, Margin
    End If

End Sub

Public Property Let ComboDropDownWidth(m_cboBox As ComboBox, ByVal NewComboDropDownWidth As Long)
    SendMessage m_cboBox.hWnd, CB_SETDROPPEDWIDTH, NewComboDropDownWidth, 0
End Property

Public Property Get ComboDropDownWidth(m_cboBox As ComboBox) As Long
    ComboDropDownWidth = SendMessage(m_cboBox.hWnd, CB_GETDROPPEDWIDTH, 0, 0)
End Property

Public Sub DrawGradient(picTarget As PictureBox, _
                        Optional mfAngle As Single = 90, _
                        Optional mlColor1 As Long = &H404040, _
                        Optional mlColor2 As Long = &HC0C0C0)

    With mGradient
        .Angle = mfAngle
        .Color1 = mlColor1
        .Color2 = mlColor2
        .Draw picTarget
    End With

    picTarget.Refresh
End Sub

Public Sub DrawRotatedText(ByRef Canvas As Object, _
                           ByVal txt As String, _
                           ByVal x As Single, _
                           ByVal Y As Single, _
                           ByVal font_name As String, _
                           ByVal Size As Long, _
                           ByVal Angle As Single, _
                           ByVal weight As FontWeight, _
                           ByVal Italic As Boolean, _
                           ByVal Underline As Boolean, _
                           ByVal Strikethrough As Boolean)

    Dim newfont     As Long
    Dim oldfont     As Long
    Dim nEscapement As Long
    Dim nHeight     As Long

    nEscapement = Angle * 10
    nHeight = -MulDiv(Size, GetDeviceCaps(Canvas.hDC, LOGPIXELSY), 72)
    newfont = CreateFont(nHeight, 0, nEscapement, nEscapement, weight, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_LH_ANGLES, PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, "Arial")
    oldfont = SelectObject(Canvas.hDC, newfont)
    Canvas.CurrentX = x
    Canvas.CurrentY = Y
    Canvas.Print txt
    newfont = SelectObject(Canvas.hDC, oldfont)
    DeleteObject newfont
End Sub
