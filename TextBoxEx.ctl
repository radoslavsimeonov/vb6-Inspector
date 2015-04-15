VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TextBoxEx 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   FillColor       =   &H00FFFFFF&
   PropertyPages   =   "TextBoxEx.ctx":0000
   ScaleHeight     =   1350
   ScaleWidth      =   4365
   Begin VB.PictureBox picClearText 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3000
      MouseIcon       =   "TextBoxEx.ctx":0026
      MousePointer    =   99  'Custom
      Picture         =   "TextBoxEx.ctx":0178
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picDrop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000080FF&
      Height          =   150
      Left            =   3240
      MouseIcon       =   "TextBoxEx.ctx":05FA
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picPin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   840
      Picture         =   "TextBoxEx.ctx":074C
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   30
      Width           =   135
   End
   Begin VB.PictureBox picUnpin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   960
      Picture         =   "TextBoxEx.ctx":0798
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   30
      Width           =   135
   End
   Begin VB.Timer tmrSlide 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   3600
      Top             =   0
   End
   Begin MSComctlLib.ListView lstListData 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1085
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   5263440
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   45
      TabIndex        =   3
      Tag             =   "EDIT"
      Text            =   "Text1"
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "lblCaption"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   690
   End
   Begin VB.Line lLeft 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   0
      Y1              =   480
      Y2              =   510
   End
   Begin VB.Line lRight 
      BorderColor     =   &H000080FF&
      X1              =   4320
      X2              =   4320
      Y1              =   480
      Y2              =   515
   End
   Begin VB.Line lDown 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   4320
      Y1              =   510
      Y2              =   510
   End
End
Attribute VB_Name = "TextBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Const ALTERNATE = 1
Private Const WINDING = 2

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function SetPolyFillMode _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nPolyFillMode As Long) As Long

Private Declare Function GetPolyFillMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Polygon _
                Lib "gdi32" (ByVal hDC As Long, _
                             lpPoint As POINTAPI, _
                             ByVal nCount As Long) As Long

Private Const SB_BOTH = 3
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Declare Function ShowScrollBar _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal wBar As Long, _
                              ByVal bShow As Long) As Long

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function SystemParametersInfo _
                Lib "user32" _
                Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                               ByVal uParam As Long, _
                                               lpvParam As Any, _
                                               ByVal fuWinIni As Long) As Long

Private Const GWL_EXSTYLE      As Long = -20
Private Const WS_EX_TOOLWINDOW As Long = 128
Private Declare Function SetParent _
                Lib "user32" (ByVal hWndChild As Long, _
                              ByVal hWndNewParent As Long) As Long

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type POINTL
    x       As Long
    y       As Long
End Type

Private Type Msg
    hWnd    As Long
    Message As Long
    wParam  As Long
    lParam  As Long
    time    As Long
    pt      As POINTL
End Type

Dim ControlPos As RECT
Dim bolSliding As Boolean
Dim ColsHeight As Long
Dim oFont      As New clsTextSize

Private Declare Function GetWindowRect _
                Lib "user32" (ByVal hWnd As Long, _
                              lpRect As RECT) As Long

Private Declare Function GetMessage _
                Lib "user32" _
                Alias "GetMessageA" (lpMsg As Msg, _
                                     ByVal hWnd As Long, _
                                     ByVal wMsgFilterMin As Long, _
                                     ByVal wMsgFilterMax As Long) As Long

Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage _
                Lib "user32" _
                Alias "DispatchMessageA" (lpMsg As Msg) As Long

Public Enum eType
    TextBox = 0
    ComboBox
End Enum

Public Enum eComboType
    DropDownCombo
    DropDownList
End Enum

Public Enum ePinned
    None = 0
    Pin
    Unpin
End Enum

Private Const PRP_BCOLOR         As String = "BackColor"                   'Color property name
Private Const DEF_BCOLOR         As Long = &H404040
Private Const PRP_FCOLOR         As String = "ForeColor"                   'Color property name
Private Const DEF_FCOLOR         As Long = &HFFFFFF
Private Const PRP_CAPTION        As String = "Caption"
Private Const DEF_CAPTION        As String = "Caption"
Private Const PRP_CAPTIONCOLOR   As String = "CaptionColor"
Private Const DEF_CAPTIONCOLOR   As String = &H80FF&
Private Const PRP_UNDERLINECOLOR As String = "UnderlineColor"
Private Const DEF_UNDERLINECOLOR As String = &HFFFFFF
Private Const PRP_TEXT           As String = "Text"
Private Const DEF_TEXT           As String = ""
Private Const PRP_LOCKED         As String = "Locked"
Private Const DEF_LOCKED         As String = "False"
Private Const PRP_PASSWORD       As String = "Password"
Private Const DEF_PASSWORD       As String = "False"
Private Const PRP_MINPASS        As String = "MinPass"
Private Const DEF_MINPASS        As Integer = "8"
Private Const PRP_TIPS           As String = "Tips"
Private Const DEF_TIPS           As String = ""
Private Const PRP_REQUIRED       As String = "Required"
Private Const DEF_REQUIRED       As String = "False"
Private Const PRP_ENABLED        As String = "Enabled"
Private Const DEF_ENABLED        As String = "True"
Private Const PRP_COLUMNS        As String = "Columns"
Private Const DEF_COLUMNS        As Long = "1"
Private Const PRP_MAXROWS        As String = "MaxRowsToDisplay"
Private Const DEF_MAXROWS        As Long = "8"
Private Const PRP_TYPE           As String = "ControlType"
Private Const DEF_TYPE           As Integer = eType.TextBox
Private Const PRP_COMBOSTYLE     As String = "ComboStyle"
Private Const DEF_COMBOSTYLE     As Integer = eComboType.DropDownCombo
Private Const PRP_SORTED         As String = "Sorted"
Private Const DEF_SORTED         As Boolean = False
Private Const PRP_PINNED         As String = "Pinned"
Private Const DEF_PINNED         As Integer = ePinned.None
Private Const PRP_PINTOOLTIP     As String = "PinToolTip"
Private Const DEF_PINTOOLTIP     As String = ""
Private m_bcolor                 As OLE_COLOR
Private m_fcolor                 As OLE_COLOR
Private m_caption                As String
Private m_text                   As String
Private m_caption_color          As OLE_COLOR
Private m_uncerline_color        As OLE_COLOR
Private m_enabled                As Boolean
Private m_required               As Boolean
Private m_locked                 As Boolean
Private m_password               As Boolean
Private m_min_pass               As Integer
Private m_tips                   As String
Private m_maxRows                As Long
Private m_control_type           As eType
Private m_combo_style            As eComboType
Private m_sorted                 As Boolean
Private m_pinned                 As ePinned
Private m_pin_tool_tip           As String

Event TextChange()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event ShowDropDown()
Event HideDropDown()
Event ItemSelect(ByRef Item As ListItem)
Event ItemPinned(state As ePinned)

Private Const REQUIRED_COLOR  As Long = &H2141FF
Private blnHasFocus           As Boolean
Private mblnSelfItemClickFire As Boolean
Private bDelKey               As Boolean
Private bStrFound             As Boolean
Private lColumnsSize          As Long

Public Function IsPassStrong() As Boolean

    If IsComplex(Text1.Text) = "Very Strong" Then IsPassStrong = True Else IsPassStrong = False
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    Enabled = m_enabled
End Property

Public Property Let Enabled(val As Boolean)
    m_enabled = val

    If Ambient.UserMode Then
        PropertyChanged DEF_ENABLED
    End If

    lblCaption.Enabled = m_enabled
    UserControl.Enabled = m_enabled
End Property

Public Property Get SelectedItem() As ListItem
    If (Not lstListData.SelectedItem Is Nothing) Then Set SelectedItem = lstListData.SelectedItem
End Property

Public Property Get Required() As Boolean
Attribute Required.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    Required = m_required
End Property

Public Property Let Required(val As Boolean)
    m_required = val

    If Ambient.UserMode Then
        PropertyChanged DEF_REQUIRED
    End If

    lblCaption.ForeColor = IIf(m_required, REQUIRED_COLOR, m_caption_color)
End Property

Public Property Get Pinned() As ePinned
    Pinned = m_pinned
End Property

Public Property Let Pinned(val As ePinned)
    m_pinned = val

    If Ambient.UserMode Then
        PropertyChanged DEF_PINNED
    End If

    TogglePin
End Property

Public Property Get PinToolTip() As String
    PinToolTip = m_pin_tool_tip
End Property

Public Property Let PinToolTip(val As String)
    m_pin_tool_tip = val

    If Ambient.UserMode Then
        PropertyChanged DEF_PINTOOLTIP
    End If

    picPin.ToolTipText = m_pin_tool_tip
    m_pinned = (Len(m_pin_tool_tip) > 1) + 2
    TogglePin
End Property

Private Sub TogglePin()

    On Error Resume Next

    Select Case m_pinned
        Case ePinned.None
            picPin.Visible = False
            picUnpin.Visible = False
        Case ePinned.Pin
            picPin.Visible = True
            picPin.SetFocus
            picUnpin.Visible = False
        Case ePinned.Unpin
            picPin.Visible = False
            picPin.ToolTipText = ""
            picUnpin.Visible = True
            picUnpin.SetFocus
    End Select

    RaiseEvent ItemPinned(m_pinned)
End Sub

Public Property Get Sorted() As Boolean
    Sorted = m_sorted
End Property

Public Property Let Sorted(val As Boolean)
    m_sorted = val

    If Ambient.UserMode Then
        PropertyChanged DEF_SORTED
    End If

    lstListData.Sorted = m_sorted
End Property

Public Property Get ControlType() As eType
    ControlType = m_control_type
End Property

Public Property Let ControlType(val As eType)
    m_control_type = val

    If Ambient.UserMode Then
        PropertyChanged DEF_TYPE
    End If

    Select Case m_control_type
        Case eType.TextBox
            picDrop.Visible = False
        Case eType.ComboBox
            picDrop.Visible = True
    End Select

End Property

Public Property Get ComboStyle() As eComboType
    ComboStyle = m_combo_style
End Property

Public Property Let ComboStyle(val As eComboType)
    m_combo_style = val

    If Ambient.UserMode Then
        PropertyChanged DEF_COMBOSTYLE
    End If

End Property

Public Property Get MinPassLen() As Integer
Attribute MinPassLen.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    MinPassLen = m_min_pass
End Property

Public Property Let MinPassLen(val As Integer)
    m_min_pass = val

    If Ambient.UserMode Then
        PropertyChanged DEF_MINPASS
    End If

End Property

Public Property Let SelStart(val As Integer)
    Text1.SelStart = val
End Property

Public Property Let SelLength(val As Integer)
    Text1.SelLength = val
End Property

Public Property Get SelLength() As Integer
Attribute SelLength.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    SelLength = Text1.SelLength
End Property

Public Property Get Tips() As String
Attribute Tips.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    Tips = m_tips
End Property

Public Property Let Tips(val As String)
    m_tips = val

    If Ambient.UserMode Then
        PropertyChanged DEF_TIPS
    End If

    Call UserControl_ExitFocus
End Property

Public Property Get Password() As Boolean
Attribute Password.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    Password = m_password
End Property

Public Property Let Password(val As Boolean)
    m_password = val

    If Ambient.UserMode Then
        PropertyChanged DEF_PASSWORD
    End If

    Text1.PasswordChar = IIf(m_password, "*", "")
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    Locked = m_locked
End Property

Public Property Let Locked(val As Boolean)
    m_locked = val

    If Ambient.UserMode Then
        PropertyChanged DEF_LOCKED
    End If

    Text1.Locked = m_locked
    Text1.MousePointer = IIf(m_locked, vbArrow, vbIbeam)
End Property

Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    Text = m_text
End Property

Public Property Let Text(txt As String)
    m_text = txt

    If Ambient.UserMode Then
        PropertyChanged DEF_TEXT
    End If

    Text1.Text = m_text
    picClearText.Visible = False
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    Caption = m_caption
End Property

Public Property Let Caption(txt As String)
    m_caption = txt

    If Ambient.UserMode Then
        PropertyChanged DEF_CAPTION
    End If

    lblCaption.Caption = m_caption
End Property

Public Property Get UnderlineColor() As OLE_COLOR
    UnderlineColor = m_uncerline_color
End Property

Public Property Let UnderlineColor(sColor As OLE_COLOR)
    m_uncerline_color = sColor

    If Ambient.UserMode Then
        PropertyChanged DEF_UNDERLINECOLOR
    End If

    SetUnderlineColor m_uncerline_color
End Property

Private Sub SetUnderlineColor(color As OLE_COLOR)
    lLeft.BorderColor = color
    lDown.BorderColor = color
    lRight.BorderColor = color
    picDrop.FillColor = color
    picDrop.ForeColor = color
End Sub

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = m_caption_color
End Property

Public Property Let CaptionColor(sColor As OLE_COLOR)
    m_caption_color = sColor

    If Ambient.UserMode Then
        PropertyChanged DEF_CAPTIONCOLOR
    End If

    lblCaption.ForeColor = m_caption_color
End Property

Public Property Get BackgroundColor() As OLE_COLOR
    BackgroundColor = m_bcolor
End Property

Public Property Let BackgroundColor(sColor As OLE_COLOR)
    m_bcolor = sColor

    If Not Ambient.UserMode Then
        PropertyChanged PRP_BCOLOR
    End If

    Text1.BackColor = m_bcolor
    lblCaption.BackColor = m_bcolor
    BackColor = m_bcolor
End Property

Public Property Get ForegroundColor() As OLE_COLOR
    ForegroundColor = m_fcolor
End Property

Public Property Let ForegroundColor(sColor As OLE_COLOR)
    m_fcolor = sColor

    If Not Ambient.UserMode Then
        PropertyChanged PRP_FCOLOR
    End If

    Text1.ForeColor = m_fcolor
End Property

Private Sub picDrop_Click()

    Dim udtMsg As Msg
    Dim K      As Integer

    GetWindowRect UserControl.hWnd, ControlPos
    DoEvents

    If Not lstListData.Visible Then
        RaiseEvent ShowDropDown
        Text1.SetFocus
        SetParent lstListData.hWnd, 0
        SetWindowLong lstListData.hWnd, GWL_EXSTYLE, (GetWindowLong(lstListData.hWnd, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW)
        lstListData.ZOrder 0
        UserControl.ScaleMode = 3

        If lstListData.ListItems.count < m_maxRows Then
            ColsHeight = lstListData.ListItems.count
        Else
            ColsHeight = m_maxRows
        End If

        For K = 1 To lstListData.ColumnHeaders.count
            lColumnsSize = lColumnsSize + lstListData.ColumnHeaders(K).Width + 10
        Next K

        lstListData.Height = 0
        tmrSlide.Enabled = True
        bolSliding = True

        Do While (bolSliding)

            If GetMessage(udtMsg, UserControl.hWnd, 0, 0) Then
                TranslateMessage udtMsg
                DispatchMessage udtMsg
            End If

        Loop

        UserControl.ScaleMode = 1
        lstListData.SetFocus
    Else
        RaiseEvent HideDropDown
        lstListData.Visible = False
        DoEvents
        UserControl.SetFocus
    End If

End Sub

Public Sub ColapseList()

    If lstListData.Visible Then
        lstListData.Visible = False
        tmrSlide.Enabled = False
        bolSliding = False
        DoEvents
        Call UserControl_Resize
    End If

End Sub

Private Sub lstListData_Click()
    Call ColapseList
End Sub

Private Sub lstListData_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

    On Error Resume Next

    Dim Item As ListItem

    Set Item = lstListData.HitTest(x, y)

    If Not Item Is Nothing Then
        Item.Selected = True
        lstListData.ListItems(Item.Index - 1).EnsureVisible
        lstListData.ListItems(Item.Index + 1).EnsureVisible
    End If

End Sub

Private Sub picPin_Click()
    m_pinned = Unpin
    TogglePin
End Sub

Private Sub picUnpin_Click()

    If Len(Trim$(Text1.Text)) = 0 Then Exit Sub
    m_pinned = Pin
    TogglePin
End Sub

Private Sub Text1_Click()

    If m_combo_style = DropDownList And lstListData.Visible = False Then
        DropList
        Exit Sub
    End If

    If lstListData.Visible = True Then picDrop_Click
End Sub

Private Sub tmrSlide_Timer()

    Dim lngColSize As Long
    Dim lngTop     As Long

    If lColumnsSize > UserControl.ScaleWidth - 3 Then
        lngColSize = lColumnsSize
    Else
        lngColSize = UserControl.ScaleWidth - 3
    End If

    If (ControlPos.Bottom + (oFont.TextHeight("W") + 1) * ColsHeight + 2) > lDesktopHeight Then
        lngTop = ControlPos.Top - ((oFont.TextHeight("W") + 1) * ColsHeight + 2) + lblCaption.Height
    Else
        lngTop = ControlPos.Bottom
    End If

    lstListData.Move ControlPos.Left + 1, lngTop, lngColSize, (oFont.TextHeight("W") + 1) * ColsHeight + 2
    tmrSlide.Enabled = False
    DoEvents
    bolSliding = False
    lColumnsSize = 0

    If lstListData.Visible = False Then lstListData.Visible = True
End Sub

Private Function lDesktopHeight() As Long

    Const SPI_GETWORKAREA As Long = 48

    Dim rc                As RECT

    Call SystemParametersInfo(SPI_GETWORKAREA, vbNull, rc, 0)
    lDesktopHeight = (rc.Bottom - rc.Top)
End Function

Public Sub DropList()

    If lstListData.Visible = False Then
        Call picDrop_Click
    End If

End Sub

Private Sub picClearText_MouseUp(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)

    On Error Resume Next

    m_text = ""
    Text1.Text = m_text
    lstListData.ListItems(1).Selected = True
    lstListData.ListItems(1).EnsureVisible
    ColapseList
    RaiseEvent TextChange
End Sub

Private Sub Text1_Change()
   
    If oFont.TextWidth(Text1.Text) * Screen.TwipsPerPixelX > Text1.Width Then
        Text1.ToolTipText = Text1.Text
    Else
        Text1.ToolTipText = ""
    End If

    If Text1.Text = m_tips Then
        m_text = ""
        Text1.ForeColor = &HC0C0C0
    Else
        m_text = Text1.Text
        Text1.ForeColor = m_fcolor
    End If

    If Text1.PasswordChar <> "" Then
        Text1.ForeColor = vbRed

        If IsComplex(Text1.Text) = "Very Strong" Then Text1.ForeColor = m_fcolor
    End If

    If Len(m_text) > 0 Then
        lblCaption.ForeColor = IIf(m_enabled, m_caption_color, &H808080)

        If blnHasFocus And m_locked = False Then picClearText.Visible = True
    Else

        If m_required = True Then lblCaption.ForeColor = REQUIRED_COLOR
        picClearText.Visible = False
    End If

    RaiseEvent TextChange
    AutoComplete
End Sub

Private Sub AutoComplete()

    Dim tmpText As String
    Dim itmX    As ListItem
    Dim strt    As Integer
    
    If m_control_type = ComboBox Then

        Dim ins As Integer

        mblnSelfItemClickFire = True

        If Len(Trim$(Text1.Text)) = 0 Then
            lstListData.ListItems(1).Selected = True
            lstListData.SelectedItem.EnsureVisible
            bStrFound = False
            Exit Sub
        End If

        If Not bStrFound Then Set itmX = lstListData.FindItem(Text1.Text, , , 1)
        
        If Not itmX Is Nothing Then
            bStrFound = True
            tmpText = Text1.Text
        Else
            ins = InStr(Text1.Text, " ")

            If Text1.SelStart + 1 <= ins Then Exit Sub
            tmpText = LTrim$(Mid$(Text1.Text, ins + 1, Len(Text1.Text)))
        End If

        If Len(tmpText) = 0 Then Exit Sub
        Set itmX = lstListData.FindItem(tmpText)

        If itmX Is Nothing Then Set itmX = lstListData.FindItem(tmpText, , , 1)
        
        If itmX Is Nothing Then
            bStrFound = False
            If lstListData.ListItems.count > 0 Then
                lstListData.ListItems(1).Selected = True
                lstListData.SelectedItem.EnsureVisible
            End If
        Else
            bStrFound = True
            itmX.Selected = True
            itmX.EnsureVisible

            If Not bDelKey Then
                strt = ins
                Text1.Text = ProperCase(Text1.Text + Mid$(lstListData.SelectedItem.Text, Len(tmpText) + 1, Len(lstListData.SelectedItem.Text)))
                Text1.SelStart = strt + Len(tmpText)
                Text1.SelLength = Len(Text1.Text) - strt
            End If
        End If

        bDelKey = False
    End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    RaiseEvent KeyDown(KeyCode, Shift)

    If ((KeyCode = vbKeyDown)) And (Shift = 0) Then
        lstListData.ListItems(lstListData.SelectedItem.Index + 1).Selected = True
        lstListData.ListItems(lstListData.SelectedItem.Index + 1).EnsureVisible
        mblnSelfItemClickFire = True
        lstListData_ItemClick lstListData.SelectedItem
        KeyCode = 0
    ElseIf ((KeyCode = vbKeyUp)) And (Shift = 0) Then
        lstListData.ListItems(lstListData.SelectedItem.Index - 1).Selected = True
        lstListData.ListItems(lstListData.SelectedItem.Index - 1).EnsureVisible
        mblnSelfItemClickFire = True
        lstListData_ItemClick lstListData.SelectedItem
        KeyCode = 0
    ElseIf ((KeyCode = vbKeyDown) And (Shift And vbAltMask)) Then

        If (lstListData.Visible) Then
            ColapseList
        Else
            DropList
        End If

    ElseIf (KeyCode = 13) Then

        If (lstListData.Visible) Then
            mblnSelfItemClickFire = False
            lstListData_ItemClick lstListData.SelectedItem
        Else
            Text1.SelStart = Len(Text1.Text)
        End If

    ElseIf ((KeyCode = vbKeyUp) And (Shift And vbAltMask)) Or (KeyCode = vbKeyEscape) Then
        ColapseList
    End If

    If KeyCode = vbKeyDelete Or KeyCode = 8 Then bDelKey = True

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

    If UCase$(PropertyName) = "BACKCOLOR" Then
        BackgroundColor = Ambient.BackColor
        picClearText.BackColor = Ambient.BackColor
    End If

End Sub

Private Sub UserControl_Click()

    If m_combo_style = DropDownList And lstListData.Visible = False Then
        DropList
        Exit Sub
    End If

    If lstListData.Visible = True Then picDrop_Click
End Sub

Private Sub UserControl_EnterFocus()
    SetUnderlineColor &HC0C000
    DrawComboArrow

    If Len(m_text) = 0 And Text1.Text = m_tips Then
        Text1.Text = ""
    End If

    If Len(m_text) > 0 And m_locked = False Then picClearText.Visible = True Else picClearText.Visible = False
    If (Text1.SelStart = Len(Text1.Text) Or Text1.SelStart = 0) And Text1.SelLength = 0 Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
    End If

    blnHasFocus = True
End Sub

Private Sub UserControl_ExitFocus()
    Call ColapseList
    SetUnderlineColor m_uncerline_color
    DrawComboArrow

    If Len(m_text) = 0 Then
        Text1.ForeColor = &HC0C0C0
        Text1.Text = m_tips
    Else
        If m_password = False Then Text1.ForeColor = m_fcolor
    End If

    picClearText.Visible = False
    Text1.SelStart = 0
    Text1.SelLength = 0
    blnHasFocus = False
End Sub

Private Sub UserControl_InitProperties()
    m_bcolor = DEF_BCOLOR
    m_fcolor = DEF_FCOLOR
    m_caption = DEF_CAPTION
    m_caption_color = DEF_CAPTIONCOLOR
    m_text = DEF_TEXT
    m_uncerline_color = DEF_UNDERLINECOLOR
    m_locked = DEF_LOCKED
    m_password = DEF_PASSWORD
    m_min_pass = DEF_MINPASS
    m_tips = DEF_TIPS
    m_required = DEF_REQUIRED
    m_enabled = DEF_ENABLED
    m_maxRows = DEF_MAXROWS
    Columns = Me.Columns
    m_control_type = DEF_TYPE
    m_combo_style = DEF_COMBOSTYLE
    m_sorted = DEF_SORTED
    m_pinned = DEF_PINNED
    m_pin_tool_tip = DEF_PINTOOLTIP
    Me.Columns = 1
    picClearText.Visible = False
    picClearText.BackColor = Ambient.BackColor
    UserControl.Enabled = m_enabled
    UserControl.BackColor = Ambient.BackColor
    Text1.BackColor = Ambient.BackColor
    Text1.Enabled = m_enabled
    Text1.TabStop = m_enabled
    lblCaption.BackColor = Ambient.BackColor

    If m_required = True Then
        lblCaption.ForeColor = REQUIRED_COLOR
    Else
        lblCaption.ForeColor = IIf(m_enabled, m_caption_color, &H808080)
    End If

    TogglePin
    SetUnderlineColor m_uncerline_color
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_bcolor = .ReadProperty(PRP_BCOLOR, Ambient.BackColor)
        m_fcolor = .ReadProperty(PRP_FCOLOR, DEF_FCOLOR)
        m_caption = .ReadProperty(PRP_CAPTION, Extender.Name)
        m_caption_color = .ReadProperty(PRP_CAPTIONCOLOR, DEF_CAPTIONCOLOR)
        m_text = .ReadProperty(PRP_TEXT, DEF_TEXT)
        m_locked = .ReadProperty(PRP_LOCKED, DEF_LOCKED)
        m_uncerline_color = .ReadProperty(PRP_UNDERLINECOLOR, DEF_UNDERLINECOLOR)
        m_password = .ReadProperty(PRP_PASSWORD, DEF_PASSWORD)
        m_min_pass = .ReadProperty(PRP_MINPASS, DEF_MINPASS)
        m_tips = .ReadProperty(PRP_TIPS, DEF_TIPS)
        m_required = .ReadProperty(PRP_REQUIRED, DEF_REQUIRED)
        m_enabled = .ReadProperty(PRP_ENABLED, DEF_ENABLED)
        Columns = .ReadProperty(PRP_COLUMNS, DEF_COLUMNS)
        m_maxRows = .ReadProperty(PRP_MAXROWS, DEF_MAXROWS)
        m_control_type = .ReadProperty(PRP_TYPE, DEF_TYPE)
        m_combo_style = .ReadProperty(PRP_COMBOSTYLE, DEF_COMBOSTYLE)
        m_sorted = .ReadProperty(PRP_SORTED, DEF_SORTED)
        m_pinned = .ReadProperty(PRP_PINNED, DEF_PINNED)
        m_pin_tool_tip = .ReadProperty(PRP_PINTOOLTIP, DEF_PINTOOLTIP)
    End With

    SetUnderlineColor m_uncerline_color

    Select Case m_control_type
        Case eType.TextBox
            picDrop.Visible = False
        Case eType.ComboBox
            lstListData.ListItems.Add , , ""
            picDrop.Visible = True
            lstListData.Sorted = m_sorted
    End Select

    UserControl.Enabled = m_enabled
    UserControl.BackColor = m_bcolor
    TogglePin
    picClearText.Move ScaleWidth - 16 * Screen.TwipsPerPixelX - picDrop.Width * m_control_type, Text1.Top + 35
    picClearText.BackColor = Ambient.BackColor
    picPin.Move 2 * Screen.TwipsPerPixelX, Text1.Top + 40
    picUnpin.Move 2 * Screen.TwipsPerPixelX, Text1.Top + 40
    Text1.Left = 4 * Screen.TwipsPerPixelX + picPin.Width * ((m_pinned = None) + 1)
    Text1.ForeColor = m_fcolor
    Text1.BackColor = m_bcolor
    Text1.Text = IIf(Len(m_text) = 0, m_tips, m_text)
    Text1.Locked = m_locked
    Text1.MousePointer = IIf(m_locked, vbArrow, vbIbeam)
    Text1.PasswordChar = IIf(m_password, "*", "")
    lblCaption.Caption = m_caption
    lblCaption.BackColor = m_bcolor
    lblCaption.ForeColor = IIf(m_required, REQUIRED_COLOR, m_caption_color)
    lblCaption.ToolTipText = UserControl.Extender.ToolTipText
End Sub

Public Property Get ListItems() As IListItems
    Set ListItems = lstListData.ListItems
End Property

Public Property Let ListItems(ByVal vNewValue As IListItems)
End Property

Private Sub lstListData_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lstListData_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RaiseEvent ItemSelect(lstListData.SelectedItem)

    If (Not mblnSelfItemClickFire) Then
        RaiseEvent HideDropDown
        SetParent lstListData.hWnd, 0&
        Call ColapseList
        DoEvents
        mblnSelfItemClickFire = False
    End If

    Text1.Text = Item.Text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Public Property Let ListIndex(val As Integer)

    If val = 0 Then GoTo NotFound
    If Not lstListData.ListItems(val) Is Nothing Then
        lstListData_ItemClick lstListData.ListItems(val)
        lstListData.ListItems(val).Selected = True
        lstListData.ListItems(val).EnsureVisible
    Else
NotFound:
        Set lstListData.SelectedItem = Nothing
        Text1.Text = ""
    End If

End Property

Public Property Get ListIndex() As Integer

    If Ambient.UserMode Then
        ListIndex = lstListData.SelectedItem.Index
    End If

End Property

Public Sub Clear()

    Do Until lstListData.ListItems.count = 0
        lstListData.ListItems.Remove 1
    Loop

End Sub

Public Property Get MaxRowsToDisplay() As Long
Attribute MaxRowsToDisplay.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    MaxRowsToDisplay = m_maxRows
End Property

Public Property Let MaxRowsToDisplay(ByVal vNewValue As Long)

    If vNewValue >= 1 Then
        m_maxRows = vNewValue
    Else
        m_maxRows = 1
    End If

End Property

Public Sub AutoResizeColumns()

    Dim K As Long, TotalWidth As Single, ColumnsSize As Single

    UserControl_Resize
    ResizeListColumns lstListData
    K = lstListData.ColumnHeaders.count + 1
    TotalWidth = lstListData.ColumnHeaders(K - 1).Left + lstListData.ColumnHeaders(K - 1).Width + 450 ' why 450 ? no clue...

    If lstListData.Width - TotalWidth > 0 And lstListData.Width - TotalWidth > lstListData.ColumnHeaders(K - 1).Width Then
        lstListData.ColumnHeaders(K - 1).Width = lstListData.Width - TotalWidth
    End If

End Sub

Private Sub ResizeListColumns(lstReport As ListView)

    Dim MaxVal As Single, tmpVal As Single
    Dim K      As Long, Q As Long

    If lstReport.ColumnHeaders.count = 1 Then
        lstReport.ColumnHeaders(1).Width = lstReport.Width - 20 * Screen.TwipsPerPixelX
        Exit Sub
    End If

    For K = 0 To lstReport.ColumnHeaders.count - 1
        MaxVal = oFont.TextWidth(lstReport.ColumnHeaders(K + 1).Text)

        For Q = 1 To lstReport.ListItems.count

            If K = 0 Then
                tmpVal = oFont.TextWidth(lstReport.ListItems(Q).Text)
            Else

                On Error Resume Next

                tmpVal = oFont.TextWidth(Left$(lstReport.ListItems(Q).ListSubItems(K).Text, 200))

                If Err.Number <> 0 Then
                    tmpVal = 0
                    Err.Clear
                End If

            End If

            If tmpVal > MaxVal Then MaxVal = tmpVal
        Next Q

        If MaxVal = 0 Then
            lstReport.ColumnHeaders(K + 1).Width = Screen.TwipsPerPixelX
        Else
            lstReport.ColumnHeaders(K + 1).Width = (MaxVal + 13) * Screen.TwipsPerPixelX
        End If

    Next K

End Sub

Private Sub UserControl_Resize()
On Error Resume Next

    Dim iLine As Integer

    iLine = Text1.Top + Text1.Height + 10
    lDown.X1 = 0
    lDown.X2 = ScaleWidth - 50
    lDown.Y1 = iLine
    lDown.Y2 = iLine
    lLeft.Y1 = iLine
    lLeft.Y2 = iLine - 40
    lRight.X1 = ScaleWidth - 60
    lRight.X2 = ScaleWidth - 60
    lRight.Y1 = iLine
    lRight.Y2 = iLine - 40
    Height = 515
    lstListData.Move 0, 0, UserControl.ScaleWidth
    picDrop.Move lDown.X2 - 16 * Screen.TwipsPerPixelX, lDown.Y2 - 16 * Screen.TwipsPerPixelY, 16 * Screen.TwipsPerPixelX, 16 * Screen.TwipsPerPixelY
    DrawComboArrow
    Text1.Width = ScaleWidth - (picClearText.Width + picDrop.Width * m_control_type + 130 + Text1.Left)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty(PRP_BCOLOR, m_bcolor, DEF_BCOLOR)
        Call .WriteProperty(PRP_FCOLOR, m_fcolor, DEF_FCOLOR)
        Call .WriteProperty(PRP_CAPTION, m_caption, DEF_CAPTION)
        Call .WriteProperty(PRP_CAPTIONCOLOR, m_caption_color, DEF_CAPTIONCOLOR)
        Call .WriteProperty(PRP_TEXT, m_text, DEF_TEXT)
        Call .WriteProperty(PRP_LOCKED, m_locked, DEF_LOCKED)
        Call .WriteProperty(PRP_UNDERLINECOLOR, m_uncerline_color, DEF_UNDERLINECOLOR)
        Call .WriteProperty(PRP_PASSWORD, m_password, DEF_PASSWORD)
        Call .WriteProperty(PRP_MINPASS, m_min_pass, DEF_MINPASS)
        Call .WriteProperty(PRP_TIPS, m_tips, DEF_TIPS)
        Call .WriteProperty(PRP_REQUIRED, m_required, DEF_REQUIRED)
        Call .WriteProperty(PRP_ENABLED, m_enabled, DEF_ENABLED)
        Call .WriteProperty(PRP_COLUMNS, Me.Columns, DEF_COLUMNS)
        Call .WriteProperty(PRP_MAXROWS, m_maxRows, DEF_MAXROWS)
        Call .WriteProperty(PRP_TYPE, m_control_type, DEF_TYPE)
        Call .WriteProperty(PRP_COMBOSTYLE, m_combo_style, DEF_COMBOSTYLE)
        Call .WriteProperty(PRP_SORTED, m_sorted, DEF_SORTED)
        Call .WriteProperty(PRP_PINNED, m_pinned, DEF_PINNED)
    End With

End Sub

Private Function IsComplex(sPass As String) As String

    Dim iStrength As Integer

    iStrength = 0

    If Len(sPass) >= m_min_pass Then iStrength = iStrength + 1

    If CheckValue(97, 122, sPass) Then iStrength = iStrength + 1

    If CheckValue(65, 90, sPass) Then iStrength = iStrength + 1

    If CheckValue(48, 57, sPass) Then iStrength = iStrength + 1

    If CheckValue(33, 47, sPass) Then iStrength = iStrength + 1
    If CheckValue(58, 64, sPass) Then iStrength = iStrength + 1
    If CheckValue(91, 96, sPass) Then iStrength = iStrength + 1
    If CheckValue(123, 255, sPass) Then iStrength = iStrength + 1

    Select Case iStrength
        Case 0, 1, 2
            IsComplex = "Weak"
        Case 3
            IsComplex = "Moderate"
        Case 4
            IsComplex = "Strong"
        Case 5, 6, 7, 8
            IsComplex = "Very Strong"
    End Select

    sPass = ""
End Function

Private Function CheckValue(x, y, sPass) As Boolean

    Dim iLoopVar As Integer

    iLoopVar = 0
    CheckValue = False

    For iLoopVar = x To y

        If InStr(1, sPass, Chr$(iLoopVar)) > 0 Then
            CheckValue = True
        End If

    Next

End Function

Public Property Get Columns() As Long
Attribute Columns.VB_ProcData.VB_Invoke_Property = "ControlSettings"
    Columns = lstListData.ColumnHeaders.count
End Property

Public Property Let Columns(ByVal vNewValue As Long)

    Dim K As Long, count As Long

    If lstListData.ColumnHeaders.count = 0 Then ' add columns

        For K = 1 To vNewValue
            lstListData.ColumnHeaders.Add , , Chr$(65 + K)
        Next K

    ElseIf lstListData.ColumnHeaders.count < vNewValue Then ' add more columns
        count = vNewValue - lstListData.ColumnHeaders.count

        For K = 1 To count
            lstListData.ColumnHeaders.Add , , ""
        Next K

    ElseIf lstListData.ColumnHeaders.count > vNewValue Then ' too many, so remove columns
        count = lstListData.ColumnHeaders.count - vNewValue

        For K = 1 To count
            lstListData.ColumnHeaders.Remove lstListData.ColumnHeaders.count
        Next K

    End If

End Property

Private Sub DrawComboArrow()

    Dim PentaPoints(1 To 3) As POINTAPI

    PentaPoints(1).x = 8
    PentaPoints(1).y = 16
    PentaPoints(2).x = 16
    PentaPoints(2).y = 16
    PentaPoints(3).x = 16
    PentaPoints(3).y = 8
    picDrop.FillStyle = vbSolid
    picDrop.BackColor = Ambient.BackColor

    If GetPolyFillMode(picDrop.hDC) <> WINDING Then SetPolyFillMode picDrop.hDC, WINDING
    Polygon picDrop.hDC, PentaPoints(1), 3
    SetPolyFillMode picDrop.hDC, ALTERNATE
End Sub

Private Function ProperCase(str As String) As String

    Dim i             As Integer
    Dim char          As String
    Dim blnAfterSpace As Boolean

    For i = 1 To Len(str)
        char = Mid$(str, i, 1)

        If blnAfterSpace = True Then
            char = UCase$(char)
            blnAfterSpace = False
        End If

        If char = " " Then blnAfterSpace = True
        If i = 1 Then char = UCase$(char)
        ProperCase = ProperCase & char
    Next i

End Function
