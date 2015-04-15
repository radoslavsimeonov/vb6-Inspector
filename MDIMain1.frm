VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CE08B9D4-381D-4D40-B699-E8352BA50128}#1.1#0"; "VBCCR10.OCX"
Begin VB.MDIForm MDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Inspector"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15510
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPlaceHolder 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9060
      Left            =   0
      ScaleHeight     =   9060
      ScaleWidth      =   3465
      TabIndex        =   2
      Top             =   1095
      Width           =   3465
      Begin VBCCR10.TreeView tvwFeatures 
         Height          =   2295
         Left            =   360
         TabIndex        =   8
         Top             =   3240
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   4048
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "MDIMain1.frx":0000
         PathSeparator   =   "MDIMain1.frx":002C
      End
      Begin VBCCR10.CommandButtonW CommandButtonW1 
         Height          =   495
         Left            =   720
         TabIndex        =   7
         Top             =   6060
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "MDIMain1.frx":004E
      End
      Begin MSComctlLib.ImageList ilMenuItems 
         Left            =   720
         Top             =   6840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":008C
               Key             =   "APPLICATIONS"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":0966
               Key             =   "EXIT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":6158
               Key             =   "HARDWARE"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":6A32
               Key             =   "DEVMGMT"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":730C
               Key             =   "SERVICES"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picSlider 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   6  'Cross
         ForeColor       =   &H00FFFFFF&
         Height          =   8580
         Left            =   3120
         MouseIcon       =   "MDIMain1.frx":7BE6
         MousePointer    =   99  'Custom
         ScaleHeight     =   8580
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   0
         Width           =   375
         Begin VB.Label lblSliderDown 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "< < < < <"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   1215
            Left            =   90
            MouseIcon       =   "MDIMain1.frx":7D38
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   6900
            Width           =   255
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSliderUp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "< < < < <"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   1215
            Left            =   90
            MouseIcon       =   "MDIMain1.frx":7E8A
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   0
            Width           =   255
            WordWrap        =   -1  'True
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2040
         Top             =   6840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   19
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":7FDC
               Key             =   "SHARED"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":88B6
               Key             =   "DEVMGR"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":9190
               Key             =   "REPLACE"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":9A6A
               Key             =   "EXIT1"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":A344
               Key             =   "EXIT"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":AC1E
               Key             =   "NAMING"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":B4F8
               Key             =   "SERVICES1"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":B812
               Key             =   "SERVICES"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":C0EC
               Key             =   "APPLICATIONS"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":C9C6
               Key             =   "NETWORK"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":D2A0
               Key             =   "SYSTEM"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":DB7A
               Key             =   "STARTUP"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":14E84
               Key             =   "SECURITY"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":1519E
               Key             =   "HARDWARE"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":15A78
               Key             =   "DEVMGR1"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":15ECA
               Key             =   "EVENTVWR"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":161E4
               Key             =   "USERS"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":16ABE
               Key             =   "APPWIZ"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain1.frx":17398
               Key             =   "SECPOL"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwFeatures1 
         Height          =   2535
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4471
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   1
         FullRowSelect   =   -1  'True
         Scroll          =   0   'False
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      BorderStyle     =   0  'None
      ForeColor       =   &H00535353&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   15510
      TabIndex        =   0
      Top             =   0
      Width           =   15510
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "BDUK Inspector"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   555
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   3915
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   480
         Picture         =   "MDIMain1.frx":177EA
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const F_DEVMGMT      As String = "DEVICE MANAGER"
Private Const F_SERVICES     As String = "SERVICES"
Private Const F_APPLICATIONS As String = "APPLICATIONS"
Private Const F_HARDWARE     As String = "HARDWARE"

'Private Const F_ As String = ""
Private Const F_EXIT         As String = "EXIT"

Private Sub lblSliderDown_Click()
    MinMaxWindow
End Sub

Private Sub lblSliderUp_Click()
    MinMaxWindow
End Sub

Private Sub MDIForm_Load()

    Dim nodX  As VBCCR10.TvwNode
    'Dim nodX  As Node
    Dim hMenu As Long
    Dim i     As Integer

    DoUnload = True
    'Disable form resize and remove MAX button
    'hMenu = GetSystemMenu(Me.hWnd, 0)
    'Call DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)
    'EnableMaxButton MDIMain.hwnd, False
    'SetTVBackColor tvwFeatures, &H535353
    'tvwFeatures.ImageList = ilsIcons
    tvwFeatures.ImageList = ilMenuItems
    
    Set nodX = tvwFeatures.Nodes.Add(, , F_APPLICATIONS, F_APPLICATIONS, 1)
    Set nodX = tvwFeatures.Nodes.Add(, , F_DEVMGMT, F_DEVMGMT, 2)
    Set nodX = tvwFeatures.Nodes.Add(, , F_SERVICES, F_SERVICES, 3)
    Set nodX = tvwFeatures.Nodes.Add(, , F_HARDWARE, F_HARDWARE, 4)

    Set nodX = tvwFeatures.Nodes.Add(, , F_EXIT, F_EXIT, 5)
'    Set nodX = tvwFeatures.Nodes.Add(, , F_APPLICATIONS, F_APPLICATIONS, "APPLICATIONS")
'    Set nodX = tvwFeatures.Nodes.Add(, , F_DEVMGMT, F_DEVMGMT, "DEVMGMT")
'    Set nodX = tvwFeatures.Nodes.Add(, , F_SERVICES, F_SERVICES, "SERVICES")
'    Set nodX = tvwFeatures.Nodes.Add(, , F_HARDWARE, F_HARDWARE, "HARDWARE")
'
'    Set nodX = tvwFeatures.Nodes.Add(, , F_EXIT, F_EXIT, "EXIT")

    For Each nodX In tvwFeatures.Nodes
        nodX.BackColor = &H535353
        nodX.ForeColor = vbWhite ' &H80FF&
    Next

    tvwFeatures.Nodes("EXIT").ForeColor = vbRed
    'HookForm Me.hwnd
End Sub

Private Sub MDIForm_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    picSlider.Height = Me.ScaleHeight
    DrawGradient picSlider, -0, &H535353, &H404040
    lblSliderDown.Top = Me.ScaleHeight - lblSliderDown.Height
    'Draw vetical text in slider
    DrawRotatedText picSlider, "Меню", picSlider.ScaleWidth / 3 - 150, picSlider.ScaleHeight / 2 + 1400 - picHeader.Height, "Courier", 12, 90, 700, False, False, False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    '    If MUST_RESTART = True Then
    '        If MsgBox("Трябва да рестартирате компютъра," & vbCrLf & "за да приложат новите настройки." & '            vbCrLf & "Да рестртирам ли сега ?", vbYesNo + vbExclamation, Me.Caption) = vbYes Then
    '                ExitWindowsWith "reboot"
    '        End If
    '    End If
    'UnhookForm
   Unload MDIMain
    'Unload frmUsersProfiles
End Sub

Sub picSlider_Click()
    MinMaxWindow
End Sub

Private Sub tvwFeatures_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    Dim itmX As VBCCR10.TvwNode
'    Dim itmX As Node
    Dim idx  As Integer

    Set itmX = tvwFeatures.HitTest(X, Y)

    If Not itmX Is Nothing And tvwFeatures.Width > picPlaceHolder.Width Then
        tvwFeatures.ToolTipText = itmX.Text
    Else
        tvwFeatures.ToolTipText = ""
    End If

End Sub

Private Sub MaxForm(max As Boolean)

    If (MDIMain.picPlaceHolder.Width > 1000) = max Then MinMaxWindow
End Sub

Private Sub tvwFeatures_NodeClick(ByVal Node As VBCCR10.TvwNode, ByVal Button As Integer)
'Private Sub tvwFeatures_NodeClick(ByVal Node As Node)

    If Not MDIMain.ActiveForm Is Nothing Then
        If Node.key <> MDIMain.ActiveForm.Tag Then Unload MDIMain.ActiveForm
        If DoUnload = False Then
            tvwFeatures.Nodes(MDIMain.ActiveForm.Tag).Selected = True
            Exit Sub
        End If
    End If

    Select Case Node.key
            '        Case "SECPOL"
            '            Load frmPolicy
            '            MaxForm False
        Case F_APPLICATIONS
            Load frmApplications
        Case "HARDWARE"
            Load frmHardware
            '        Case "EVENTVWR"
            '            Load frmEvents
            '            MaxForm False
        Case F_DEVMGMT
            Load frmDeviceManager
        Case F_SERVICES
            Load frmServices
            '        Case "REPLACE"
            '            Load frmReplace
            '            MaxForm True
            '        Case "NAMING"
            '            Load frmNaming
            '            MaxForm True
        Case "EXIT"
            Me.Hide
            MDIForm_Unload True
    End Select

End Sub


