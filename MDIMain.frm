VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Inspector"
   ClientHeight    =   10350
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   15510
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPlaceHolder 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      ScaleHeight     =   9255
      ScaleWidth      =   3465
      TabIndex        =   2
      Top             =   1095
      Width           =   3465
      Begin MSComctlLib.ImageList ilMenuItems 
         Left            =   2400
         Top             =   6600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":0000
               Key             =   "APPLICATIONS"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":08DA
               Key             =   "SAVEEXIT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1CBC
               Key             =   "SECPOL"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":23B6
               Key             =   "NETWORK"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":2C90
               Key             =   "OS"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":356A
               Key             =   "USERDETAILS1"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":3E44
               Key             =   "USERDETAILS"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":471E
               Key             =   "EXIT"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":9F10
               Key             =   "HARDWARE"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":A7EA
               Key             =   "DEVMGMT"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":B0C4
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
         MouseIcon       =   "MDIMain.frx":B99E
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
            MouseIcon       =   "MDIMain.frx":BAF0
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
            MouseIcon       =   "MDIMain.frx":BC42
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   0
            Width           =   255
            WordWrap        =   -1  'True
         End
      End
      Begin MSComctlLib.TreeView tvwFeatures 
         Height          =   6135
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   10821
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
      Begin VB.Frame fraExit 
         BackColor       =   &H00535353&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         MouseIcon       =   "MDIMain.frx":BD94
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   7800
         Width           =   2895
         Begin VB.Label lblExitMenu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EXIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   600
            TabIndex        =   8
            Top             =   120
            Width           =   510
         End
         Begin VB.Image imgExit 
            Height          =   480
            Left            =   0
            Picture         =   "MDIMain.frx":BEE6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
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
         Caption         =   "Inspector"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   540
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.Menu mnuInputLanguage 
      Caption         =   "ац"
      Visible         =   0   'False
      Begin VB.Menu mnuSwitchInput 
         Caption         =   "SwitchLanguage"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const F_WORKSTATION  As String = "WORKSTSTION"
Private Const F_OS           As String = "SOFTWARE"
Private Const F_HARDWARE     As String = "HARDWARE"
Private Const F_NETWORK      As String = "NETWORK"
Private Const F_POLICY       As String = "SECURITY"

Private Const F_SAVEEXIT     As String = "SAVEEXIT"

Private Sub fraExit_Click()
    Unload Me
End Sub

Private Sub imgExit_Click()
    Unload Me
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picPlaceHolder.Width < 1001 Then
        imgExit.ToolTipText = lblExitMenu.Caption
    Else
        imgExit.ToolTipText = ""
    End If
End Sub

Private Sub lblExitMenu_Click()
    Unload Me
End Sub

Private Sub lblSliderDown_Click()
    MinMaxWindow
End Sub

Private Sub lblSliderUp_Click()
    MinMaxWindow
End Sub

Private Sub MDIForm_Load()

    Dim nodX  As Node
    Dim hMenu As Long
    Dim i     As Integer

    DoUnload = True
    SetTVBackColor tvwFeatures, &H535353
    tvwFeatures.ImageList = ilMenuItems
    
    Set nodX = tvwFeatures.Nodes.Add(, , F_WORKSTATION, "WORKSTATION", "USERDETAILS")
    Set nodX = tvwFeatures.Nodes.Add(, , F_OS, "SOFTWARE", "OS")
    Set nodX = tvwFeatures.Nodes.Add(, , F_HARDWARE, "HARDWARE", "HARDWARE")
    Set nodX = tvwFeatures.Nodes.Add(, , F_NETWORK, "NETWORK", "NETWORK")
    Set nodX = tvwFeatures.Nodes.Add(, , F_POLICY, "LOCAL POLICY (beta)", "SECPOL")

    Set nodX = tvwFeatures.Nodes.Add(, , F_SAVEEXIT, "SAVE & EXIT", "SAVEEXIT")

    For Each nodX In tvwFeatures.Nodes
        nodX.BackColor = &H535353
        nodX.ForeColor = vbWhite ' &H80FF&
    Next

    tvwFeatures.Nodes("SAVEEXIT").ForeColor = &HC0C000
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub
    picSlider.Height = Me.ScaleHeight
    DrawGradient picSlider, -0, &H535353, &H404040
    lblSliderDown.Top = Me.ScaleHeight - lblSliderDown.Height
    DrawRotatedText picSlider, "Menu", picSlider.ScaleWidth / 3 - 150, picSlider.ScaleHeight / 2 + 1400 - picHeader.Height, "Courier", 12, 90, 700, False, False, False
    fraExit.Top = Me.ScaleHeight - fraExit.Height - 480

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Unload MDIMain
End Sub

Private Sub mnuSwitchInput_Click()
    SwitchInputLanguages
End Sub

Sub picSlider_Click()
    MinMaxWindow
End Sub

Private Sub MaxForm(max As Boolean)

    If (MDIMain.picPlaceHolder.Width > 1000) = max Then MinMaxWindow
End Sub

Private Sub tvwFeatures_NodeClick(ByVal Node As Node)

    If Not MDIMain.ActiveForm Is Nothing Then
        If DoUnload = False Then
            tvwFeatures.Nodes(MDIMain.ActiveForm.Tag).Selected = True
            Exit Sub
        End If
        If Node.Key = MDIMain.ActiveForm.Tag Then Exit Sub  'Unload MDIMain.ActiveForm
    End If
    
    Select Case Node.Key
        Case F_WORKSTATION
            frmUserDetails.ZOrder 0
            Load frmUserDetails
        Case F_OS
            frmOS.ZOrder 0
            Load frmOS
        Case F_HARDWARE
            frmHardware.ZOrder 0
            Load frmHardware
        Case F_NETWORK
            frmNetwork.ZOrder 0
            Load frmNetwork
        Case F_POLICY
            frmPolicy.ZOrder 0
            Load frmPolicy
        Case F_SAVEEXIT
            If ValidateWorkstationDetails Then
                If ValidateHDDDetails Then
                    If ValidateUserDetails Then
                        If ExportData(AsXML) Then Unload Me
                    End If
                End If
            End If
            
    End Select
    
End Sub


