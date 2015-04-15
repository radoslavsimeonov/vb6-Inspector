VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartUp 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Startup"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12555
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   Tag             =   "SOFTWARE"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picDefault 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      Picture         =   "frmStartUp.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ilStartup 
      Left            =   6000
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picSoftwareMenu 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   700
      Left            =   0
      ScaleHeight     =   700
      ScaleMode       =   0  'User
      ScaleWidth      =   12555
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12555
      Begin VB.PictureBox picMenuHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   700
         Left            =   0
         ScaleHeight     =   705
         ScaleWidth      =   10575
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   10575
         Begin VB.Label lblHeaderStartUp 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STARTUP  ITEMS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   9000
            MouseIcon       =   "frmStartUp.frx":058A
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   120
            Width           =   1560
            WordWrap        =   -1  'True
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   5
            X1              =   8760
            X2              =   8760
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   4
            X1              =   6840
            X2              =   6840
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   5160
            X2              =   5160
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   3720
            X2              =   3720
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   1800
            X2              =   1800
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Label lblHeaderLocalUsers 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "USERS AND GROUPS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   7080
            MouseIcon       =   "frmStartUp.frx":06DC
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblHeaderSharedFolders 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SHARED FOLDERS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   5280
            MouseIcon       =   "frmStartUp.frx":082E
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblHeaderServices 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NT SERVICES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   4005
            MouseIcon       =   "frmStartUp.frx":0980
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   120
            Width           =   930
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblHeaderOS 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SOFTWARE SUMMARY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   120
            MouseIcon       =   "frmStartUp.frx":0AD2
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblHeaderApps 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INSTALLED SOFTWARE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   2040
            MouseIcon       =   "frmStartUp.frx":0C24
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
      End
   End
   Begin MSComctlLib.ListView lvStartUp 
      Height          =   3105
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5477
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilStartup"
      ForeColor       =   16777215
      BackColor       =   5395026
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Menu mnuStartup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "New startup"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit startup"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete startup"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next

    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2
    
    With lvStartUp
        .Move 120, _
              700, _
              Me.ScaleWidth - 2 * 120, _
              Me.ScaleHeight - 700 - 120
    
        .ColumnHeaders(1).Width = .Width \ 6
        .ColumnHeaders(2).Width = .Width \ 2
        .ColumnHeaders(3).Width = .Width \ 6
        .ColumnHeaders(4).Width = .Width \ 6
    End With
End Sub

Private Sub Form_Load()
    
    With lvStartUp
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Commannd"
        .ColumnHeaders.Add , , "Publisher"
        .ColumnHeaders.Add , , "User"
        .ColumnHeaders.Add , , "Location"
                
        .View = lvwReport
        .HideColumnHeaders = False
    End With
    
    FillListViewStartUps
End Sub

Private Sub FillListViewStartUps(Optional bRefresh As Boolean = False)

    Dim cnt     As Integer
    Dim itmX    As ListItem
    Dim tPath   As String
    
    If bRefresh Then SW_START_COMMANDS = EnumStartUpCommands
    
    With lvStartUp
        .Visible = False
        .ListItems.Clear
        .SmallIcons = Nothing
        ilStartup.ListImages.Clear
        ilStartup.ListImages.Add , , picDefault.Picture
        .SmallIcons = ilStartup
    End With

    For cnt = 0 To UBound(SW_START_COMMANDS)
        With SW_START_COMMANDS(cnt)
            If .Command <> vbNullString Then
                tPath = GetFilePathWithOutParams(.Command)
                
                ilStartup.ListImages.Add , "K" & CStr(cnt), _
                            GetAssocIcon(tPath, False, True)
                
                Set itmX = lvStartUp.ListItems.Add(, , .CommandName, , _
                           "K" & CStr(cnt))
                itmX.SubItems(1) = .Command
                itmX.SubItems(2) = .Vendor
                itmX.SubItems(3) = .UserRange
                itmX.SubItems(4) = .Location
            End If
        End With
    Next cnt
    
    lvStartUp.Visible = True
End Sub

Private Sub lblHeaderApps_Click()
    frmApplications.ZOrder 0
    frmApplications.Show
End Sub

Private Sub lblHeaderLocalUsers_Click()
    frmLocalUsers.ZOrder 0
    frmLocalUsers.Show
End Sub

Private Sub lblHeaderOS_Click()
    frmOS.ZOrder 0
    frmOS.Show
End Sub

Private Sub lblHeaderServices_Click()
    frmServices.ZOrder 0
    frmServices.Show
End Sub

Private Sub lblHeaderSharedFolders_Click()
    frmSharedFolders.ZOrder 0
    frmSharedFolders.Show
End Sub

Private Sub lvStartUp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        FillListViewStartUps True
        lvStartUp.SetFocus
    End If
End Sub

Private Sub lvStartUp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
       
    If Button = vbRightButton Then
        If lvStartUp.HitTest(x, y) Is Nothing Then
            mnuEdit.Visible = False
            mnuDelete.Visible = False
        Else
            mnuEdit.Visible = True
            mnuDelete.Visible = True
        End If
        PopupMenu mnuStartup
    End If
End Sub

Private Sub mnuDelete_Click()
    If lvStartUp.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("Delete following startup item: " & lvStartUp.SelectedItem.Text & " ?", _
                vbQuestion + vbYesNo, "Delete startup item") = vbNo Then Exit Sub
    
    DeleteStartupItem SW_START_COMMANDS(lvStartUp.SelectedItem.Index - 1)
    
    FillListViewStartUps True
End Sub

Private Sub mnuRefresh_Click()
    FillListViewStartUps True
    lvStartUp.SetFocus
End Sub
