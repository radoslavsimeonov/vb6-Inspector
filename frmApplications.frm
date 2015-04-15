VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmApplications 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Installed programs"
   ClientHeight    =   3360
   ClientLeft      =   6015
   ClientTop       =   450
   ClientWidth     =   11250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   11250
   Tag             =   "SOFTWARE"
   WindowState     =   2  'Maximized
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
      ScaleWidth      =   11250
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   11250
      Begin VB.PictureBox picMenuHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   700
         Left            =   0
         ScaleHeight     =   705
         ScaleWidth      =   10575
         TabIndex        =   2
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
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   9000
            MouseIcon       =   "frmApplications.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   8
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
            MouseIcon       =   "frmApplications.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   7
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
            MouseIcon       =   "frmApplications.frx":02A4
            MousePointer    =   99  'Custom
            TabIndex        =   6
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
            MouseIcon       =   "frmApplications.frx":03F6
            MousePointer    =   99  'Custom
            TabIndex        =   5
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
            MouseIcon       =   "frmApplications.frx":0548
            MousePointer    =   99  'Custom
            TabIndex        =   4
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
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   2040
            MouseIcon       =   "frmApplications.frx":069A
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
      End
   End
   Begin MSComctlLib.ImageList ilApplications 
      Left            =   2880
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmApplications.frx":07EC
            Key             =   "APPLICATION"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPrograms 
      Height          =   2025
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   3572
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilApplications"
      ForeColor       =   16777215
      BackColor       =   5395026
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   240
      X2              =   240
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Menu mnuApplications 
      Caption         =   "Uninstall"
      Visible         =   0   'False
      Begin VB.Menu mnuUninstall 
         Caption         =   "Uninstall"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuSepartator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmApplications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_hookedLV As Boolean

Private Sub FillListViewValues(Optional bRefresh As Boolean = False)

    Dim i    As Integer
    Dim itmX As ListItem

    If bRefresh Then SW_APPLICATIONS = EnumApplications
    lvwPrograms.ListItems.Clear

    For i = 0 To UBound(SW_APPLICATIONS)
        With SW_APPLICATIONS(i)
            If .AppName <> vbNullString Then
                Set itmX = lvwPrograms.ListItems.Add(, "K" & str$(i), .AppName, , "APPLICATION")
                itmX.SubItems(1) = Trim$(.Version)
                itmX.ListSubItems.Add(Text:=.InstalledOn).Tag = .InstalledOn
                itmX.SubItems(3) = .Publisher
            End If
        End With
    Next

    mnuUninstall.Enabled = False
    
    With lvwPrograms
        If .ListItems.count > 0 Then
            .ListItems(.GetFirstVisible.Index).Selected = True
            .ListItems(.GetFirstVisible.Index).EnsureVisible
        End If
    End With
End Sub

Private Sub Form_Load()

    With lvwPrograms
        .ColumnHeaders.Add , , "Name", 4800
        .ColumnHeaders.Add , , "Version", 1600
        .ColumnHeaders.Add , , "Install on", 1400
        .ColumnHeaders.Add , , "Publisher", 2200
        .View = lvwReport
        .Sorted = True
    End With

    DoEvents
    FillListViewValues
End Sub

Private Sub Form_Resize()
On Error Resume Next

    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2

    If Me.WindowState = vbMinimized Then Exit Sub
    Static rc As RECT

    With lvwPrograms
        .Visible = False
        
        .Move L_TABLE_LEFT, _
              700, _
              Me.ScaleWidth - 2 * L_TABLE_LEFT, _
              Me.ScaleHeight - 700 - 120
        
        .ColumnHeaders(1).Width = .Width / 2
        .ColumnHeaders(2).Width = .Width / 7
        .ColumnHeaders(3).Width = .Width / 7
        .ColumnHeaders(4).Width = .Width / 5 - 200
        
        .Visible = True
    End With
    
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

Private Sub lblHeaderStartUp_Click()
    frmStartUp.ZOrder 0
    frmStartUp.Show
End Sub

Private Sub lvwPrograms_Click()
    
    If lvwPrograms.SelectedItem Is Nothing Then Exit Sub
    
    If lvwPrograms.SelectedItem.Index <> -1 Then mnuUninstall.Enabled = True
End Sub

Private Sub lvwPrograms_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim i         As Long
    Dim oListItem As ListItem

    If ColumnHeader.Index = 2 Then Exit Sub

    For i = 1 To lvwPrograms.ColumnHeaders.count

        If (i <> ColumnHeader.Index) Then
            lvwPrograms.ColumnHeaders(i).Tag = ""
        End If

    Next i

    If ColumnHeader.Index = 3 Then

        For Each oListItem In lvwPrograms.ListItems
            oListItem.SubItems(2) = Format$(oListItem.ListSubItems(2).Tag, "yyyymmddHHMMSS")
        Next oListItem

    End If

    If ColumnHeader.Tag = "ASC" Then
        ColumnHeader.Tag = "DESC"
        lvwPrograms.SortOrder = lvwDescending
    Else
        ColumnHeader.Tag = "ASC"
        lvwPrograms.SortOrder = lvwAscending
    End If

    lvwPrograms.SortKey = ColumnHeader.Index - 1
    lvwPrograms.Sorted = True

    If ColumnHeader.Index = 3 Then

        For Each oListItem In lvwPrograms.ListItems
            oListItem.SubItems(2) = oListItem.ListSubItems(2).Tag
        Next oListItem

    End If

End Sub

Private Sub lvwPrograms_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 Then
        FillListViewValues True
    ElseIf KeyCode = vbKeyDelete And Shift = 1 Then
        mnuUninstall_Click
    End If

End Sub

Private Sub lvwPrograms_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        If lvwPrograms.HitTest(x, y) Is Nothing Then
            mnuUninstall.Visible = False
            mnuSepartator.Visible = False
        Else
            mnuUninstall.Visible = True
            mnuSepartator.Visible = True
        End If
        PopupMenu mnuApplications
    End If
    
End Sub

Private Sub mnuRefresh_Click()
    FillListViewValues True
End Sub

Private Sub mnuUninstall_Click()
    Dim sKey As String

    Screen.MousePointer = vbHourglass

    mnuUninstall.Enabled = False
    sKey = lvwPrograms.SelectedItem.Key

    With SW_APPLICATIONS(Int(Mid$(sKey, 2, Len(sKey))))

        If .ModifyString <> "" Then
            ShellWait .ModifyString, vbNormalFocus
        Else
            ShellWait .UninstallString, vbNormalFocus
        End If

    End With

    mnuUninstall.Enabled = True
    FillListViewValues True
    
    Screen.MousePointer = vbNormal
End Sub
