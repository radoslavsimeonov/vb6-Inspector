VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocalUsers 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Local users and groups"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   12090
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
      ScaleWidth      =   12090
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   12090
      Begin VB.PictureBox picMenuHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   700
         Left            =   0
         ScaleHeight     =   705
         ScaleWidth      =   10575
         TabIndex        =   20
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
            MouseIcon       =   "frmUsers.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   26
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
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   7080
            MouseIcon       =   "frmUsers.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   25
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
            MouseIcon       =   "frmUsers.frx":02A4
            MousePointer    =   99  'Custom
            TabIndex        =   24
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
            MouseIcon       =   "frmUsers.frx":03F6
            MousePointer    =   99  'Custom
            TabIndex        =   23
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
            MouseIcon       =   "frmUsers.frx":0548
            MousePointer    =   99  'Custom
            TabIndex        =   22
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
            MouseIcon       =   "frmUsers.frx":069A
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
      End
   End
   Begin MSComctlLib.ImageList ilUsers 
      Left            =   11280
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   255
      ImageWidth      =   24
      ImageHeight     =   24
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":07EC
            Key             =   "ADMIN1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":5FDE
            Key             =   "ADMIN0"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":B7D0
            Key             =   "USER1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":10FC2
            Key             =   "USER0"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilGroups 
      Left            =   10560
      Top             =   4800
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
            Picture         =   "frmUsers.frx":167B4
            Key             =   "GROUP"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameGroups 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Local groups"
      ForeColor       =   &H0000FF00&
      Height          =   4095
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   11775
      Begin VB.PictureBox picGroupButtons 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10320
         ScaleHeight     =   255
         ScaleWidth      =   1215
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1215
         Begin VB.Image imgDeleteMembersGroup 
            Height          =   240
            Left            =   960
            MouseIcon       =   "frmUsers.frx":16D4E
            MousePointer    =   99  'Custom
            Picture         =   "frmUsers.frx":16EA0
            ToolTipText     =   "Delete group (key: Del)"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgAddMembersGroup 
            Height          =   240
            Left            =   0
            MouseIcon       =   "frmUsers.frx":18272
            MousePointer    =   99  'Custom
            Picture         =   "frmUsers.frx":183C4
            ToolTipText     =   "Add group (key: Ins)"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgEditMemberGroup 
            Height          =   240
            Left            =   480
            MouseIcon       =   "frmUsers.frx":19796
            MousePointer    =   99  'Custom
            Picture         =   "frmUsers.frx":198E8
            ToolTipText     =   "Edit group (key: Enter)"
            Top             =   0
            Width           =   240
         End
      End
      Begin MSComctlLib.ListView lvGroups 
         Height          =   2055
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilGroups"
         ForeColor       =   16777215
         BackColor       =   5395026
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame frameLocalUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Local users"
      ForeColor       =   &H0000FF00&
      Height          =   4215
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   11775
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "&Save"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CheckBox chkMustChangePassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "User must change password "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   7
         Top             =   1200
         Width           =   2640
      End
      Begin VB.CheckBox chkPasswordNeverExpire 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Password never expires"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   9
         Top             =   1920
         Width           =   2520
      End
      Begin VB.CheckBox chkCanChangePassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "User cannot change password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   8
         Top             =   1560
         Width           =   2880
      End
      Begin VB.CheckBox chkAccountLocked 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Account is locked out"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   11
         Top             =   2640
         Width           =   1920
      End
      Begin VB.CheckBox chkAccountDisabled 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Account is disabled"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   10
         Top             =   2280
         Width           =   2280
      End
      Begin VB.ListBox lstGroups 
         Appearance      =   0  'Flat
         BackColor       =   &H00525252&
         ForeColor       =   &H00FFFFFF&
         Height          =   810
         ItemData        =   "frmUsers.frx":1ACBA
         Left            =   3360
         List            =   "frmUsers.frx":1ACBC
         TabIndex        =   6
         Top             =   2880
         Width           =   3885
      End
      Begin Inspector.TextBoxEx lblLastLogon 
         Height          =   510
         Left            =   7800
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   900
         Caption         =   "Last logon"
         Locked          =   -1  'True
      End
      Begin Inspector.TextBoxEx txtPassword1 
         Height          =   510
         Left            =   3360
         TabIndex        =   4
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   900
         Caption         =   "Password"
         Password        =   -1  'True
      End
      Begin Inspector.TextBoxEx txtDescription 
         Height          =   510
         Left            =   3360
         TabIndex        =   3
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   900
         Caption         =   "Description"
      End
      Begin Inspector.TextBoxEx txtFullUserName 
         Height          =   510
         Left            =   3360
         TabIndex        =   2
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   900
         Caption         =   "Full name"
      End
      Begin Inspector.TextBoxEx txtUserName 
         Height          =   510
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   900
         Caption         =   "Username"
         Locked          =   -1  'True
      End
      Begin Inspector.TextBoxEx txtPassword2 
         Height          =   510
         Left            =   5400
         TabIndex        =   5
         Top             =   2040
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   900
         Caption         =   ""
         Password        =   -1  'True
      End
      Begin MSComctlLib.ListView lvwUsers 
         Height          =   3375
         Left            =   240
         TabIndex        =   0
         Top             =   320
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilUsers"
         SmallIcons      =   "ilUsers"
         ForeColor       =   16777215
         BackColor       =   5395026
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Image imgEditGroups 
         Height          =   240
         Left            =   6480
         MouseIcon       =   "frmUsers.frx":1ACBE
         MousePointer    =   99  'Custom
         Picture         =   "frmUsers.frx":1AE10
         ToolTipText     =   "Edit user's groups (key: Enter)"
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image imgDeleteUser 
         Height          =   240
         Left            =   2880
         MouseIcon       =   "frmUsers.frx":1C1E2
         MousePointer    =   99  'Custom
         Picture         =   "frmUsers.frx":1C334
         ToolTipText     =   "Delete user (key: Del)"
         Top             =   3800
         Width           =   240
      End
      Begin VB.Image imgAddUser 
         Height          =   240
         Left            =   2400
         MouseIcon       =   "frmUsers.frx":1D706
         MousePointer    =   99  'Custom
         Picture         =   "frmUsers.frx":1D858
         ToolTipText     =   "New user (key: Ins)"
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image imgDeleteGroup 
         Height          =   240
         Left            =   6960
         MouseIcon       =   "frmUsers.frx":1EC2A
         MousePointer    =   99  'Custom
         Picture         =   "frmUsers.frx":1ED7C
         ToolTipText     =   "Remove user from group (key: Del)"
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image imgRefresh 
         Height          =   225
         Left            =   240
         MouseIcon       =   "frmUsers.frx":2014E
         MousePointer    =   99  'Custom
         Picture         =   "frmUsers.frx":202A0
         Stretch         =   -1  'True
         ToolTipText     =   "Refresh (key: F5)"
         Top             =   3800
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Member of"
         ForeColor       =   &H000080FF&
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   16
         Top             =   2640
         Width           =   750
      End
   End
   Begin VB.Menu mnuGroups 
      Caption         =   "Groups"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add group"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit group"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete group"
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
Attribute VB_Name = "frmLocalUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsNew           As Boolean
Private bClicked    As Boolean

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


Private Sub chkAccountDisabled_Click()
    If bClicked = True Then cmdSave.Enabled = True
End Sub

Private Sub chkAccountLocked_Click()
    If bClicked = True Then cmdSave.Enabled = True
End Sub

Private Sub chkCanChangePassword_Click()
    chkMustChangePassword.Enabled = (Not chkCanChangePassword.Value = 1) _
                                    And (Not chkPasswordNeverExpire.Value = 1)
    If bClicked = True Then cmdSave.Enabled = True
End Sub

Private Sub chkMustChangePassword_Click()
    chkCanChangePassword.Enabled = chkMustChangePassword.Value < 1
    chkPasswordNeverExpire.Enabled = chkMustChangePassword.Value < 1
    If bClicked = True Then cmdSave.Enabled = True
End Sub

Private Sub chkPasswordNeverExpire_Click()
    chkMustChangePassword.Enabled = (Not chkCanChangePassword.Value = 1) _
                                    And (Not chkPasswordNeverExpire.Value = 1)
    If bClicked = True Then cmdSave.Enabled = True
End Sub

Private Sub Form_Load()
                       
    bClicked = False
                       
    lvwUsers.ColumnHeaders.Add , , "User"
    lvwUsers.ColumnHeaders(1).Width = lvwUsers.Width
    lvwUsers.View = lvwReport
    lvwUsers.HideSelection = False
    
    lvGroups.ColumnHeaders.Add , , "Group", 4000
    lvGroups.ColumnHeaders.Add , , "Description", lvGroups.Width - 4300
    lvGroups.View = lvwReport
    lvGroups.HideSelection = False
      
    FillUsersListView
    psubLocalGroupEnum
    
End Sub

Private Sub psubLocalGroupEnum()
    Dim lngRet As Long
    Dim lngCount As Long
    Dim i As Long
    Dim tLGI1VB() As LOCALGROUP_INFO_1_VB
    Dim itmX As ListItem
    Dim lngType As Long
    
    lvGroups.ListItems.Clear

    lngRet = swLocalGroups.LocalGroupEnum1(lngCount, tLGI1VB())
    If lngRet = 0 Then
        For i = 0 To lngCount - 1
            With tLGI1VB(i)
                Set itmX = lvGroups.ListItems.Add(, , .vb_lgrpi1_name, , "GROUP")
                itmX.SubItems(1) = .vb_lgrpi1_comment
            End With
        Next i
    Else
        MsgBox mErr.fncGetErrorString(lngRet), vbCritical
    End If
    
    With lvGroups
        If .ListItems.count > 0 Then
            .ListItems(.GetFirstVisible.Index).Selected = True
            .ListItems(.GetFirstVisible.Index).EnsureVisible
        End If
    End With
    
End Sub

Private Sub FillUsersListView()
    
    Dim i, x            As Integer
    Dim itmX            As ListItem
    Dim picType         As String
       
    lvwUsers.ListItems.Clear
    
    For x = 0 To UBound(SW_LOCAL_USERS)
        With SW_LOCAL_USERS(x)
            If .Name <> vbNullString Then
                picType = .Group & .AccountDisabled
                Set itmX = lvwUsers.ListItems.Add(, .Name, .Name, , picType)
                itmX.Tag = CStr(.Index)
            End If
        End With
    Next x
    
    Set itmX = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsNew = True Then
        If CancelNewUser(lvwUsers.ListItems.count - 1) = False Then
            DoUnload = False
            Cancel = 1
        Else
            DoUnload = True
        End If
    End If
End Sub

Private Sub imgAddGroup_Click()

End Sub

Private Sub Form_Resize()
On Error Resume Next
    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2

    frameLocalUsers.Width = Me.ScaleWidth - 2 * 120
    
    With frameGroups
        .Width = Me.ScaleWidth - 2 * 120
        .Height = Me.ScaleHeight - .Top - 120
    End With
    
    With lvGroups
        .Width = frameGroups.Width - 2 * 240
        .Height = frameGroups.Height - 3 * 240
        .ColumnHeaders(2).Width = lvGroups.Width - 4300
    End With
    
    With picGroupButtons
        .Left = frameGroups.Width - .Width - 240
        .Top = frameGroups.Height - .Height - 120
    End With

End Sub

Private Sub imgAddMembersGroup_Click()
             
    frmGroupMembers.MY_LocalGroupName = ""
    frmGroupMembers.MY_NewMode = True
    frmGroupMembers.Show vbModal

    Call psubLocalGroupEnum
    
    lvGroups.ListItems(1).Selected = True
    lvGroups.ListItems(1).EnsureVisible
End Sub

Private Sub imgAddUser_Click()
Dim itmX As ListItem
    
    If cmdSave.Enabled = True Then
        If MsgBox("User " & txtUserName.Text & " is modified." & vbCrLf & _
            "Do you want to save changes?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            cmdSave_Click
        End If
        cmdSave.Enabled = False
    End If
    
    IsNew = True
    Set itmX = lvwUsers.ListItems.Add(, "NewUser", "New user", , "USER0")
    itmX.Tag = "new"
    itmX.Selected = True
    
    ResetControlsValues
    
    imgAddUser.Visible = False
    imgDeleteUser.Visible = True
    txtUserName.SetFocus
End Sub

Private Sub imgDeleteGroup_Click()
    If lstGroups.ListIndex = -1 Then Exit Sub
    
    lstGroups.RemoveItem (lstGroups.ListIndex)
    cmdSave.Enabled = True
End Sub

Private Sub imgDeleteMembersGroup_Click()
    Dim lngRet As Long
    If lvGroups.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("Are you sure you want to delete group " & lvGroups.SelectedItem, vbExclamation + vbYesNo) = vbYes Then
        lngRet = swLocalGroups.LocalGroupDel(lvGroups.SelectedItem)
        If lngRet <> 0 Then
            MsgBox mErr.fncGetErrorString(lngRet)
        End If

        Call psubLocalGroupEnum
    End If
End Sub

Private Sub imgDeleteUser_Click()
Dim sUser As String
    
    If IsNew = True Then
        CancelNewUser lvwUsers.ListItems.count - 1
        Exit Sub
    End If
    
    If lvwUsers.SelectedItem Is Nothing Then Exit Sub
    
    sUser = lvwUsers.SelectedItem.Text
    
    If MsgBox("Are you sure you want to delete user " & sUser & "?", _
        vbExclamation + vbYesNo, Me.Caption) = vbYes Then
            If DeleteAccount(sUser) = True Then
                MsgBox "User was deleted successfully", vbInformation, Me.Caption
                
                lvwUsers.ListItems.Remove (lvwUsers.SelectedItem.Index)
                
                cmdSave.Enabled = False
                
                SW_LOCAL_USERS = EnumAccounts
                FillUsersListView
                
                lvwUsers_ItemClick lvwUsers.ListItems(1)
                lvwUsers.ListItems(1).Selected = True
            Else
                MsgBox "User WAS NOT deleted.", vbCritical, "FAILED"
        End If
    End If
End Sub

Private Sub lblAddUser_Click()
Dim itmX As ListItem
    
    If cmdSave.Enabled = True Then
        If MsgBox("User " & txtUserName.Text & " is modified." & vbCrLf & _
            "Do you want to save the changes?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            cmdSave_Click
        End If
        cmdSave.Enabled = False
    End If
    
    IsNew = True
    Set itmX = lvwUsers.ListItems.Add(, "NewUser", "New user", , "USER0")
    itmX.Tag = "new"
    itmX.Selected = True
    
    ResetControlsValues
    
    imgAddUser.Visible = False
    imgDeleteUser.Visible = True
    txtUserName.SetFocus
End Sub

Private Sub ResetControlsValues()
    bClicked = False
    
    txtUserName.Text = ""
    txtUserName.Locked = False
    txtFullUserName.Text = ""
    txtDescription.Text = ""
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtPassword1.UnderlineColor = vbRed
    txtPassword2.UnderlineColor = vbRed
    
    lstGroups.Clear
    lstGroups.AddItem "Users"
    
    lblLastLogon.Text = "Never"
        
    chkMustChangePassword.Value = 0
    chkCanChangePassword.Value = 0
    chkPasswordNeverExpire.Value = 0
    chkAccountDisabled.Value = 0
    chkAccountDisabled.Enabled = True
    chkAccountLocked.Value = 0
    chkAccountLocked.Enabled = False
    bClicked = True
End Sub

Private Sub lblRemoveGroup_Click()
    If lstGroups.ListIndex = -1 Then Exit Sub
    
    lstGroups.RemoveItem (lstGroups.ListIndex)
    cmdSave.Enabled = True
End Sub

Private Sub lblRemoveUser_Click()
    Dim sUser As String

    If IsNew = True Then
        CancelNewUser lvwUsers.ListItems.count - 1
        Exit Sub
    End If

    sUser = lvwUsers.SelectedItem.Text

    If MsgBox("Are you sure you want to delete user " & sUser & "?", vbExclamation + vbYesNo, Me.Caption) = vbYes Then

        If DeleteAccount(sUser) = True Then
            MsgBox "User was deleted successfully.", vbInformation, Me.Caption

            lvwUsers.ListItems.Remove (lvwUsers.SelectedItem.Index)

            cmdSave.Enabled = False

            lvwUsers_ItemClick lvwUsers.ListItems(1)
            lvwUsers.ListItems(1).Selected = True
        Else
            MsgBox "User WAS NOT deleted.", vbCritical
        End If
    End If

End Sub

Private Sub imgEditGroups_Click()
Dim sUsers As String
Dim aGroups() As String
Dim i As Integer
Dim cnt As Integer

    If Trim$(txtUserName.Text) = "" Then Exit Sub
    
    cnt = lstGroups.ListCount - 1
    
    If cnt > -1 Then
        ReDim aGroups(cnt)
        For i = 0 To cnt
            aGroups(i) = lstGroups.List(i)
        Next i
    End If
    
    frmLocalUserGroups.MY_LocalUserName = txtUserName.Text
    frmLocalUserGroups.MY_LocalUserGroups = aGroups
    frmLocalUserGroups.Show vbModal
    
    If frmLocalUserGroups.MY_IsCancel = False Then
        aGroups = frmLocalUserGroups.MY_LocalUserGroups
        lstGroups.Clear
        
        If Len(Join$(aGroups)) > 0 Then
            For i = 0 To UBound(aGroups)
                lstGroups.AddItem aGroups(i)
            Next i
        End If
        cmdSave.Enabled = True
    End If
End Sub

Private Sub imgEditMemberGroup_Click()
    lvGroups_DblClick
End Sub

Private Sub imgRefresh_Click()

    Dim tmpUserIndex As Integer

    If Not lvwUsers.SelectedItem Is Nothing Then _
        tmpUserIndex = lvwUsers.SelectedItem.Index
    
    SW_LOCAL_USERS = EnumAccounts
    FillUsersListView
    
    If tmpUserIndex > 0 And lvwUsers.ListItems.count >= tmpUserIndex Then
        With lvwUsers.ListItems(tmpUserIndex)
            .Selected = True
            .EnsureVisible
            lvwUsers_ItemClick lvwUsers.ListItems(tmpUserIndex)
        End With
    Else
        If lvwUsers.ListItems.count >= 1 Then
            With lvwUsers.ListItems(1)
                .Selected = True
                .EnsureVisible
                lvwUsers_ItemClick lvwUsers.ListItems(1)
            End With
        End If
    End If
End Sub

Private Sub lblHeaderSharedFolders_Click()
    frmSharedFolders.ZOrder 0
    frmSharedFolders.Show
End Sub

Private Sub lblHeaderStartUp_Click()
    frmStartUp.ZOrder 0
    frmStartUp.Show
End Sub

Private Sub lstGroups_GotFocus()
    lstGroups.BackColor = &H626262
End Sub

Private Sub lstGroups_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 45, 107 'Enter
            imgEditGroups_Click
        Case 46 ' DEL key presse
            lblRemoveGroup_Click
    End Select
End Sub

Private Sub lstGroups_LostFocus()
    lstGroups.BackColor = &H525252
End Sub

Private Sub lvGroups_DblClick()
    
    Dim tmpIndex        As Integer
        
    If lvGroups.SelectedItem Is Nothing Then Exit Sub
    
    tmpIndex = lvGroups.SelectedItem.Index
    
    frmGroupMembers.MY_LocalGroupName = lvGroups.SelectedItem
    frmGroupMembers.MY_NewMode = False
    frmGroupMembers.Show vbModal

    Call psubLocalGroupEnum
    
    imgRefresh_Click

    With lvGroups.ListItems(tmpIndex)
        .Selected = True
        .EnsureVisible
    End With
End Sub

Private Sub lvGroups_GotFocus()
    lvGroups.BackColor = &H626262
End Sub

Private Sub lvGroups_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            imgEditMemberGroup_Click
        Case 46 ' DEL key presse
            imgDeleteMembersGroup_Click
        Case 45, 107 ' INS key pressed
            imgAddMembersGroup_Click
    End Select
End Sub

Private Sub lvGroups_LostFocus()
    lvGroups.BackColor = &H525252
End Sub

Private Sub lvGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        If lvGroups.HitTest(x, y) Is Nothing Then
            mnuEdit.Visible = False
            mnuDelete.Visible = False
        Else
            mnuEdit.Visible = True
            mnuDelete.Visible = True
        End If
        PopupMenu mnuGroups
    End If
End Sub

Private Sub lvwUsers_GotFocus()
    lvwUsers.BackColor = &H626262
End Sub

Private Sub lvwUsers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim idx             As Long
    Dim i               As Long
    Dim sGroups()       As String
    Dim sDate, sTime    As Date
    
    If IsNew = True And Item.Tag <> "new" Then
        CancelNewUser Item.Index
        Exit Sub
    ElseIf IsNew = True And Item.Tag = "new" Then
        Exit Sub
    End If
    
    If cmdSave.Enabled = True Then
        If MsgBox("User " & txtUserName.Text & " is modified." & vbCrLf & _
            "Do you want to save the changes?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                cmdSave_Click
        End If
        
        cmdSave.Enabled = False
    End If
    
    idx = CLng(Item.Tag)
    
    bClicked = False
    
    With SW_LOCAL_USERS(idx)
        txtUserName.Text = .Name
        txtFullUserName.Text = .FullName
        txtDescription.Text = .Description
        txtPassword1.Text = "": txtPassword2.Text = ""
        txtPassword1.MinPassLen = .MaxPasswordLen
        txtPassword2.MinPassLen = .MaxPasswordLen
        lstGroups.Clear
        sGroups = Split(.Groups, ",")
        For i = 0 To UBound(sGroups)
            lstGroups.AddItem sGroups(i)
        Next
        bClicked = False
        chkMustChangePassword.Value = .PasswordExpired
        chkCanChangePassword.Value = .CannotChangePassword
        chkPasswordNeverExpire.Value = .PasswordNeverExpires
        chkAccountDisabled.Value = .AccountDisabled
        chkAccountDisabled.Enabled = .Name <> CurrentUser
        chkAccountLocked.Value = .AccountLocked
        chkAccountLocked.Enabled = .AccountLocked = 1
        bClicked = True

        If Trim$(str$(.LastLogin)) <> "00:00:00" Then
            lblLastLogon.Text = .LastLogin
        Else
            lblLastLogon.Text = "Never"
        End If
        
        imgDeleteUser.Visible = (Not LCase$(.Name) = "administrator") _
                                And (Not LCase$(.Name) = "guest") _
                                And (Not LCase$(.Name) = LCase$(CurrentUser))
    End With
    
    bClicked = True
End Sub

Private Sub cmdSave_Click()

Dim idx As Integer
Dim tmpKey As String
    
    tmpKey = lvwUsers.SelectedItem.Text
    
    txtUserName.Text = Trim$(txtUserName.Text)
    
    If CheckUserExists(txtUserName.Text) And IsNew = True Then
        MsgBox "Username is currently in use, " & vbCrLf & _
                "please select different username.", vbExclamation, "Conflict"
        txtUserName.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If SaveChanges = False Then Exit Sub

    cmdSave.Enabled = False
    
    IsNew = False

    imgAddUser.Visible = True
    txtUserName.Locked = True
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtPassword1.UnderlineColor = vbWhite
    txtPassword2.UnderlineColor = vbWhite
    
    SW_LOCAL_USERS = EnumAccounts
    FillUsersListView
    
    Screen.MousePointer = vbNormal
    
    lvwUsers.ListItems(tmpKey).Selected = True

End Sub

Private Function SaveChanges() As Boolean
Dim sUser   As SoftwareLocalUser
Dim i       As Integer
Dim sGroups As String
    If ComparePasswords = False Then Exit Function
      
    For i = 0 To lstGroups.ListCount - 1
        If sGroups <> "" Then sGroups = sGroups & ","
        sGroups = sGroups & lstGroups.List(i)
    Next
    
    With sUser
        .Name = txtUserName.Text
        .FullName = txtFullUserName.Text
        .Description = txtDescription.Text
        .Password = txtPassword1.Text
        .Groups = sGroups
        .PasswordExpired = chkMustChangePassword.Value
        .CannotChangePassword = chkCanChangePassword.Value
        .PasswordNeverExpires = chkPasswordNeverExpire.Value
        .AccountDisabled = chkAccountDisabled.Value
    End With
    
    SaveChanges = SaveAccount(sUser)
    
    If SaveChanges = False Then
        MsgBox "Changes WAS NOT saved.", vbCritical, Me.Caption
    End If
    
End Function

Private Function ComparePasswords() As Boolean
Dim Msg As String
    If (Len(Trim$(txtPassword1.Text)) = 0 And Len(Trim$(txtPassword2.Text)) = 0) _
        And IsNew Then
        ComparePasswords = False
        Msg = "Password for new users is mandatory!"
    ElseIf (Len(Trim$(txtPassword1.Text)) = 0 And Len(Trim$(txtPassword2.Text)) = 0) _
        And IsNew = False Then
        ComparePasswords = True
        Exit Function
    Else
        If txtPassword1.Text = txtPassword2.Text Then
            If txtPassword1.IsPassStrong = False Or txtPassword2.IsPassStrong = False Then
                Msg = "Password does not meet policy requirements." & vbCrLf & _
                "Your password must be at least x characters" & _
                "cannot repeat any of your previous x passwords; must contain" & _
                "capitals, numerals or punctuation; and cannot contain your" & _
                "account or full name. Please type a different password. Type" & _
                "a password which meets these requirements in both text boxes."
            Else
                ComparePasswords = True
                Exit Function
            End If
        Else
            Msg = "Passwords does not match." & vbCrLf & "Please enter passwords again."
        End If
    End If
    MsgBox Msg, vbExclamation, Me.Caption
    txtPassword1.Text = "": txtPassword2.Text = "": txtPassword1.SetFocus
    ComparePasswords = False
End Function

Private Sub lvwUsers_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46 ' DEL key pressed
            lblRemoveUser_Click
        Case 45, 107 ' INS key pressed
            lblAddUser_Click
    End Select
    lvwUsers.SetFocus
End Sub

Private Sub lvwUsers_LostFocus()
    lvwUsers.BackColor = &H525252
End Sub

Private Sub mnuAdd_Click()
    imgAddMembersGroup_Click
End Sub

Private Sub mnuDelete_Click()
    imgDeleteMembersGroup_Click
End Sub

Private Sub mnuEdit_Click()
    lvGroups_DblClick
End Sub

Private Sub mnuRefresh_Click()
    Call psubLocalGroupEnum
    
End Sub

Private Sub txtDescription_TextChange()
    If bClicked = True And NotEmpty(txtDescription.Text) Then _
        cmdSave.Enabled = True
End Sub

Private Sub txtFullUserName_TextChange()
    If bClicked = True And NotEmpty(txtFullUserName.Text) Then _
        cmdSave.Enabled = True
End Sub

Private Sub txtPassword1_TextChange()
    If bClicked = True And NotEmpty(txtPassword1.Text) Then _
        cmdSave.Enabled = True
    
    If Len(Trim$(txtPassword1.Text)) = 0 And IsNew Then
        txtPassword1.UnderlineColor = vbRed
    Else
        txtPassword1.UnderlineColor = vbWhite
    End If
End Sub

Private Sub txtPassword2_TextChange()
    If bClicked = True And NotEmpty(txtPassword2.Text) Then _
        cmdSave.Enabled = True
        
    If Len(Trim$(txtPassword2.Text)) = 0 And IsNew Then
        txtPassword2.UnderlineColor = vbRed
    Else
        txtPassword2.UnderlineColor = vbWhite
    End If
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 And IsNew = True Then
        lblRemoveUser_Click
        lvwUsers.SetFocus
    End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
Dim i               As Integer
Dim ForbbidenKey    As String
Dim arrKeys()       As String
    
    If Len(txtUserName.Text) > 19 Then KeyAscii = 0
    ForbbidenKey = "34,42,43,44,47,58,59,60,61,62,63,64,91,92,93,124"
    arrKeys = Split(ForbbidenKey, ",")
    For i = 0 To UBound(arrKeys)
        If KeyAscii = Int(arrKeys(i)) Then KeyAscii = 0
    Next i
End Sub

Private Sub txtUserName_LostFocus()

    txtUserName.Text = Trim$(txtUserName.Text)
    
    If IsNew = True Then
        If txtUserName.Text <> "" Then
            lvwUsers.SelectedItem.Text = txtUserName.Text
        Else
            txtUserName.SetFocus
        End If
    End If
End Sub

Private Function CancelNewUser(lItem As Integer) As Boolean
    
    If Trim$(txtUserName.Text) <> "" Then
         If MsgBox("Do you want to cancel new record?", vbExclamation + vbYesNo, Me.Caption) = vbNo Then
            lvwUsers.ListItems("NewUser").Selected = True
            lvwUsers.ListItems("NewUser").Text = txtUserName.Text
            CancelNewUser = False
            Exit Function
        End If
    End If
    
    IsNew = False
    cmdSave.Enabled = False
    imgAddUser.Visible = True
    txtUserName.Locked = True
    txtPassword1.UnderlineColor = vbWhite
    txtPassword2.UnderlineColor = vbWhite
    lvwUsers.ListItems.Remove "NewUser"
    If lItem > 0 And lItem <= lvwUsers.ListItems.count Then
        lvwUsers_ItemClick lvwUsers.ListItems(lItem)
        lvwUsers.ListItems(lItem).Selected = True
    End If
    CancelNewUser = True
    DoUnload = True
End Function

Private Sub txtUserName_TextChange()
    
    If bClicked = True And NotEmpty(txtUserName.Text) Then _
        cmdSave.Enabled = True
    
    If Len(Trim$(txtUserName.Text)) = 0 Then
        txtUserName.UnderlineColor = vbRed
        cmdSave.Enabled = Trim$(txtUserName.Text) <> ""
    Else
        txtUserName.UnderlineColor = vbWhite
    End If
End Sub

Private Function CheckUserExists(ByVal sName As String) As Boolean
Dim i As Integer

    CheckUserExists = False
    
    For i = 0 To UBound(SW_LOCAL_USERS)
        If LCase$(Trim$(SW_LOCAL_USERS(i).Name)) = LCase$(Trim$(sName)) Then _
                            CheckUserExists = True
    Next i
        
End Function
