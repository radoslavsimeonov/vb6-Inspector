VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSharedFolders 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Shared folders"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13395
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
   ScaleHeight     =   5745
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   Tag             =   "SOFTWARE"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ilShares 
      Left            =   6960
      Top             =   960
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
            Picture         =   "frmSharedFolders.frx":0000
            Key             =   "SHARE"
         EndProperty
      EndProperty
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
      ScaleWidth      =   13395
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13395
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
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   9000
            MouseIcon       =   "frmSharedFolders.frx":13E2
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
            MouseIcon       =   "frmSharedFolders.frx":1534
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
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   5280
            MouseIcon       =   "frmSharedFolders.frx":1686
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
            MouseIcon       =   "frmSharedFolders.frx":17D8
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
            MouseIcon       =   "frmSharedFolders.frx":192A
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
            MouseIcon       =   "frmSharedFolders.frx":1A7C
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
      End
   End
   Begin MSComctlLib.ListView lvShares 
      Height          =   2265
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilShares"
      ForeColor       =   16777215
      BackColor       =   5395026
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Menu mnuShares 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "New share"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit share"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete share"
         Shortcut        =   +{DEL}
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
Attribute VB_Name = "frmSharedFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    With lvShares
        .ColumnHeaders.Add , , "Name", .Width / 6
        .ColumnHeaders.Add , , "Folder path", .Width / 6
        .ColumnHeaders.Add , , "Type", .Width / 6
        .ColumnHeaders.Add , , "Description", .Width / 6
        .ColumnHeaders.Add , , "User limit", .Width / 6
        .ColumnHeaders.Add , , "Client connections", .Width / 6
        
        .View = lvwReport
        .HideColumnHeaders = False
    End With
    
    psubShareEnum
    'FillListViewShares
End Sub

Private Sub Form_Resize()
On Error Resume Next

    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2
    
    
    With lvShares
    
        lvShares.Move 120, _
                      700, _
                      Me.ScaleWidth - 2 * 120, _
                      Me.ScaleHeight - 700 - 120
        
        .ColumnHeaders(1).Width = .Width \ 6
        .ColumnHeaders(2).Width = .Width \ 6
        .ColumnHeaders(3).Width = .Width \ 6
        .ColumnHeaders(4).Width = .Width \ 6
        .ColumnHeaders(5).Width = .Width \ 6
        .ColumnHeaders(6).Width = .Width \ 6
    End With
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

Private Sub lblHeaderStartUp_Click()
    frmStartUp.ZOrder 0
    frmStartUp.Show
End Sub

Private Sub psubShareEnum()
    Dim lngRet As Long
    Dim lngCount As Long
    Dim i As Long
    Dim tSI502VB() As SHARE_INFO_502_VB
    Dim itmX As ListItem
    Dim strType As String
    
    lvShares.ListItems.Clear

    lngRet = mNetShare.ShareEnum502(lngCount, tSI502VB())
    
    If lngRet = 0 Then
        For i = 0 To lngCount - 1
            With tSI502VB(i)
                Set itmX = lvShares.ListItems.Add(, , .vb_shi502_netname, , "SHARE")
                itmX.SubItems(1) = .vb_shi502_path
                strType = ""
                If (.vb_shi502_type = STYPE_DISKTREE) Then
                    strType = "Disk Drive"
                Else
                    If (.vb_shi502_type And STYPE_SPECIAL) Then
                        strType = "Admin Share"
                        .vb_shi502_type = .vb_shi502_type Xor STYPE_SPECIAL
                    End If
                    If (.vb_shi502_type And STYPE_IPC) Then
                        strType = strType & " IPC"
                        .vb_shi502_type = .vb_shi502_type Xor STYPE_IPC
                    End If
                    If (.vb_shi502_type And STYPE_DEVICE) Then strType = strType & " Communication device"
                    If (.vb_shi502_type And STYPE_PRINTQ) Then strType = strType & " Print Queue"
                End If
                itmX.SubItems(2) = Trim$(strType)
                itmX.SubItems(3) = .vb_shi502_remark
                If .vb_shi502_max_uses = -1 Then
                    itmX.SubItems(4) = "Unlimited"
                Else
                    itmX.SubItems(4) = .vb_shi502_max_uses
                End If
                itmX.SubItems(5) = .vb_shi502_current_uses
            End With
        Next i
    Else
        MsgBox mErr.fncGetErrorString(lngRet), vbCritical
    End If
    
    lvShares.ListItems(lvShares.SelectedItem.Index).Selected = False
End Sub

Private Sub lvShares_DblClick()
    
    If lvShares.SelectedItem Is Nothing Then Exit Sub
    
    frmSharedFoldersEdit.MY_NewMode = False
    frmSharedFoldersEdit.MY_ShareName = lvShares.SelectedItem
    frmSharedFoldersEdit.Show vbModal

    Call psubShareEnum
    
End Sub

Private Sub lvShares_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Call psubShareEnum
    End If
End Sub

Private Sub lvShares_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim itmX As ListItem
    
    Set itmX = lvShares.HitTest(x, y)
    
    If Button = vbRightButton Then
                mnuEdit.Visible = False
        mnuDelete.Visible = False
        
        If Not itmX Is Nothing Then
            If InStr(itmX.Text, "$") = 0 And InStr(UCase$(itmX.Text), "IPC") = 0 Then
                mnuEdit.Visible = True
                mnuDelete.Visible = True
            End If
        End If
        PopupMenu mnuShares
    End If
    
End Sub

Private Sub mnuAdd_Click()
    frmSharedFoldersEdit.MY_NewMode = True
    frmSharedFoldersEdit.Show vbModal

    Call psubShareEnum
End Sub

Private Sub mnuAdd2_Click()
    frmSharedFoldersEdit.MY_NewMode = True
    frmSharedFoldersEdit.Show vbModal

    Call psubShareEnum
End Sub

Private Sub mnuDelete_Click()
    
    Dim lngRet As Long
    
    If lvShares.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("Delete shared folder " & lvShares.SelectedItem & "?", _
              vbExclamation + vbYesNo) = vbYes Then
        
        lngRet = mNetShare.ShareDel(lvShares.SelectedItem)
        
        If lngRet <> 0 Then
            MsgBox mErr.fncGetErrorString(lngRet), vbCritical
        End If
    End If

    Call psubShareEnum
End Sub

Private Sub mnuEdit_Click()
    lvShares_DblClick
End Sub

Private Sub mnuRefresh_Click()
    Call psubShareEnum
End Sub

Private Sub mnuRefreshList_Click()
    Call psubShareEnum
End Sub
