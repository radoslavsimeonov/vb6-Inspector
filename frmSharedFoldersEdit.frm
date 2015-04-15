VERSION 5.00
Begin VB.Form frmSharedFoldersEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   ". . ."
      Height          =   315
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtComment 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox txtShareName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   4575
   End
   Begin VB.Frame fraConnLimit 
      Caption         =   "User limit"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtMaxUses 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   630
         Width           =   735
      End
      Begin VB.OptionButton optMaxUses 
         Caption         =   "Alllowed users"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   690
         Width           =   1575
      End
      Begin VB.OptionButton optMaxUses 
         Caption         =   "Maximum allowed"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label lblSpecial 
      Caption         =   "This has been shared for administrative purposes. The share data cannot be set."
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3480
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblSharedFolder 
      Alignment       =   1  'Right Justify
      Caption         =   "Folder path"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   187
      Width           =   1515
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   1147
      Width           =   1515
   End
   Begin VB.Label lblSharedName 
      Alignment       =   1  'Right Justify
      Caption         =   "Share name"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   667
      Width           =   1515
   End
End
Attribute VB_Name = "frmSharedFoldersEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private P_booNewMode As Boolean
Private P_strShareName As String
Private P_tSI2 As SHARE_INFO_2_VB

Public Property Let MY_NewMode(ByVal vNewValue As Boolean)
    P_booNewMode = vNewValue
End Property

Public Property Let MY_ShareName(ByVal vNewValue As String)
    P_strShareName = vNewValue
End Property

Private Sub cmdSave_Click()
    Dim lngRet As Long
    
    If Len(txtShareName.Text) = 0 Then
        MsgBox "Please enter shared name.", vbCritical
        Exit Sub
    End If
    
    If P_booNewMode Then
        With P_tSI2
            .vb_shi2_netname = txtShareName.Text
            .vb_shi2_path = lblPath.Caption
            .vb_shi2_remark = txtComment.Text
            If optMaxUses(0).Value Then
                .vb_shi2_max_uses = -1
            Else
                .vb_shi2_max_uses = CLng(txtMaxUses.Text)
            End If
            .vb_shi2_type = STYPE_DISKTREE
            .vb_shi2_passwd = vbNullString
        End With
        lngRet = mNetShare.ShareAdd2(P_tSI2)
    Else
        With P_tSI2
            .vb_shi2_netname = txtShareName.Text
            .vb_shi2_remark = txtComment.Text
            If optMaxUses(0).Value Then
                .vb_shi2_max_uses = -1
            Else
                .vb_shi2_max_uses = CLng(txtMaxUses.Text)
            End If
        End With
        lngRet = mNetShare.ShareSetInfo2(P_strShareName, P_tSI2)
    End If
    If lngRet = 0 Then
        Unload Me
    Else
        MsgBox mErr.fncGetErrorString(lngRet), vbCritical
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim strPath As String
    mSHBF.hWnd = Me.hWnd
    mSHBF.Title = "Choose shared folder"
    If mNetShare.NetShareLocalCheck Then
        mSHBF.RootFolder = CSIDL_DRIVES
        mSHBF.Flags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN
        strPath = mSHBF.Show
    Else
        strPath = InputBox("Не може да споделите мрежови ресурси")
    End If
    If strPath = "" Then Exit Sub
    lblPath.Caption = strPath
End Sub

Private Sub Form_Load()
    Dim lngRet As Long
    
    If P_booNewMode Then
        Me.Caption = "New shared folder"
    Else
        Me.Caption = "Edit shared folder"
        Command3.Enabled = False
        lngRet = mNetShare.ShareGetInfo2(P_strShareName, P_tSI2)
        If lngRet = 0 Then
            With P_tSI2
                lblPath.Caption = .vb_shi2_path
                txtShareName.Locked = True
                txtShareName.Text = .vb_shi2_netname
                txtComment.Text = .vb_shi2_remark
                If .vb_shi2_max_uses = -1 Then
                    optMaxUses(0).Value = True
                Else
                    optMaxUses(1).Value = True
                    txtMaxUses.Text = .vb_shi2_max_uses
                End If
                If .vb_shi2_type And STYPE_SPECIAL Then
                    lblSpecial.Visible = True
                    txtShareName.Enabled = False
                    txtComment.Enabled = False
                    cmdSave.Enabled = False
                    optMaxUses(0).Enabled = False
                    optMaxUses(1).Enabled = False
                    txtMaxUses.Enabled = False
                End If
            End With
        Else
            MsgBox mErr.fncGetErrorString(lngRet), vbCritical
        End If
    End If
End Sub

Private Sub optMaxUses_Click(Index As Integer)
    With txtMaxUses
        If optMaxUses(0).Value Then
            .Enabled = False
            .BackColor = vbButtonFace
        Else
            .Enabled = True
            .BackColor = vbWindowBackground
        End If
    End With
End Sub

Private Sub txtMaxUses_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
