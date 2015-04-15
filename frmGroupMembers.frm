VERSION 5.00
Begin VB.Form frmGroupMembers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit group members"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "&Remove"
      Height          =   340
      Left            =   5280
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add"
      Height          =   340
      Left            =   5280
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   340
      Left            =   5280
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   340
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtGroupComment 
      Height          =   720
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtGroupName 
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Group members"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   645
      Width           =   795
   End
   Begin VB.Label lblGroupName 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   435
      TabIndex        =   0
      Top             =   135
      Width           =   435
   End
End
Attribute VB_Name = "frmGroupMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private P_booNewMode As Boolean
Private P_strLocalGroupName As String

Public Property Let MY_NewMode(ByVal vNewValue As Boolean)
    P_booNewMode = vNewValue
End Property

Public Property Let MY_LocalGroupName(ByVal vNewValue As String)
    P_strLocalGroupName = vNewValue
End Property

Private Sub psubLocalGroupMemberEnum()
    Dim lngRet As Long
    Dim tLGMI2() As LOCALGROUP_MEMBERS_INFO_2_VB
    Dim lngCount As Long
    Dim i As Long
    List1.Clear
    lngRet = swLocalGroups.LocalGroupGetMembers2(P_strLocalGroupName, _
                                                  lngCount, _
                                                  tLGMI2())
    If lngRet = 0 Then
        For i = 0 To lngCount - 1
            List1.AddItem tLGMI2(i).vb_lgrmi2_domainandname
            List1.ItemData(List1.NewIndex) = tLGMI2(i).vb_lgrmi2_sid
        Next i
    Else
        MsgBox mErr.fncGetErrorString(lngRet), vbCritical
    End If
End Sub

Private Sub Command1_Click()
    Dim lngRet As Long
    If P_booNewMode Then
        lngRet = swLocalGroups.LocalGroupAdd1(txtGroupName.Text, txtGroupComment.Text)
    Else
        lngRet = swLocalGroups.LocalGroupSetInfo1002(P_strLocalGroupName, txtGroupComment.Text)
    End If
    If lngRet = 0 Then
        Unload Me
    Else
        MsgBox mErr.fncGetErrorString(lngRet)
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    
    Dim lngRet      As Long
    Dim newUser     As String
    
    newUser = InputBox("Please enter new group member username for group: " & _
                        vbCrLf & vbCrLf & Chr$(34) & P_strLocalGroupName & Chr$(34), _
                        "Add group members")
    
    If Not NotEmpty(newUser) Then Exit Sub
    
    lngRet = swLocalGroups.LocalGroupAddMembers3(P_strLocalGroupName, newUser)
    If lngRet <> 0 Then
        MsgBox mErr.fncGetErrorString(lngRet), vbCritical
    End If
    
    'Form3.Show vbModal
    Call psubLocalGroupMemberEnum
End Sub

Private Sub Command4_Click()
    Dim lngRet As Long
    If List1.ListIndex = -1 Then Exit Sub
    lngRet = swLocalGroups.LocalGroupDelMembers0(P_strLocalGroupName, _
                                                  List1.ItemData(List1.ListIndex))
    If lngRet <> 0 Then
        MsgBox mErr.fncGetErrorString(lngRet), vbCritical
    End If
    Call psubLocalGroupMemberEnum
End Sub

Private Sub Form_Load()
    Dim lngRet As Long
    Dim tLGI1 As LOCALGROUP_INFO_1_VB
    
    If P_booNewMode Then
        Me.Caption = "Create new local group"
        lblGroupName.Caption = "Група"
        txtGroupName.Visible = True
        Command3.Enabled = False
        Command4.Enabled = False
    Else
        Me.Caption = "Edit local group [" & P_strLocalGroupName & "]"
        txtGroupName.Visible = False
        Command3.Enabled = True
        Command4.Enabled = False
        lngRet = swLocalGroups.LocalGroupGetInfo1(P_strLocalGroupName, tLGI1)
        If lngRet = 0 Then
            lblGroupName.Caption = "Група     " & tLGI1.vb_lgrpi1_name
            txtGroupComment.Text = tLGI1.vb_lgrpi1_comment
            'Form3.MY_LocalGroupName = tLGI1.vb_lgrpi1_name
        Else
            MsgBox mErr.fncGetErrorString(lngRet), vbCritical
        End If
        Call psubLocalGroupMemberEnum
    End If
End Sub

Private Sub List1_Click()
    If List1.ListIndex <> -1 Then
        Command4.Enabled = True
    Else
        Command4.Enabled = False
    End If
End Sub
