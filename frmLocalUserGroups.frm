VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocalUserGroups 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit user groups"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8790
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvLocalGroups 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox txtSearchGroup 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Œ "
      Height          =   360
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1110
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<-- Remove"
      Height          =   360
      Left            =   3743
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add    -->"
      Height          =   360
      Left            =   3743
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvUserGroups 
      Height          =   2895
      Left            =   5280
      TabIndex        =   8
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblUserGroups 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User groups"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   870
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local groups"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmLocalUserGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private P_strLocalUserName As String
Private P_strLocalUserGroups() As String
Private P_IsCancel As Boolean

Public Property Let MY_LocalUserName(ByVal vNewValue As String)
    P_strLocalUserName = vNewValue
End Property

Public Property Get MY_LocalUserGroups() As String()
    MY_LocalUserGroups = P_strLocalUserGroups
End Property

Property Let MY_LocalUserGroups(sValue() As String)
    P_strLocalUserGroups = sValue
End Property

Property Get MY_IsCancel() As Boolean
    MY_IsCancel = P_IsCancel
End Property

Private Sub cmdAdd_Click()
    If lvLocalGroups.SelectedItem Is Nothing Then Exit Sub
        
    lvUserGroups.ListItems.Add , , lvLocalGroups.SelectedItem.Text
    lvLocalGroups.ListItems.Remove (lvLocalGroups.SelectedItem.Index)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub LoadLocalUsersGroups()

    Dim lngRet As Long
    Dim lngCount As Long
    Dim tLGUI() As String
    Dim i As Integer
    Dim itmX As ListItem
    Dim itmX1 As ListItem
    Dim cnt As Integer
    
'    lngRet = swLocalGroups.UserGetLocalGroups(P_strLocalUserName, 0, lngCount, tLGUI())
'    If lngRet = 0 Then
'        For i = 0 To lngCount - 1
'            lstUserGroups.AddItem tLGUI(i)
'        Next i
'    End If

    lvUserGroups.ListItems.Clear
    If Len(Join$(P_strLocalUserGroups)) = 0 Then Exit Sub
    
    cnt = UBound(P_strLocalUserGroups)

    For i = 0 To cnt
        Set itmX = lvUserGroups.ListItems.Add(, , P_strLocalUserGroups(i))
        Set itmX1 = lvLocalGroups.FindItem(P_strLocalUserGroups(i))
        If Not itmX1 Is Nothing Then lvLocalGroups.ListItems.Remove (itmX1.Index)
    Next i
End Sub

Private Sub LoadLocalGrouplist()
    Dim lngRet As Long
    Dim lngCount As Long
    Dim i As Long
    Dim itmX As ListItem
    Dim tLGI1VB() As LOCALGROUP_INFO_0_VB
    Dim lngType As Long
    
    lvLocalGroups.ListItems.Clear

    lngRet = swLocalGroups.LocalGroupEnum0(lngCount, tLGI1VB())
    If lngRet = 0 Then
        For i = 0 To lngCount - 1
            With tLGI1VB(i)
                Set itmX = lvLocalGroups.ListItems.Add(, , .vb_lgrpi0_name)
            End With
        Next i
    Else
        MsgBox mErr.fncGetErrorString(lngRet), vbCritical
    End If
End Sub

Private Sub cmdRemove_Click()
    If lvUserGroups.SelectedItem Is Nothing Then Exit Sub
    
    lvLocalGroups.ListItems.Add , , lvUserGroups.SelectedItem.Text
    lvUserGroups.ListItems.Remove (lvUserGroups.SelectedItem.Index)
End Sub

Private Sub cmdSave_Click()
    Dim lngRet As Long
    Dim tGroups() As String
    Dim i As Long
    Dim cnt As Integer
    
    cnt = lvUserGroups.ListItems.count - 1
    If cnt > -1 Then
        ReDim tGroups(cnt)
        For i = 0 To cnt
            tGroups(i) = lvUserGroups.ListItems(i + 1).Text
        Next i
    End If
    
'    lngRet = swLocalGroups.UserSetGroups0(P_strLocalUserName, tGroups())
'    If lngRet <> 0 Then
'        MsgBox mErr.fncGetErrorString(lngRet), vbCritical
'    End If

    P_strLocalUserGroups = tGroups
    P_IsCancel = False
    Unload Me
End Sub

Private Sub Form_Load()
    
    P_IsCancel = True
    
    lvLocalGroups.ColumnHeaders.Add , , "LocalGroups", lvLocalGroups.Width
    lvUserGroups.ColumnHeaders.Add , , "UserGroups", lvUserGroups.Width
    
    LoadLocalGrouplist
    LoadLocalUsersGroups
End Sub

Private Sub lvLocalGroups_DblClick()
    cmdAdd_Click
End Sub

Private Sub lvLocalGroups_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdAdd_Click
End Sub


Private Sub lvUserGroups_DblClick()
    cmdRemove_Click
End Sub

Private Sub lvUserGroups_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdRemove_Click
End Sub

Private Sub txtSearchGroup_Change()
    Dim itmX As ListItem
    
    Set itmX = lvLocalGroups.FindItem(txtSearchGroup.Text, , , lvwPartial)
    
    If Not itmX Is Nothing Then
        itmX.Selected = True
        itmX.EnsureVisible
    End If
End Sub

Private Sub txtSearchGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not lvLocalGroups.SelectedItem Is Nothing Then
            cmdAdd_Click
        End If
    End If
End Sub
