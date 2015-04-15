VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserServices 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User services"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12105
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
   ScaleHeight     =   3270
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbServiceName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox txtServiceAddress 
      Height          =   360
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdNew 
      Height          =   360
      Left            =   10680
      Picture         =   "frmUserServices.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "New"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      Height          =   360
      Left            =   11640
      Picture         =   "frmUserServices.frx":13D2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Clear"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   11160
      Picture         =   "frmUserServices.frx":27A4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   10800
      TabIndex        =   8
      Top             =   2760
      Width           =   1110
   End
   Begin VB.TextBox txtServicePeriod 
      Height          =   360
      Left            =   9120
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtServiceValue 
      Height          =   360
      Left            =   6960
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvServices 
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblServiceAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Website address  / Applications name"
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   0
      Width           =   2685
   End
   Begin VB.Label lblServicePeriod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expire date"
      Height          =   195
      Left            =   9120
      TabIndex        =   11
      Top             =   0
      Width           =   825
   End
   Begin VB.Label lblServiceValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User name"
      Height          =   195
      Left            =   6960
      TabIndex        =   10
      Top             =   0
      Width           =   765
   End
   Begin VB.Label lblServiceName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service name"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "frmUserServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ResidentServie    As String = "Permanent"

Private bEdit                   As Boolean

Private P_strUserServices()     As String
Private P_IsCancel              As Boolean

Public Property Get MY_UserServices() As String()
    MY_UserServices = P_strUserServices
End Property

Property Let MY_UserServices(sValue() As String)
    P_strUserServices = sValue
End Property

Public Sub MY_ClearServices()
    Erase P_strUserServices
End Sub

Private Sub Initialize()

    Dim itmX As ListItem
    Dim gap As Integer
    
    gap = 120
    
    With lvServices
        .ColumnHeaders.Add , , "Service name", cmbServiceName.Width
        .ColumnHeaders.Add , , "Address", txtServiceAddress.Width + 120
        .ColumnHeaders.Add , , "Value", txtServiceValue.Width + 120
        .ColumnHeaders.Add , , "Period", txtServicePeriod.Width
        
        .View = lvwReport
    End With
    
    With cmbServiceName
        .AddItem "Web access"
        .AddItem "E-mail account"
        .AddItem "MS SharePoint account"
        .AddItem "MS Lync account"
        .AddItem "Controlled-access website"
    
    End With
    
    txtServicePeriod = ResidentServie
    
    FillUserServices
    
End Sub

Private Sub FillUserServices()
       
    Dim i       As Integer
    Dim itmX    As ListItem
    Dim tSRV()  As String
    
    If Len(Join$(P_strUserServices)) = 0 Then Exit Sub
    
    lvServices.ListItems.Clear
    
    For i = 0 To UBound(P_strUserServices)
        tSRV = Split(P_strUserServices(i), vbNullChar)
        
        Set itmX = lvServices.ListItems.Add(, , tSRV(0))
        itmX.SubItems(1) = tSRV(1)
        itmX.SubItems(2) = tSRV(2)
        itmX.SubItems(3) = tSRV(3)
    Next i
    
    SelectLastListItem

End Sub

Private Sub SelectLastListItem()
    
    Dim cnt As Integer

    With lvServices
        If .ListItems.count > 0 Then
            cnt = .ListItems.count
            
            .ListItems(cnt).Selected = True
            .ListItems(cnt).EnsureVisible
            lvServices_ItemClick lvServices.ListItems(cnt)
        End If
    End With
End Sub

Private Sub SaveChanges()
    
    Dim lvCnt       As Integer
    Dim cnt         As Integer
    Dim aService()  As String
    Dim itmX        As ListItem
    Dim x           As Integer
    
    lvCnt = lvServices.ListItems.count - 1
    
    If lvCnt < 0 Then
        P_strUserServices = aService
        Exit Sub
    End If
    
    x = 0
    
    For cnt = 0 To lvCnt
        
        Set itmX = lvServices.ListItems(cnt + 1)
        
        With itmX
            If Trim$(.Text) <> vbNullString Then
                ReDim Preserve aService(x)
                
                aService(x) = .Text & vbNullChar & _
                                .SubItems(1) & vbNullChar & _
                                .SubItems(2) & vbNullChar & _
                                .SubItems(3)
                x = x + 1
            End If
        End With
    Next cnt
    
    P_strUserServices = aService
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()

    cmbServiceName.Text = ""
    txtServiceAddress = vbNullString
    txtServiceValue = vbNullString
    txtServicePeriod = ResidentServie
    
    bEdit = False
    
    cmdNew.Enabled = True
    
    With lvServices
        If Not .SelectedItem Is Nothing Then
            If .SelectedItem.Text = "" Then
                
                .ListItems.Remove (.SelectedItem.Index)
            
            Else
            
                If MsgBox("Are you sure you want to remove " & .SelectedItem.Text & " ?", _
                          vbQuestion + vbYesNo) = vbYes Then
                    .ListItems.Remove (.SelectedItem.Index)
                End If
            
            End If
        End If
    End With
      
    SelectLastListItem

End Sub

Private Sub cmdNew_Click()
    
    Dim itmX As ListItem
    
    EnableControls
    
    bEdit = False
    
    cmbServiceName.Text = vbNullString
    txtServiceAddress = vbNullString
    txtServiceValue = vbNullString
    txtServicePeriod = ResidentServie
        
    Set itmX = lvServices.ListItems.Add
    itmX.Selected = True
    itmX.EnsureVisible
    
    cmbServiceName.SetFocus
    
    cmdNew.Enabled = False
End Sub

Private Sub EnableControls()
    
    cmbServiceName.Enabled = True
    txtServiceValue.Enabled = True
    txtServicePeriod.Enabled = True
    txtServiceAddress.Enabled = True
    
    cmdClear.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()

    Dim itmX As ListItem
    
    If cmbServiceName.Text = vbNullString Then Exit Sub
    
    If lvServices.SelectedItem Is Nothing Then
       Set itmX = lvServices.ListItems.Add
    End If

    With lvServices
        .SelectedItem.Text = cmbServiceName.Text
        .SelectedItem.SubItems(1) = txtServiceAddress
        .SelectedItem.SubItems(2) = txtServiceValue
        .SelectedItem.SubItems(3) = txtServicePeriod
    End With

    
    cmdNew.Enabled = True
    
    If Not bEdit Then cmdNew_Click

End Sub

Private Sub Form_Load()
    Call Initialize
    
    bEdit = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveChanges
End Sub

Private Sub lvServices_DblClick()
    txtServiceAddress.SetFocus
End Sub

Private Sub lvServices_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim itmX As ListItem
    
    EnableControls
    
    bEdit = True
    
    With Item
        cmbServiceName.Text = Item.Text
        txtServiceAddress = Item.SubItems(1)
        txtServiceValue = Item.SubItems(2)
        txtServicePeriod = Item.SubItems(3)
    End With
    
End Sub

Private Sub txtServicePeriod_GotFocus()
    If txtServicePeriod = ResidentServie Then _
        txtServicePeriod = ""
End Sub

Private Sub txtServicePeriod_LostFocus()
    If txtServicePeriod = "" Then _
        txtServicePeriod = ResidentServie
End Sub
