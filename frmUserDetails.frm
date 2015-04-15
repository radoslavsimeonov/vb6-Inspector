VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserDetails 
   BackColor       =   &H00404040&
   Caption         =   "Workstation and users details"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
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
   ScaleHeight     =   9495
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   Tag             =   "USERDETAILS"
   WindowState     =   2  'Maximized
   Begin VB.Frame FraПотребителскиДанни 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Workstation users"
      ForeColor       =   &H0000FF00&
      Height          =   6255
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   11775
      Begin MSComctlLib.ImageList ilWorkstation 
         Left            =   6840
         Top             =   5400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserDetails.frx":0000
               Key             =   "OWNER"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserDetails.frx":13E2
               Key             =   "HARDDISK"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserDetails.frx":197C
               Key             =   "PRINTER"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserDetails.frx":2D5E
               Key             =   "USER"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserDetails.frx":4140
               Key             =   "STICKER"
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkOwner 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Owner"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Save"
         Height          =   360
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5640
         Width           =   1110
      End
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Clear"
         Height          =   360
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5640
         Width           =   1110
      End
      Begin VB.CheckBox chkNAVY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "NAVY"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   2040
         Width           =   735
      End
      Begin Inspector.TextBoxEx txtPhone 
         Height          =   510
         Left            =   10080
         TabIndex        =   19
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   900
         Caption         =   "Phone"
      End
      Begin Inspector.TextBoxEx txtBuilding 
         Height          =   510
         Left            =   7080
         TabIndex        =   17
         Top             =   2880
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   900
         Caption         =   "Building"
      End
      Begin Inspector.TextBoxEx txtFunction 
         Height          =   510
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   900
         Caption         =   "Function"
      End
      Begin Inspector.TextBoxEx txtFirstName 
         Height          =   510
         Left            =   2880
         TabIndex        =   11
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   900
         Caption         =   "First name"
      End
      Begin Inspector.TextBoxEx txtMiddleName 
         Height          =   510
         Left            =   5040
         TabIndex        =   12
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   900
         Caption         =   "Middle name"
      End
      Begin Inspector.TextBoxEx txtSurName 
         Height          =   510
         Left            =   7200
         TabIndex        =   13
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   900
         Caption         =   "Last name"
      End
      Begin Inspector.TextBoxEx txtRoom 
         Height          =   510
         Left            =   9120
         TabIndex        =   18
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   900
         Caption         =   "Room"
      End
      Begin Inspector.TextBoxEx cmbRanks 
         Height          =   510
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   900
         Caption         =   "Title"
         MaxRowsToDisplay=   21
         ControlType     =   1
         ComboStyle      =   1
      End
      Begin MSComctlLib.ListView lvUsers 
         Height          =   1575
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilWorkstation"
         ForeColor       =   16777215
         BackColor       =   5395026
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin Inspector.TextBoxEx txtUserName 
         Height          =   510
         Left            =   9360
         TabIndex        =   14
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         Caption         =   "Account name"
      End
      Begin MSComctlLib.ListView lvServices 
         Height          =   1335
         Left            =   240
         TabIndex        =   24
         Top             =   3960
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   5395026
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin Inspector.TextBoxEx txtDepartment 
         Height          =   510
         Left            =   4320
         TabIndex        =   16
         Top             =   2880
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   900
         Caption         =   "Department"
      End
      Begin VB.Image imgDeleteUser 
         Height          =   240
         Left            =   11400
         MouseIcon       =   "frmUserDetails.frx":5522
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":5674
         ToolTipText     =   "Delete (key: Del)"
         Top             =   1695
         Width           =   240
      End
      Begin VB.Image imgAddUser 
         Height          =   240
         Left            =   11400
         MouseIcon       =   "frmUserDetails.frx":6A46
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":6B98
         ToolTipText     =   "Add (key: +)"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgEditUser 
         Height          =   240
         Left            =   11400
         MouseIcon       =   "frmUserDetails.frx":7F6A
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":80BC
         ToolTipText     =   "Edit (key: Enter)"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image imgEditServices 
         Height          =   240
         Left            =   11400
         MouseIcon       =   "frmUserDetails.frx":948E
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":95E0
         ToolTipText     =   "Edit (key: Enter)"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label lblUserServices 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User services and accounts"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   285
         TabIndex        =   27
         Top             =   3720
         Width           =   1965
      End
   End
   Begin VB.Frame frameWorkstation 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Workstation"
      ForeColor       =   &H0000FF00&
      Height          =   2775
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   11775
      Begin Inspector.TextBoxEx txtSocketSKS 
         Height          =   510
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   900
         Caption         =   "Network socket label"
      End
      Begin Inspector.TextBoxEx cmbClassification 
         Height          =   510
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   900
         Caption         =   "Security clearance"
         Required        =   -1  'True
         MaxRowsToDisplay=   20
         ControlType     =   1
         ComboStyle      =   1
      End
      Begin Inspector.TextBoxEx txtBookNo 
         Height          =   510
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   900
         Caption         =   "Book number"
         Required        =   -1  'True
      End
      Begin Inspector.TextBoxEx txtBookDate 
         Height          =   510
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   900
         Caption         =   "Date"
         Required        =   -1  'True
      End
      Begin MSComctlLib.ListView lvHardDisks 
         Height          =   975
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilWorkstation"
         ForeColor       =   16777215
         BackColor       =   5395026
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvSecurityStickers 
         Height          =   855
         Left            =   3120
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilWorkstation"
         ForeColor       =   16777215
         BackColor       =   5395026
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvDevices 
         Height          =   855
         Left            =   6240
         TabIndex        =   6
         Top             =   1800
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilWorkstation"
         ForeColor       =   16777215
         BackColor       =   5395026
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory number of locally connected print devices"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   6240
         TabIndex        =   28
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Image imgEditDevice 
         Height          =   240
         Left            =   11355
         MouseIcon       =   "frmUserDetails.frx":A9B2
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":AB04
         ToolTipText     =   "Edit (key: Enter)"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgEditHardDisk 
         Height          =   240
         Left            =   11400
         MouseIcon       =   "frmUserDetails.frx":BED6
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":C028
         ToolTipText     =   "Edit  (key: Enter)"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgEditSticker 
         Height          =   240
         Left            =   5600
         MouseIcon       =   "frmUserDetails.frx":D3FA
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":D54C
         ToolTipText     =   "Edit (key: Enter)"
         Top             =   2080
         Width           =   240
      End
      Begin VB.Image imgAddSticker 
         Height          =   240
         Left            =   5600
         MouseIcon       =   "frmUserDetails.frx":E91E
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":EA70
         ToolTipText     =   "Add (key: +)"
         Top             =   1780
         Width           =   240
      End
      Begin VB.Image imgDeleteSticker 
         Height          =   240
         Left            =   5600
         MouseIcon       =   "frmUserDetails.frx":FE42
         MousePointer    =   99  'Custom
         Picture         =   "frmUserDetails.frx":FF94
         ToolTipText     =   "Delete (key: Del)"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label lblSecurityStickers 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Case security stickers"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   3120
         TabIndex        =   23
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label lblHardDisks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUserDetails.frx":11366
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmUserDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Const LV_GAP As String = "   "

Dim bEditUser As Boolean
Dim RankLabels() As Elements

Private Function ValidateWorkstationData() As Boolean
    
    ValidateWorkstationData = False
    
    If cmbClassification.SelectedItem.Index = 1 Then
        MsgBox "Please select security clearance", vbExclamation, "Error"
        cmbClassification.SetFocus
        Exit Function
    End If
    
    If ValidateTxt(txtBookNo) = False Then Exit Function
    If ValidateTxt(txtBookDate) = False Then Exit Function
    
    ValidateWorkstationData = True
End Function

Private Function ValidateUserData() As Boolean
      
    ValidateUserData = False
    
    If cmbRanks.SelectedItem.Index = 1 Then
        MsgBox "Please select user's tite", vbExclamation, "Error"
        cmbRanks.SetFocus
        Exit Function
    End If
    
    If ValidateTxt(txtFirstName) = False Then Exit Function
    If ValidateTxt(txtMiddleName) = False Then Exit Function
    If ValidateTxt(txtSurName) = False Then Exit Function
    If ValidateTxt(txtFunction) = False Then Exit Function
    
    ValidateUserData = True
End Function

Private Sub AddEditSecuritySticker(Optional newValue As String = "")
    
    Dim newSticker      As String
    Dim inputMessage    As String
    Dim itmX            As ListItem
    Dim itmX2           As ListItem
    Dim i               As Integer
    Dim cnt             As Integer
    
    inputMessage = "Please enter security sticker number" & vbCrLf & _
                   "sealed the computer case." & vbCrLf & _
                   vbCrLf & _
                   "Only numerical input allowed."
    
    newSticker = InputBox(inputMessage, _
                          IIf(newValue = "", "Add new security sticker", _
                            "Edit security sticker"), _
                          newValue)
    
    If IsNumeric(newSticker) Then
        If newValue <> "" Then
            Set itmX = lvSecurityStickers.FindItem(newSticker)
            Set itmX2 = lvSecurityStickers.FindItem(newValue)
            If Not itmX2 Is Nothing And itmX Is Nothing Then
                itmX2.Text = newSticker
                itmX2.Selected = True
                itmX2.EnsureVisible
            End If
        Else
            Set itmX = lvSecurityStickers.FindItem(newSticker)
            If itmX Is Nothing Then
                Set itmX = lvSecurityStickers.ListItems.Add(, , newSticker, , "STICKER")
                itmX.Selected = True
                itmX.EnsureVisible
            End If
        End If
        
        FillStickersArray
    End If
    lvSecurityStickers.SetFocus
    ValidateCaseStickers
End Sub

Private Sub ValidateCaseStickers()
    If (lvSecurityStickers.ListItems.count > 0 And cmbClassification.SelectedItem.Index > 2) _
        Or cmbClassification.SelectedItem.Index < 3 Then
        lblSecurityStickers.ForeColor = &H80FF&
    Else
        lblSecurityStickers.ForeColor = &H2141FF
    End If
End Sub

Private Sub AddEditDeviceInventaryNo(Optional newValue As String = "")
    
    Dim sDevInvNo As String
    Dim sMessage As String
    Dim sDevice As String
    Dim itmX As ListItem
    Dim idx As Integer
            
    If lvDevices.SelectedItem Is Nothing Then Exit Sub
        
    sDevice = lvDevices.SelectedItem.Text
    sMessage = "Please enter inventory number of " & vbCrLf & _
               sDevice
    
    sDevInvNo = InputBox(sMessage, "Inventory number", newValue)
    
    lvDevices.SelectedItem.SubItems(2) = sDevInvNo
    
    idx = lvDevices.SelectedItem.Tag
    
    HW_PRINTERS(idx).InventaryNo = sDevInvNo
    
    
End Sub

Private Sub FillHardDisks()

    Dim itmX As ListItem
    Dim i As Integer
    
    lvHardDisks.ListItems.Clear
       
    For i = 0 To UBound(HW_HARDDISKS)
        With HW_HARDDISKS(i)
            If .Removable = False Then
                If .Model <> vbNullString Then
                    Set itmX = lvHardDisks.ListItems.Add(, _
                        "HDD" & .Index, .Model & " (" & .Size & " GB)", , "HARDDISK")
                    
                    itmX.SubItems(1) = .SerialNumber
                    
                    If Len(Trim$(.Registry.InventaryNo)) > 0 Then _
                        itmX.SubItems(2) = IIf(.Registry.RegistryNo <> "", _
                                            .Registry.RegistryNo & "/", "") & _
                                            .Registry.InventaryNo & _
                                           IIf(.Registry.InventaryDate <> "", _
                                            "/" & .Registry.InventaryDate, "")
                End If
            End If
        End With
    Next i
    
    AutoSizeListViewColumns lvHardDisks
End Sub

Private Sub FillDevices()
    Dim itmX As ListItem
    Dim i As Integer
    
    lvDevices.ListItems.Clear
    
    For i = 0 To UBound(HW_PRINTERS)
        With HW_PRINTERS(i)
            If .Model <> vbNullString And .IsLocal Then
                
                Set itmX = lvDevices.ListItems.Add(, , .Model, , "PRINTER")
                
                itmX.SubItems(1) = "Printer"
                itmX.Tag = i
                
                If .InventaryNo <> vbNullString Then
                    itmX.SubItems(2) = .InventaryNo
                End If
            End If
        End With
    Next i
End Sub

Private Sub FillStickers()
    
    Dim i       As Integer
    
    lvSecurityStickers.ListItems.Clear
    
    If Len(Join(REG_WORKSTATION.CaseStickers)) > 0 Then
        For i = 0 To UBound(REG_WORKSTATION.CaseStickers)
            lvSecurityStickers.ListItems.Add , , REG_WORKSTATION.CaseStickers(i)
        Next
    End If
End Sub

Private Sub FillStickersArray()
    
    Dim itmX As Variant
    Dim cnt  As Integer
    Dim i    As Integer
    
    cnt = lvSecurityStickers.ListItems.count - 1
    
    If cnt > -1 Then
        ReDim REG_WORKSTATION.CaseStickers(cnt)
        
        For i = 0 To cnt
            REG_WORKSTATION.CaseStickers(i) = _
                lvSecurityStickers.ListItems(i + 1).Text
        Next
    Else
        Erase REG_WORKSTATION.CaseStickers
    End If
End Sub

Private Sub chkNAVY_Click()
    If chkNAVY.Value = 0 Then
        chkNAVY.ForeColor = &HC0C0C0
        SwitchForcesRanks 0, cmbRanks
    Else
        chkNAVY.ForeColor = &HFF00&
        SwitchForcesRanks 1, cmbRanks
    End If
End Sub

Private Sub chkOwner_Click()
    If chkOwner.Value = 0 Then
        chkOwner.ForeColor = &HC0C0C0
    Else
        chkOwner.ForeColor = &HFF00&
    End If
End Sub

Private Sub cmbClassification_ItemSelect(Item As MSComctlLib.ListItem)
    If Item.Index <> 1 Then
        REG_WORKSTATION.Classification = cmbClassification.SelectedItem.Key
    Else
        REG_WORKSTATION.Classification = vbNullString
    End If
    ValidateCaseStickers
End Sub

Private Sub cmbClassification_LostFocus()
    If cmbClassification.SelectedItem.Index = 1 Then _
        cmbClassification.Text = ""
    ValidateCaseStickers
End Sub

Private Sub cmdClear_Click()
    ResetUserDetails
    If lvUsers.ListItems.count > 0 Then _
        lvUsers.SelectedItem.Selected = False
End Sub

Private Sub ResetUserDetails()
    chkNAVY.Value = 0
    chkOwner.Value = 0
    
    cmbRanks.Text = vbNullString
    
    txtFirstName.Text = vbNullString
    txtMiddleName.Text = vbNullString
    txtSurName.Text = vbNullString
    txtUserName.Text = vbNullString
    txtFunction.Text = vbNullString
    txtDepartment.Text = vbNullString
    txtBuilding.Text = vbNullString
    txtRoom.Text = vbNullString
    txtPhone.Text = vbNullString
    
    lvServices.ListItems.Clear
    
    frmUserServices.MY_ClearServices
        
    bEditUser = False
    cmbRanks.SetFocus
End Sub

Private Sub cmdSave_Click()
    
    Dim cnt         As Integer
    
    If ValidateUserData = False Then Exit Sub
    
    cnt = GetUserIndex
    
    With REG_USERS(cnt)
        .Building = txtBuilding.Text
        .FirstName = txtFirstName.Text
        .Department = txtDepartment.Text
        .Function = txtFunction.Text
        .IsNavy = chkNAVY.Value = 1
        .IsOwner = ChooseOwner(cnt)
        .MiddleName = txtMiddleName.Text
        .Phone = txtPhone.Text
        .Rank = cmbRanks.Text
        .Room = txtRoom.Text
        .SurName = txtSurName.Text
        .username = txtUserName.Text
        .UserServices = SplitUserServices
    End With

    If Not bEditUser Then _
        ReDim Preserve REG_USERS(cnt + 1)

    bEditUser = False
    
    lvUsers.Visible = False
    
    FillUsersListView
    AutoSizeListViewColumns lvUsers
    
    lvUsers.ListItems(cnt + 1).Selected = True
    lvUsers.ListItems(cnt + 1).EnsureVisible
    
    lvUsers.Visible = True
    
    ResetUserDetails

End Sub

Private Function SplitUserServices() As RegisterUserService()
    
    Dim lvCnt       As Integer
    Dim cnt         As Integer
    Dim aService()  As RegisterUserService
    Dim itmX        As ListItem
    
    ReDim aService(0)
    
    lvCnt = lvServices.ListItems.count - 1
    
    If lvCnt < 0 Then
        SplitUserServices = aService
        Exit Function
    End If
    
    ReDim aService(lvCnt)
    
    For cnt = 0 To lvCnt
        
        Set itmX = lvServices.ListItems(cnt + 1)
        
        With itmX
            aService(cnt).ServiceName = .Text
            aService(cnt).ServiceAddress = .SubItems(1)
            aService(cnt).ServiceValue = .SubItems(2)
            aService(cnt).ServicePeriod = .SubItems(3)
        End With
    Next cnt
    
    SplitUserServices = aService
End Function

Private Function GetUserIndex() As Integer

    Dim cnt     As Integer
    
        If bEditUser Then
            cnt = lvUsers.SelectedItem.Index - 1
        Else
            cnt = UBound(REG_USERS)
            ReDim Preserve REG_USERS(cnt)
        End If
    
    GetUserIndex = cnt
End Function

Private Function ChooseOwner(ByVal cnt As Integer) As Boolean
    
    Dim bNewOwner   As Boolean
    Dim iOwner      As Integer
    Dim ownerMsg    As String
        
    bNewOwner = (chkOwner.Value = 1)
        
    If bNewOwner And modUserDetails.bWorkstationHasOwner Then
        iOwner = modUserDetails.WorkstationOwnerIndex
        
        If cnt <> iOwner Then
            With REG_USERS(iOwner)
                ownerMsg = "The current workstation owner is: " & vbCrLf & _
                            .Rank & " " & .FirstName & " " & .SurName & vbCrLf & _
                            "Do you want change owner with:" & vbCrLf & _
                            cmbRanks.Text & " " & txtFirstName.Text & _
                            " " & txtSurName.Text & " ?"
                bNewOwner = (MsgBox(ownerMsg, vbQuestion + vbYesNo, "Conflict") = vbYes)
                
                If bNewOwner Then modUserDetails.WorkstationOwnerIndex = cnt
                
                .IsOwner = Not bNewOwner
            End With
        End If
    ElseIf bNewOwner And Not modUserDetails.bWorkstationHasOwner Then
        modUserDetails.bWorkstationHasOwner = True
        modUserDetails.WorkstationOwnerIndex = cnt
    End If
    
    ChooseOwner = bNewOwner
End Function

Private Function AreUsersInitiliazed() As Boolean
    Dim bBound As Boolean
 
    On Error Resume Next
    bBound = IsNumeric(UBound(REG_USERS))
    On Error GoTo 0
    
    AreUsersInitiliazed = bBound
End Function

Private Sub FillUsersListView()
    
    Dim i           As Integer
    Dim itmX        As ListItem
    Dim strNames    As String
    
    
    lvUsers.ListItems.Clear
    
    If AreUsersInitiliazed = False Then Exit Sub
    
    lvUsers.Visible = False

    bWorkstationHasOwner = False
    
    For i = 0 To UBound(REG_USERS)
        With REG_USERS(i)
            If .FirstName <> vbNullString And .SurName <> vbNullString Then
                strNames = .Rank & " " & _
                           .FirstName & " " & _
                           .MiddleName & " " & _
                           .SurName
                
                Set itmX = lvUsers.ListItems.Add(, , strNames & LV_GAP, , "USER")
                
                If .IsOwner Then
                    itmX.SmallIcon = "OWNER"
                    bWorkstationHasOwner = True
                End If
                            
                itmX.SubItems(1) = .username & LV_GAP
                itmX.SubItems(2) = .Function & LV_GAP
                itmX.SubItems(3) = .Building & LV_GAP
                itmX.SubItems(4) = .Room & LV_GAP
                itmX.SubItems(5) = .Phone & LV_GAP
            End If
        End With
    Next
    
    lvUsers.Visible = True
    
End Sub

Private Sub Form_Load()
    
    bEditUser = False
    
    RankLabels = Fill_Rank(cmbRanks)
    cmbRanks.AutoResizeColumns
    
    With cmbClassification
        .Clear
        .ListItems.Add 1, , ""
        .ListItems.Add 2, "unclassified", "Unclassified"
        .ListItems.Add 3, "restricted", "Public Trust Position"
        .ListItems.Add 4, "confidential", "Confidential"
        .ListItems.Add 5, "secret", "Secret"
        .ListItems.Add 6, "topsecret", "Top Secret"
        .AutoResizeColumns
    End With
    
    With lvHardDisks
        .ColumnHeaders.Add , , "Hard Disk"
        .ColumnHeaders.Add , , "Serial number"
        .ColumnHeaders.Add , , "Inv No"
    End With
    
    With lvDevices
        .ColumnHeaders.Add , , "Device", .Width / 2
        .ColumnHeaders.Add , , "Type", .Width / 4
        .ColumnHeaders.Add , , "Inv No", .Width / 4
    End With
    
    With lvServices
        .ColumnHeaders.Add , , "Service", .Width / 4
        .ColumnHeaders.Add , , "Address", .Width / 4
        .ColumnHeaders.Add , , "Username", .Width / 4
        .ColumnHeaders.Add , , "Expire date", .Width / 4
        
        .View = lvwReport
        .HideColumnHeaders = False
    End With
    
    With lvUsers
        .ColumnHeaders.Add , , "User"
        .ColumnHeaders.Add , , "Account"
        .ColumnHeaders.Add , , "Function"
        .ColumnHeaders.Add , , "Building"
        .ColumnHeaders.Add , , "Room"
        .ColumnHeaders.Add , , "Phone"
    End With
    
    lvSecurityStickers.ColumnHeaders.Add , , _
        "Number", lvSecurityStickers.Width - 280
    
    FillHardDisks
    FillStickers
    FillWorkstationData
    FillUsersListView
    FillDevices

End Sub

Private Sub FillWorkstationData()
    With REG_WORKSTATION
        If .Classification <> "" Then
            cmbClassification.ListItems(.Classification).Selected = True
            cmbClassification.Text = cmbClassification.SelectedItem.Text
        End If
        txtBookNo.Text = .BookNo
        txtBookDate.Text = .BookDate
        txtSocketSKS.Text = .SocketSKS
    End With

End Sub

Private Sub FillUserServices(aSrv() As String)
       
    Dim i       As Integer
    Dim itmX    As ListItem
    Dim tSRV()  As String
    
    lvServices.ListItems.Clear
    
    If Len(Join$(aSrv)) = 0 Then Exit Sub
    
    For i = 0 To UBound(aSrv)
        tSRV = Split(aSrv(i), vbNullChar)
        
        Set itmX = lvServices.ListItems.Add(, , tSRV(0))
        itmX.SubItems(1) = tSRV(1)
        itmX.SubItems(2) = tSRV(2)
        itmX.SubItems(3) = tSRV(3)
    Next i
    
    'AutoSizeListViewColumns lvServices
End Sub

Private Sub imgAddService_Click()

End Sub

Private Sub imgAddSticker_Click()
    AddEditSecuritySticker
End Sub

Private Sub imgAddUser_Click()
    cmdClear_Click
End Sub

Private Sub imgDeleteSticker_Click()
    If Not lvSecurityStickers.SelectedItem Is Nothing Then
        lvSecurityStickers.ListItems.Remove _
            (lvSecurityStickers.SelectedItem.Index)
        lvSecurityStickers.SetFocus
        FillStickersArray
    End If
    ValidateCaseStickers
End Sub

Private Sub imgDeleteUser_Click()
    
    If lvUsers.SelectedItem Is Nothing Or lvUsers.ListItems.count = 0 Then _
        Exit Sub
    
    If MsgBox("Are you sure you want to delete selected user?", _
              vbYesNo + vbQuestion, _
              "Delete user") = vbNo Then Exit Sub
    
    ResetUserDetails
    RemoveUserFromArray lvUsers.SelectedItem.Index - 1
End Sub

Private Sub imgEditHardDisk_Click()
    If Not lvHardDisks.SelectedItem Is Nothing Then _
        lvHardDisks_DblClick
End Sub

Private Sub SentServicesForEdit()
    
    Dim lvCnt       As Integer
    Dim cnt         As Integer
    Dim aService()  As String
    Dim itmX        As ListItem
    
    lvCnt = lvServices.ListItems.count - 1
    
    If lvCnt < 0 Then
        Exit Sub
    End If
    
    ReDim aService(lvCnt)
    
    For cnt = 0 To lvCnt
        
        Set itmX = lvServices.ListItems(cnt + 1)
        
        With itmX
            aService(cnt) = .Text & vbNullChar & _
                            .SubItems(1) & vbNullChar & _
                            .SubItems(2) & vbNullChar & _
                            .SubItems(3)
        End With
    Next cnt
    
    frmUserServices.MY_UserServices = aService
End Sub

Private Sub imgEditServices_Click()
    
    SentServicesForEdit
   
    frmUserServices.Show vbModal
    
    FillUserServices frmUserServices.MY_UserServices
End Sub

Private Sub imgEditSticker_Click()
    If Not lvSecurityStickers.SelectedItem Is Nothing Then _
        AddEditSecuritySticker lvSecurityStickers.SelectedItem.Text
End Sub

Private Sub imgEditUser_Click()
    lvUsers_DblClick
End Sub

Private Sub lblEditHardDisk_Click()
    If Not lvHardDisks.SelectedItem Is Nothing Then _
        lvHardDisks_DblClick
End Sub

Private Sub lvDevices_DblClick()
    If Not lvDevices.SelectedItem Is Nothing Then _
        AddEditDeviceInventaryNo lvDevices.SelectedItem.SubItems(2)
End Sub

Private Sub lvDevices_GotFocus()
    lvDevices.BackColor = &H626262
End Sub

Private Sub lvDevices_LostFocus()
    lvDevices.BackColor = &H525252
End Sub

Public Sub lvHardDisks_DblClick()

    If lvHardDisks.SelectedItem Is Nothing Then Exit Sub
    If cmbClassification.SelectedItem.Index = 1 Then
        MsgBox "Please select workstation security clearane first.", vbExclamation
        cmbClassification.SetFocus
        Exit Sub
    End If

    If lvHardDisks.SelectedItem.Key <> "" Then
        RegisterHDDIndex = Int(Mid$(lvHardDisks.SelectedItem.Key, 4))
        Load frmRegister
        frmRegister.Show vbModal
    End If
    
    FillHardDisks
    lvHardDisks.ListItems("HDD" & RegisterHDDIndex).Selected = True
    lvHardDisks.ListItems("HDD" & RegisterHDDIndex).EnsureVisible
    
End Sub

Private Sub lvHardDisks_GotFocus()
    lvHardDisks.BackColor = &H626262
End Sub

Private Sub lvHardDisks_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then lvHardDisks_DblClick
End Sub

Private Sub lvHardDisks_LostFocus()
    lvHardDisks.BackColor = &H525252
End Sub

Private Sub lvSecurityStickers_DblClick()
    If Not lvSecurityStickers.SelectedItem Is Nothing Then _
        AddEditSecuritySticker lvSecurityStickers.SelectedItem.Text
End Sub

Private Sub lvSecurityStickers_GotFocus()
    lvSecurityStickers.BackColor = &H626262
End Sub

Private Sub lvSecurityStickers_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            imgEditSticker_Click
        Case vbKeyDelete
            imgDeleteSticker_Click
        Case 107, 187
            imgAddSticker_Click
    End Select
End Sub

Private Sub lvSecurityStickers_LostFocus()
    lvSecurityStickers.BackColor = &H525252
End Sub

Private Sub lvServices_DblClick()
    imgEditServices_Click
End Sub

Private Sub lvServices_GotFocus()
    lvServices.BackColor = &H626262
End Sub

Private Sub lvServices_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then imgEditServices_Click
End Sub

Private Sub lvServices_LostFocus()
    lvServices.BackColor = &H525252
End Sub

Private Sub lvUsers_DblClick()

    Dim cnt As Integer
    
   If Not lvUsers.SelectedItem Is Nothing Then
        
        cnt = lvUsers.SelectedItem.Index - 1
        
        If cnt = -1 Then Exit Sub
        
        With REG_USERS(cnt)
            txtBuilding.Text = .Building
            txtFirstName.Text = .FirstName
            txtFunction.Text = .Function
            txtDepartment.Text = .Department
            chkNAVY.Value = IIf(.IsNavy, 1, 0)
            chkOwner.Value = IIf(.IsOwner, 1, 0)
            txtMiddleName.Text = .MiddleName
            txtPhone.Text = .Phone
            cmbRanks.Text = .Rank
            txtRoom.Text = .Room
            txtSurName.Text = .SurName
            txtUserName.Text = .username
            ReFillUserServices .UserServices
        End With
        
        bEditUser = True
    End If
End Sub

Private Sub ReFillUserServices(arServices() As RegisterUserService)
On Error GoTo Err

    Dim i       As Integer
    Dim itmX    As ListItem
    Dim tSRV()  As String

    lvServices.ListItems.Clear

    For i = 0 To UBound(arServices)
        With arServices(i)
            Set itmX = lvServices.ListItems.Add(, , .ServiceName)
            itmX.SubItems(1) = .ServiceAddress
            itmX.SubItems(2) = .ServiceValue
            itmX.SubItems(3) = .ServicePeriod
        End With
    Next i

Err:

End Sub

Private Sub lvUsers_GotFocus()
    lvUsers.BackColor = &H626262
End Sub

Private Sub lvUsers_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            imgEditUser_Click
        Case vbKeyDelete
            imgDeleteUser_Click
        Case 107, 187
            imgAddUser_Click
    End Select
End Sub

Private Sub lvUsers_LostFocus()
    lvUsers.BackColor = &H525252
End Sub

Private Sub txtBookDate_TextChange()
    REG_WORKSTATION.BookDate = Trim$(txtBookDate.Text)
End Sub

Private Sub txtBookNo_TextChange()
    REG_WORKSTATION.BookNo = Trim$(txtBookNo.Text)
End Sub

Private Sub txtBuilding_KeyPress(KeyAscii As Integer)
    If Len(txtBuilding.Text) = 0 Or txtBuilding.SelLength = Len(txtBuilding.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtDepartment_KeyPress(KeyAscii As Integer)
    If Len(txtDepartment.Text) = 0 Or txtDepartment.SelLength = Len(txtDepartment.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    If Len(txtFirstName.Text) = 0 Or txtFirstName.SelLength = Len(txtFirstName.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtFunction_KeyPress(KeyAscii As Integer)
    If Len(txtFunction.Text) = 0 Or txtFunction.SelLength = Len(txtFunction.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
    If Len(txtMiddleName.Text) = 0 Or txtMiddleName.SelLength = Len(txtMiddleName.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtSocketSKS_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtSocketSKS_LostFocus()
    REG_WORKSTATION.SocketSKS = Trim$(txtSocketSKS.Text)
End Sub

Private Sub txtSurName_KeyPress(KeyAscii As Integer)
    If Len(txtSurName.Text) = 0 Or txtSurName.SelLength = Len(txtSurName.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub RemoveUserFromArray(ByVal idx As Integer)
    Dim i As Integer
    
    If idx < 0 Or idx > UBound(REG_USERS) Then Exit Sub
    
    If idx < UBound(REG_USERS) Then
        For i = idx To UBound(REG_USERS) - 1
            REG_USERS(i) = REG_USERS(i + 1)
        Next i
    End If
    
    If UBound(REG_USERS) = 0 Then
        Erase REG_USERS
    Else
        ReDim Preserve REG_USERS(UBound(REG_USERS) - 1)
    End If

    FillUsersListView
End Sub
