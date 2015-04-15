VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServices 
   BackColor       =   &H00404040&
   Caption         =   "NT Services"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   18960
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
      ScaleWidth      =   18960
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   18960
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
            MouseIcon       =   "frmServices.frx":0000
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
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   7080
            MouseIcon       =   "frmServices.frx":0152
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
            MouseIcon       =   "frmServices.frx":02A4
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
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   4005
            MouseIcon       =   "frmServices.frx":03F6
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
            MouseIcon       =   "frmServices.frx":0548
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
            MouseIcon       =   "frmServices.frx":069A
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.PictureBox picContainers 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   0
      Left            =   360
      ScaleHeight     =   5415
      ScaleWidth      =   7455
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   7455
      Begin VB.CheckBox chkDisabled 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Disabled"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2820
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkManual 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Manual"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2820
         TabIndex        =   6
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame fraServicesList 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   6345
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Visible         =   0   'False
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   0
            Max             =   10000
            Scrolling       =   1
         End
         Begin VB.TextBox txtDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   2895
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   600
            Width           =   2415
         End
         Begin MSComctlLib.ListView lvwServices 
            Height          =   2265
            Left            =   2700
            TabIndex        =   12
            Top             =   120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   3995
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ilServices"
            ForeColor       =   16777215
            BackColor       =   5395026
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label lbStartStop 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   210
            Left            =   120
            MouseIcon       =   "frmServices.frx":07EC
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblRestart 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Restart"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   210
            Left            =   720
            MouseIcon       =   "frmServices.frx":093E
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   525
         End
      End
      Begin VB.CheckBox chkStopped 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Stopped"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4260
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Automatic"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2820
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkNameOnly 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Search only in service name"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   585
         Width           =   2535
      End
      Begin VB.CheckBox chkHideMS 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Hide by Microsoft"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5700
         MaskColor       =   &H00000080&
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox chkStarted 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Started"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4260
         TabIndex        =   2
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin Inspector.TextBoxEx txtSearch 
         Height          =   510
         Left            =   180
         TabIndex        =   1
         Top             =   0
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   900
         Caption         =   "Search service"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   5700
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Service status"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   4260
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Startup type"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   2820
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   0
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ilServices 
      Left            =   8640
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":0A90
            Key             =   "SERVICE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":102A
            Key             =   "STOPPED"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":1184
            Key             =   "DISABLED"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picWrapper 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8160
      ScaleHeight     =   825
      ScaleWidth      =   1425
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Menu mnuList 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuListStart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuListStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuListRestart 
         Caption         =   "Restart"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListStartUp 
         Caption         =   "Startup type"
         Begin VB.Menu mnuListStartAutomatic 
            Caption         =   "Automatic"
         End
         Begin VB.Menu mnuListStartManual 
            Caption         =   "Manual"
         End
         Begin VB.Menu mnuListStartDisabled 
            Caption         =   "Disabled"
         End
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
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAuto_Click()
    FillListViewValues
End Sub

Private Sub chkDisabled_Click()
    FillListViewValues
End Sub

Private Sub chkHideMS_Click()
    FillListViewValues
End Sub

Private Sub chkManual_Click()
    FillListViewValues
End Sub

Private Sub chkNameOnly_Click()
    FillListViewValues
End Sub

Private Sub chkStarted_Click()
    FillListViewValues
End Sub

Private Sub chkStopped_Click()
    FillListViewValues
End Sub

Private Sub Form_Load()
   
    lvwServices.ColumnHeaders.Add , , "Name", 3700
    lvwServices.ColumnHeaders.Add , , "Status", 900
    lvwServices.ColumnHeaders.Add , , "Startup", 900
    lvwServices.ColumnHeaders.Add , , "Publisher", 2000
    lvwServices.ColumnHeaders.Add , , "Description"
    lvwServices.View = lvwReport
    lvwServices.Sorted = True
    ProgressBar1.Top = lbStartStop.Top
    FillListViewValues
    
    picContainers(0).ZOrder 0
    
End Sub

Private Sub FillListViewValues(Optional bRefresh As Boolean = False)

    On Error GoTo Err

    Dim x    As Long
    Dim itmX As ListItem
    Dim tmp  As Integer

    tmp = lvwServices.SelectedItem.Index

    If bRefresh Then SW_SERVICES = EnumServices
    
    lvwServices.Visible = False
    lvwServices.ListItems.Clear

    For x = 0 To UBound(SW_SERVICES)
        With SW_SERVICES(x)
            If Not Trim$(.Caption) = "" Then
                If ElementSearch(txtSearch.Text, x) Then
                    Set itmX = lvwServices.ListItems.Add(, "K" & x, .Caption, , "SERVICE")
                    itmX.SubItems(1) = .state
    
                    If .state = "Stopped" Then
                        itmX.SmallIcon = "STOPPED"
                        itmX.ListSubItems(1).ForeColor = &H8080FF
                    End If
    
                    itmX.SubItems(2) = .StartMode
                    
                    If .StartMode = "Disabled" Then
                        itmX.SmallIcon = "DISABLED"
                        itmX.ListSubItems(2).ForeColor = &H8080FF
                    End If
                    
                    itmX.SubItems(3) = .Manufacturer
                    itmX.SubItems(4) = .Description
                End If
            End If
        End With
    Next x

    Set itmX = Nothing
    lvwServices.ListItems(tmp).Selected = True
    lvwServices_ItemClick lvwServices.ListItems(tmp)
    lvwServices.SelectedItem.EnsureVisible
    lvwServices.Visible = True
    Exit Sub
Err:
    tmp = 1

    Resume Next

End Sub

Private Function ElementSearch(ByVal strSearch As String, idx As Long) As Boolean

    Dim arrSearch() As String
    Dim i           As Integer

    With SW_SERVICES(idx)

        If InStr(LCase$(.Manufacturer), "microsoft") > 0 And chkHideMS.Value = 1 Then
            ElementSearch = False
            Exit Function
        End If

        If ((chkStarted.Value = 1 And .state = "Running") _
            Or (chkStopped.Value = 1 And .state = "Stopped")) _
            And ((chkAuto.Value = 1 And .StartMode = "Auto") _
            Or (chkManual.Value = 1 And .StartMode = "Manual") _
            Or (chkDisabled.Value = 1 And .StartMode = "Disabled")) Then
                ElementSearch = True
        End If

        If ElementSearch = False Then Exit Function
        
        arrSearch = Split(strSearch, " ")

        For i = 0 To UBound(arrSearch)
            ElementSearch = False
            strSearch = LCase$(arrSearch(i))

            If (InStr(LCase$(.Caption), strSearch) > 0 _
                Or (InStr(LCase$(.Description), strSearch) > 0 _
                And chkNameOnly.Value = 0)) Then
                    ElementSearch = True
            End If

            If ElementSearch = False Then Exit For
        Next i

    End With

End Function

Private Sub Form_Resize()
On Error Resume Next
    
    Dim i As Integer
    
    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2
    
    
    With picWrapper
        .Move 120, _
              820, _
              Me.ScaleWidth - 120, _
              Me.ScaleHeight - 700
    End With
    
    For i = 0 To picContainers.count - 1
        With picWrapper
            picContainers(i).Move .Left, _
                                   .Top, _
                                   .Width, _
                                   .Height

        End With
    Next i
    
    
    
    'SERVICES tab
    fraServicesList.Move 0, _
                         1020, _
                         picContainers(0).ScaleWidth, _
                         picContainers(0).ScaleHeight - 900
    
    lvwServices.Move 2600, _
                     60, _
                     fraServicesList.Width - 2600 - 120, _
                     fraServicesList.Height - 300 - 120
                     
    txtDescription.Height = fraServicesList.Height - 600
    
    If lvwServices.ColumnHeaders(5).Left < lvwServices.Width - 300 Then
        lvwServices.ColumnHeaders(5).Width = _
                                        fraServicesList.Width - _
                                        lvwServices.ColumnHeaders(5).Left - _
                                        lvwServices.Left - 480
    End If
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

Private Sub lblHeaderSharedFolders_Click()
    frmSharedFolders.ZOrder 0
    frmSharedFolders.Show
End Sub

Private Sub lblHeaderStartUp_Click()
    frmStartUp.ZOrder 0
    frmStartUp.Show
End Sub

Private Sub lblRestart_Click()
    ExecuteMethod "StopService"
    ExecuteMethod "StartService"
End Sub

Private Sub lvwServices_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim i As Long

    For i = 1 To lvwServices.ColumnHeaders.count

        If (i <> ColumnHeader.Index) Then
            lvwServices.ColumnHeaders(i).Tag = ""
        End If

    Next i

    If ColumnHeader.Tag = "ASC" Then
        ColumnHeader.Tag = "DESC"
        lvwServices.SortOrder = lvwDescending
    Else
        ColumnHeader.Tag = "ASC"
        lvwServices.SortOrder = lvwAscending
    End If

    lvwServices.SortKey = ColumnHeader.Index - 1
    lvwServices.Sorted = True
End Sub

Private Sub lvwServices_DblClick()
    
    If lvwServices.SelectedItem Is Nothing Then Exit Sub
    
    MsgBox SW_SERVICES(Replace(lvwServices.SelectedItem.Key, "K", "")).Description
End Sub

Private Sub lvwServices_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim idx As Long
    Dim sDesc As String
    
    If Item Is Nothing Then Exit Sub
        
    idx = Replace(Item.Key, "K", "")
        
    With SW_SERVICES(idx)
        
        sDesc = .Caption
        sDesc = sDesc & _
                IIf(InStr(LCase$(.Caption), "(" & LCase$(.Name) & ")") = 0, _
                " (" & .Name & ")", "")
        sDesc = sDesc & vbCrLf & vbCrLf & _
                "Description:" & vbCrLf & vbCrLf & _
                .Description & _
                vbCrLf & vbCrLf & _
                "Path to executable:" & _
                vbCrLf & vbCrLf & _
                .PathName

        txtDescription.Text = sDesc
                              
        mnuListStartDisabled.Enabled = Not .StartMode = "Disabled"
        mnuListStartAutomatic.Enabled = Not .StartMode = "Auto"
        mnuListStartManual.Enabled = Not .StartMode = "Manual"
        
        If .StartMode = "Disabled" _
                Or (.AcceptStop = False _
                And .state = "Running") Then
                
            lbStartStop.Visible = .StartMode = "Disabled" _
                                    And .state = "Running"
            lblRestart.Visible = False
            mnuListStart.Visible = False
            mnuListRestart.Visible = False
            mnuListStop.Visible = .StartMode = "Disabled" _
                                    And .state = "Running"
            mnuSeparator1.Visible = mnuListStop.Visible
        Else
            lbStartStop.Visible = True
            lblRestart.Visible = .state = "Running"
            mnuSeparator1.Caption = "-"
            mnuSeparator1.Visible = True
            mnuListRestart.Visible = .state = "Running"
            mnuListStop.Visible = .state = "Running"
            mnuListStart.Visible = Not .state = "Running"
            lbStartStop.Caption = _
                IIf(.state = "Running", "Stop", "Start")
            lbStartStop.Tag = _
                IIf(.state = "Running", "StopService", "StartService")
        End If
        
    End With

End Sub

Private Sub lvwServices_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        FillListViewValues True
    End If
End Sub

'Private Sub lvwServices_MouseDown(Button As Integer, _
'                                  Shift As Integer, _
'                                  x As Single, _
'                                  y As Single)
'
'    Dim itmX As ListItem
'    Dim idx  As Integer
'
'    Set itmX = lvwServices.HitTest(x, y)
'    lvwServices_ItemClick itmX
'
'    If Not itmX Is Nothing Then
'        Me.lvwServices.SelectedItem.Selected = False
'        Set lvwServices.SelectedItem = itmX
'        idx = Replace(itmX.Key, "K", "")
'
'        If Button = vbRightButton And lvwServices.ListItems.count > 0 Then
'            PopupMenu mnuList
'        End If
'    End If
'
'End Sub

Private Sub lvwServices_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        If lvwServices.HitTest(x, y) Is Nothing Then
            mnuListStart.Visible = False
            mnuListStop.Visible = False
            mnuListRestart.Visible = False
            mnuSeparator.Visible = False
            mnuSeparator1.Visible = False
            mnuListStartUp.Visible = False
        Else
            mnuSeparator.Visible = True
            'mnuSeparator1.Visible = True
            mnuListStartUp.Visible = True
        End If
        PopupMenu mnuList
    End If
    
End Sub

Private Sub mnuListRestart_Click()
    lblRestart_Click
End Sub

Private Sub mnuListStart_Click()
    lbStartStop_Click
End Sub

Private Sub mnuListStartAutomatic_Click()
    ExecuteMethod "Automatic"
End Sub

Private Sub mnuListStartDisabled_Click()
    ExecuteMethod "Disabled"
End Sub

Private Sub mnuListStartManual_Click()
    ExecuteMethod "Manual"
End Sub

Private Sub mnuListStop_Click()
    lbStartStop_Click
End Sub

Private Sub lbl_Click()
    ExecuteMethod "StopService"
    ExecuteMethod "StartService"
End Sub

Private Sub lbStartStop_Click()
    ExecuteMethod lbStartStop.Tag
End Sub

Private Sub ExecuteMethod(action As String)

    Dim tmp As String

    tmp = lvwServices.SelectedItem.Key
    ServiceMethods action, _
                   SW_SERVICES(Replace(lvwServices.SelectedItem.Key, "K", "")).Name
    FillListViewValues True
End Sub

Private Sub mnuRefresh_Click()
    FillListViewValues True
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    FillListViewValues
End Sub

Private Sub txtSearch_TextChange()
    If Trim$(txtSearch.Text) = "" Then FillListViewValues
End Sub

