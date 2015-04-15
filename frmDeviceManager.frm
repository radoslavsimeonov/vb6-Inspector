VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeviceManager 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Device manager"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   9825
   Tag             =   "HARDWARE"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMethods 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   255
      ScaleWidth      =   5895
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   5895
      Begin VB.OptionButton optDisabled 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Изключено"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton optEnabled 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Включено"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Промени"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   4920
         MouseIcon       =   "frmDeviceManager.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   0
         Width           =   690
      End
      Begin VB.Label lblMethonds 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Състояние"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   825
      End
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
      ScaleWidth      =   9825
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   9825
      Begin VB.PictureBox picMenuHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   700
         Left            =   0
         ScaleHeight     =   705
         ScaleWidth      =   5535
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   5535
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   3840
            X2              =   3840
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
         Begin VB.Label lblHeaderSummary 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HARDWARE SUMMARY"
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
            Left            =   60
            MouseIcon       =   "frmDeviceManager.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   120
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblHeaderDeviceManager 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DEVICE MANAGER"
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
            Left            =   2160
            MouseIcon       =   "frmDeviceManager.frx":02A4
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDevicesPrinters 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRINT   DEVICES"
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
            Left            =   4065
            MouseIcon       =   "frmDeviceManager.frx":03F6
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   120
            Width           =   1530
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Frame frameOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   5895
      Begin VB.Frame frameDevice 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   5895
         Begin MSComctlLib.ListView lvProperties 
            Height          =   2175
            Left            =   960
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   360
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   3836
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   33023
            BackColor       =   12632256
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial CYR"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Field"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   0
            X2              =   5880
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.Label lblInstructions 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For device details (manufacturer, model, serial number, setting and etc.) select device from the left-hand list."
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   5610
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDeviceName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   120
         Width           =   5040
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgDeviceIcon 
         Height          =   375
         Left            =   120
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CheckBox chkHiddenDevices 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Show &hidden devices"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin MSComctlLib.ImageList ilDevices 
      Left            =   360
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   4210752
      MaskColor       =   4210752
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   6588
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ilDevices"
      Appearance      =   0
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   2640
      MouseIcon       =   "frmDeviceManager.frx":0548
      MousePointer    =   99  'Custom
      Picture         =   "frmDeviceManager.frx":069A
      Stretch         =   -1  'True
      ToolTipText     =   "Опресни списъка"
      Top             =   2040
      Width           =   225
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   45
   End
End
Attribute VB_Name = "frmDeviceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub LoadDevices(Optional bRefresh As Boolean = False)

    Dim selIndex As String

    ilDevices.ListImages.Clear

    If Not TV.SelectedItem Is Nothing Then selIndex = TV.SelectedItem.Key
    If bRefresh Then
        Erase HW_DEVICES
        GetDevProp "", HW_DEVICES
    End If

    ClassImageList = GetClassImageList
    FillImageListWithClassImageList ClassImageList, ilDevices
    FillTV selIndex
End Sub

Private Sub chkHiddenDevices_Click()
    LoadDevices
End Sub

Private Sub Form_Activate()
    chkHiddenDevices.Caption = "Show hidden devices"
    optEnabled.Caption = "Enable"
    optDisabled.Caption = "Disable"
    lblMethonds.Caption = "Status"
    
    frameOptions.BackColor = &H404040
    lvProperties.BackColor = &H404040
    frameDevice.BackColor = &H404040
    
    SetTVBackColor TV, &H404040
    LoadDevices
End Sub

Private Sub FillTV(Optional ByVal selIndex As String)

    Dim nodX As Node
    Dim nodR As String
    Dim x    As Long
    Dim i    As Integer

    TV.Visible = False
    TV.Nodes.Clear

    For x = 0 To UBound(HW_DEVICES)

        With HW_DEVICES(x)
            If .ClassDesc <> vbNullString Then
                nodR = .ClassDesc
    
                If Not NodeExists(TV, nodR) Then
                    Set nodX = TV.Nodes.Add(, , nodR, .ClassDesc, GetClassImageListIndex(ClassImageList, .ClassGuid))
                    nodX.Sorted = True
                    nodX.BackColor = &H404040
                    nodX.ForeColor = vbWhite
                    Set nodX = Nothing
                End If
    
                If ((chkHiddenDevices.Value = 0) And LCase$(.ClassDesc) <> "non-plug and play drivers" And .Hidden = False) Or ((chkHiddenDevices.Value = 1)) Then
                    Set nodX = TV.Nodes.Add(nodR, tvwChild, "H" & .Index, IIf(.FriendlyName <> "", .FriendlyName, .DeviceDesc), GetClassImageListIndex(ClassImageList, .ClassGuid))
                    nodX.Sorted = True
                    nodX.BackColor = &H404040
                    nodX.ForeColor = vbWhite
    
                    If .Enabled = False Then
                        nodX.ForeColor = &HC0C0FF
                        nodX.Parent.ForeColor = &HC0C0FF
                    End If
    
                    If nodR = "Other devices" Then
                        nodX.ForeColor = &HC0FFFF
                        nodX.Parent.ForeColor = &HC0FFFF
                    End If
                End If
    
DoNext:
                Set nodX = Nothing
            End If
        End With

    Next x

    For i = TV.Nodes.count To 1 Step -1

        If TV.Nodes(i).Children = 0 And TV.Nodes(i).Parent Is Nothing Then
            TV.Nodes.Remove (TV.Nodes(i).Index)
        End If

    Next i

    If (selIndex = "" And TV.Nodes.count > 0) Or (selIndex <> "" And Not NodeExists(TV, selIndex)) Then selIndex = TV.Nodes(1).FirstSibling.Key
    
    TV.Visible = True
    
    If selIndex = "" Then Exit Sub
    
    TV.Nodes(selIndex).Selected = True
    TV.Nodes(selIndex).EnsureVisible
    TV_NodeClick TV.Nodes(selIndex)
    TV.Nodes(selIndex).Expanded = False
End Sub

Private Function NodeExists(TV As TreeView, ByVal sKey As String) As Boolean

    Dim nd As Node

    On Error Resume Next

    Set nd = TV.Nodes(sKey)
    NodeExists = (Err = 0)
    Set nd = Nothing
End Function

Private Sub Form_Resize()
On Error Resume Next

    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2
        
    TV.Move L_TABLE_LEFT, _
            1200, _
            Me.ScaleWidth - L_TABLE_LEFT, _
            Me.ScaleHeight - TV.Top - 120
    
    chkHiddenDevices.Move 120, 750
    
    picMethods.Move Me.ScaleWidth - picMethods.Width - 220, 750
    
    frameOptions.Move Me.ScaleWidth - frameOptions.Width - L_TABLE_LEFT * 2, _
                      TV.Top, _
                      5895, _
                      Me.ScaleHeight - 800 - 700
                      
    lblInstructions.Move 360, _
                         lblDeviceName.Top + lblDeviceName.Height + 460, _
                         frameOptions.Width - 720, _
                         lblInstructions.Height
    
    frameDevice.Move 0, _
                     lblDeviceName.Top + lblDeviceName.Height + 220, _
                     frameOptions.Width, _
                     frameOptions.Height - frameDevice.Top
                     
    lvProperties.Move 0, 220, frameDevice.Width, frameDevice.Height - 220
End Sub

Private Sub Image2_Click()

End Sub

Private Sub lblApply_Click()

    Dim K As Integer

    If TV.SelectedItem Is Nothing Then Exit Sub

    lblApply.Enabled = False
    K = Mid$(TV.SelectedItem.Key, 2)

    Screen.MousePointer = vbHourglass
    
    If EnableDevice(HW_DEVICES(K).Index, optEnabled.Value) = True Then
        LoadDevices True
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub lblDeviceName_Change()
    lblDeviceName.Top = 180 - lblDeviceName.Height \ 2
    frameDevice.Top = lblDeviceName.Top + lblDeviceName.Height + 220
    Form_Resize
End Sub

Private Sub lblDevicesPrinters_Click()
    frmPeriphery.ZOrder 0
    frmPeriphery.Show
End Sub

Private Sub lblHeaderSummary_Click()
    frmHardware.ZOrder 0
    frmHardware.Show
End Sub

Private Sub optDisabled_Click()
    lblApply.Enabled = (optDisabled.Value = (optEnabled.Tag <> ""))
End Sub

Private Sub optEnabled_Click()
    lblApply.Enabled = (optEnabled.Value = (optEnabled.Tag = ""))
End Sub

Private Sub TV_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 Then
        LoadDevices True
    End If

End Sub

Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim K  As Integer
    Dim td As HardwareDevice

    If Node.Parent Is Nothing Then
        lblDeviceName.Caption = Node.Text
        imgDeviceIcon.Picture = ilDevices.ListImages(Node.Image).Picture
        If Node.Key <> "Computer" Then
            picMethods.Visible = False
            frameDevice.Visible = False
        Else
            picMethods.Visible = False
            frameDevice.Visible = True
            modFillListView.MY_ListView = lvProperties
            FillComputerDetails lvProperties
        End If
        
        Exit Sub
    End If

    K = Mid$(Node.Key, 2)

    With HW_DEVICES(K)
        lblDeviceName.Caption = IIf(.FriendlyName <> "", .FriendlyName, .DeviceDesc)
        imgDeviceIcon.Picture = ilDevices.ListImages(GetClassImageListIndex(ClassImageList, .ClassGuid)).Picture

        If .CanDisable Then
            optDisabled.Enabled = True
        Else
            optDisabled.Enabled = False
        End If

        If .Enabled Then
            optEnabled.Tag = "1"
            optEnabled.Value = 1
        Else
            optEnabled.Tag = ""
            optDisabled.Value = 1
        End If
                                
        lblApply.Enabled = False
        picMethods.Visible = True
        frameDevice.Visible = True
    End With
    
    FillDeviceDetails HW_DEVICES(K), lvProperties
End Sub

