VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNetwork 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Network adapter settings"
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12060
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
   ScaleHeight     =   9675
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.OptionButton opDNSServers 
      BackColor       =   &H00404040&
      Caption         =   "Obtain DNS servers autoatically"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   19
      Top             =   6360
      Width           =   5655
   End
   Begin VB.OptionButton opDNSServers 
      BackColor       =   &H00404040&
      Caption         =   "Use the following DNS server addresses"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   18
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Frame frameNetConfing 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Network settings"
      Enabled         =   0   'False
      ForeColor       =   &H000080FF&
      Height          =   6015
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   11775
      Begin VB.OptionButton optIPAddress 
         BackColor       =   &H00404040&
         Caption         =   "Use the following IP address"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   360
         Left            =   4560
         TabIndex        =   8
         Top             =   5280
         Width           =   990
      End
      Begin VB.Frame fraDNS 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   5415
         Begin VB.TextBox txtDNS2 
            Appearance      =   0  'Flat
            BackColor       =   &H00626262&
            Height          =   315
            Left            =   3480
            TabIndex        =   7
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtDNS1 
            Appearance      =   0  'Flat
            BackColor       =   &H00626262&
            Height          =   315
            Left            =   3480
            TabIndex        =   6
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label lblIPAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preffered DNS server"
            ForeColor       =   &H000080FF&
            Height          =   195
            Index           =   4
            Left            =   1650
            TabIndex        =   17
            Top             =   480
            Width           =   1545
         End
         Begin VB.Label lblIPAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alternate DNS server"
            ForeColor       =   &H000080FF&
            Height          =   195
            Index           =   3
            Left            =   1665
            TabIndex        =   16
            Top             =   1020
            Width           =   1530
         End
      End
      Begin VB.Frame fraIPAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   5415
         Begin VB.TextBox txtGateway 
            Appearance      =   0  'Flat
            BackColor       =   &H00626262&
            Height          =   315
            Left            =   3480
            TabIndex        =   5
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txtMask 
            Appearance      =   0  'Flat
            BackColor       =   &H00626262&
            Height          =   315
            Left            =   3480
            TabIndex        =   4
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox txtIPAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H00626262&
            Height          =   315
            Left            =   3480
            TabIndex        =   3
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblIPAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default gateway"
            ForeColor       =   &H000080FF&
            Height          =   195
            Index           =   2
            Left            =   1965
            TabIndex        =   14
            Top             =   1500
            Width           =   1200
         End
         Begin VB.Label lblIPAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Primary subnet mask"
            ForeColor       =   &H000080FF&
            Height          =   195
            Index           =   1
            Left            =   1695
            TabIndex        =   13
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label lblIPAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Primary IP address"
            ForeColor       =   &H000080FF&
            Height          =   195
            Index           =   0
            Left            =   1830
            TabIndex        =   12
            Top             =   420
            Width           =   1350
         End
      End
      Begin VB.OptionButton optIPAddress 
         BackColor       =   &H00404040&
         Caption         =   "Obtain an IP address automatically ( enable DHCP )"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
   End
   Begin MSComctlLib.ImageList ilNetwork 
      Left            =   6960
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetwork.frx":0000
            Key             =   "NETADAPTER"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvNetworkAdapters 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilNetwork"
      ForeColor       =   16777215
      BackColor       =   5395026
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Image imgReset 
      Height          =   225
      Left            =   120
      MouseIcon       =   "frmNetwork.frx":08DA
      MousePointer    =   99  'Custom
      Picture         =   "frmNetwork.frx":0A2C
      Stretch         =   -1  'True
      ToolTipText     =   "Опресни списъка (клавиш: F5)"
      Top             =   2760
      Width           =   225
   End
   Begin VB.Label lblNetworkAdapterList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Network adapters"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frmNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bRefresh As Boolean

Private Sub cmdSave_Click()
    
    Dim cnt As Integer
    
    cnt = lvNetworkAdapters.SelectedItem.Tag
    
    Screen.MousePointer = vbHourglass
    
    If optIPAddress(0).Value = True Then
        If Not SetNetConfig(HW_NETWORK_ADAPTERS(cnt).GUID, _
                            HW_NETWORK_ADAPTERS(cnt).PNPDeviceId) Then MsgBox "Error saving"
    Else
        ' Some kind of form validation
        If Not ValidateIPForm Then
            MsgBox "Моля въведете коректни параметри на мрежовата карта", vbExclamation, "Грешка"
            Exit Sub
        End If
        
        If Not IsGatewayValid Then
            If MsgBox("Основния шлюз не е в същата подмрежа," & vbCrLf & _
                       "дефинирана от IP адреса и подмрежовата маска." & vbCrLf & _
                       vbCrLf & "Желаете ли да продъжите?", _
                       vbExclamation + vbYesNo, "Конфликт") = vbNo Then Exit Sub
        End If
    
    
        ' Applying new IP addresses
        If SetNetConfig(HW_NETWORK_ADAPTERS(cnt).GUID, _
                        HW_NETWORK_ADAPTERS(cnt).PNPDeviceId, _
                        txtIPAddress, _
                        txtMask, _
                        txtGateway, _
                        txtDNS1, _
                        txtDNS2) _
                        Then
            MsgBox "Настройките бяха наложени успешно.", vbInformation
        Else
            MsgBox "Възникна проблем при прилагането на новите настройки.", _
                    vbExclamation, "Грешка"
        End If
    End If
    
    imgReset_Click
    
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_Load()
    lvNetworkAdapters.ColumnHeaders.Add , , "Име"
    lvNetworkAdapters.ColumnHeaders.Add , , "Име на връзката"
    lvNetworkAdapters.ColumnHeaders.Add , , "Свързаност"
    lvNetworkAdapters.ColumnHeaders.Add , , "IP адрес"
    lvNetworkAdapters.ColumnHeaders.Add , , "MAC адрес"
    lvNetworkAdapters.ColumnHeaders.Add , , "DHCP"
    
    FillNetworkAdapters
    
    If Not lvNetworkAdapters.SelectedItem Is Nothing Then
     lvNetworkAdapters.SelectedItem.Selected = False
    End If
End Sub

Private Sub FillNetworkAdapters(Optional bReload As Boolean = False)
On Error Resume Next

    Dim i       As Integer
    Dim cnt     As Integer
    Dim itmX    As ListItem
    
    If bReload Then HW_NETWORK_ADAPTERS = EnumNetworkAdapters
    
    cnt = UBound(HW_NETWORK_ADAPTERS)
    
    lvNetworkAdapters.ListItems.Clear
    
    For i = 0 To cnt
        With HW_NETWORK_ADAPTERS(i)
            If .Model <> vbNullString Then
                Set itmX = lvNetworkAdapters.ListItems.Add(, , .Model, , "NETADAPTER")
                itmX.Tag = i
                If (.Configuration.ConnectionStatus = 2) Then itmX.ForeColor = vbGreen
                
                itmX.SubItems(1) = .Configuration.NetConnectionID
                itmX.SubItems(2) = WMINetConnectorStatus(.Configuration.ConnectionStatus)
                itmX.SubItems(3) = .Configuration.IP(0)
                itmX.SubItems(4) = .MACAddress
                itmX.SubItems(5) = IIf(.Configuration.DHCPEnabled, "Enable", "Disable")
            End If
        End With
    Next i
    
    AutoSizeListViewColumns lvNetworkAdapters
End Sub

Private Sub Form_Resize()
On Error Resume Next

    With lvNetworkAdapters
        .Left = 120
        .Width = Me.ScaleWidth - 240
    End With
    
    With frameNetConfing
        .Width = Me.ScaleWidth - 240
        .Height = Me.ScaleHeight - .Top - 120
    
    End With
End Sub

Private Sub imgReset_Click()
    
    Dim itmX As Integer
    
    If Not lvNetworkAdapters.SelectedItem Is Nothing Then
        itmX = lvNetworkAdapters.SelectedItem.Index
    End If
    
    FillNetworkAdapters True
    
    bRefresh = True
    
    If itmX > 0 Then
        lvNetworkAdapters.ListItems(itmX).Selected = True
        lvNetworkAdapters.ListItems(itmX).EnsureVisible
        lvNetworkAdapters_ItemClick lvNetworkAdapters.ListItems(itmX)
    End If
    
    EnableIPField
    EnableDNSField
    
End Sub

Private Sub lvNetworkAdapters_GotFocus()
    lvNetworkAdapters.BackColor = &H626262
End Sub

Private Sub lvNetworkAdapters_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Static itmIndex As Integer
    
    Dim aDNS() As String
    Dim cnt As Integer
        
    If Item.Index = itmIndex And Not bRefresh Then Exit Sub
        
    bRefresh = False
    
    ResetIPFields
    ResetDNSFields
    
    frameNetConfing.Enabled = True
    
    With HW_NETWORK_ADAPTERS(Item.Tag).Configuration
        
        If .DHCPEnabled Or Not HasElements(.IP) Then
            optIPAddress(0).Value = True
            opDNSServers(0).Value = True
        Else
            optIPAddress(1).Value = True
            opDNSServers(1).Value = True
        End If
        
        If HasElements(.IP) Then txtIPAddress = .IP(0)
        If HasElements(.mask) Then txtMask = .mask(0)
        If HasElements(.GateWay) Then txtGateway = .GateWay(0)
        
        If Len(Join(.DNS)) > 0 Then
            cnt = UBound(.DNS)
            txtDNS1 = .DNS(0)
            If cnt > 0 Then txtDNS2 = .DNS(1)
        End If
       
    End With
    
    itmIndex = Item.Index
End Sub

Private Sub lvNetworkAdapters_LostFocus()
    lvNetworkAdapters.BackColor = &H525252
End Sub

Private Sub opDNSServers_Click(Index As Integer)
    Select Case Index
        Case 0
            fraDNS.Enabled = False
        Case 1
            fraDNS.Enabled = True
    End Select
    
    EnableDNSField
End Sub

Private Sub optIPAddress_Click(Index As Integer)
    Select Case Index
        Case 0
            fraIPAddress.Enabled = False
            opDNSServers(0).Enabled = True
            opDNSServers(0).Value = True
        Case 1
            fraIPAddress.Enabled = True
            opDNSServers(0).Enabled = False
            opDNSServers(1).Value = True
    End Select
    
    EnableIPField
End Sub

Private Sub ResetIPFields()
    txtIPAddress = vbNullString
    txtMask = vbNullString
    txtGateway = vbNullString
    
    txtIPAddress.BackColor = vbWhite
    txtMask.BackColor = vbWhite
    txtGateway.BackColor = vbWhite
End Sub

Private Sub ResetDNSFields()
    txtDNS1 = vbNullString
    txtDNS2 = vbNullString
    
    txtDNS1.BackColor = vbWhite
    txtDNS2.BackColor = vbWhite
End Sub

Private Sub EnableIPField()
    Dim b As Boolean
    
    b = optIPAddress(1).Value
    
    txtIPAddress.BackColor = IIf(b, vbWhite, &H626262)
    txtMask.BackColor = IIf(b, vbWhite, &H626262)
    txtGateway.BackColor = IIf(b, vbWhite, &H626262)
End Sub

Private Sub EnableDNSField()
    Dim b As Boolean
    
    b = opDNSServers(1).Value
    
    txtDNS1.BackColor = IIf(b, vbWhite, &H626262)
    txtDNS2.BackColor = IIf(b, vbWhite, &H626262)
End Sub

Private Sub txtDNS1_LostFocus()
    ValidateIP txtDNS1
End Sub

Private Sub txtDNS2_LostFocus()
    ValidateIP txtDNS2
End Sub

Private Sub txtGateway_LostFocus()
    ValidateIP txtGateway
End Sub

Private Sub txtIPAddress_LostFocus()
    ValidateIP txtIPAddress
End Sub

Private Function ValidateIP(ctrl As TextBox) As Boolean
    Dim sIP As String
    
    sIP = Trim$(ctrl.Text)
    
    If Len(sIP) = 0 Then
        ctrl.BackColor = vbWhite
        ValidateIP = True
        Exit Function
    End If
    
    ValidateIP = RegExIsMatch(sIP, REG_EX_IP_ADDRESS2, vbNullString)
    
    If Not ValidateIP Then
        ctrl.BackColor = &HC0C0FF
    Else
        ctrl.BackColor = vbWhite
    End If
End Function

Private Function IsGatewayValid() As Boolean
    
    Dim aIP() As String
    Dim aGW() As String
    Dim cnt   As Integer
    Dim bOK   As Boolean
    
    bOK = True
    
    aIP = Split(txtIPAddress.Text, ".")
    aGW = Split(txtGateway.Text, ".")
    
    If UBound(aIP) = 3 And UBound(aGW) = 3 Then
        For cnt = 0 To 2
            If aIP(cnt) <> aGW(cnt) Then bOK = False
        Next
    Else
        bOK = False
    End If
    
    IsGatewayValid = bOK
    
End Function

Private Sub txtMask_LostFocus()
    ValidateIP txtMask
End Sub

Private Function ValidateIPForm() As Boolean
    
    Dim bVld As Boolean
    
    ValidateIPForm = False
    
    If Not ValidateIP(txtIPAddress) Then Exit Function
    If Not ValidateIP(txtMask) Then Exit Function
    If Not ValidateIP(txtGateway) Then Exit Function
    
    If Not ValidateIP(txtDNS1) Then Exit Function
    If Not ValidateIP(txtDNS2) Then Exit Function
    
    ValidateIPForm = True
    
End Function
