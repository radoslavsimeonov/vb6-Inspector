VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOS 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Software summary"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
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
   ScaleHeight     =   6060
   ScaleWidth      =   11310
   Tag             =   "SOFTWARE"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOS 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Width           =   6255
      Begin MSComctlLib.ListView lvOSDetails 
         Height          =   2265
         Left            =   360
         TabIndex        =   9
         Top             =   480
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
         SmallIcons      =   "ilOS"
         ForeColor       =   16777215
         BackColor       =   4210752
         Appearance      =   0
         NumItems        =   0
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
      ScaleWidth      =   11310
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11310
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
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   360
            Left            =   4320
            TabIndex        =   10
            Top             =   720
            Width           =   990
         End
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
            MouseIcon       =   "frmOS.frx":0000
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
            MouseIcon       =   "frmOS.frx":0152
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
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   5280
            MouseIcon       =   "frmOS.frx":02A4
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
            MouseIcon       =   "frmOS.frx":03F6
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
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   120
            MouseIcon       =   "frmOS.frx":0548
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
            MouseIcon       =   "frmOS.frx":069A
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
         End
      End
   End
   Begin MSComctlLib.ImageList ilOS 
      Left            =   8400
      Top             =   1200
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
            Picture         =   "frmOS.frx":07EC
            Key             =   "SHARE"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    With lvOSDetails
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Path"
        
        .View = lvwReport
        .HideColumnHeaders = True
    End With
    
    FillOSListView
    AutoSizeListViewColumns lvOSDetails
End Sub

Private Sub FillOSListView()

    Dim tKey    As String
    Dim cnt     As Integer
    
    modFillListView.MY_ListView = lvOSDetails
    
    With SW_OPERATING_SYSTEM
        AddListItem .CSName, "Computer name", , True
        AddListItem .Domain, "Domain/Workgroup", , True
        
        AddListItem "", ""
        AddListItem .Caption & " " & .Architecture, "Operating system", , True
        AddListItem .CSDVersion, "Service pack"
        AddListItem .Version, "Version"
        AddListItem .RegisteredUser, "Registered user"
        AddListItem .Organization, "Organization"
        If .ActivationStatus <> -1 Then _
            AddListItem GetWindowsActivationStatus(.ActivationStatus), "Windows activation", _
                        , , IIf(.ActivationStatus = 1, vbGreen, vbRed)
        AddListItem .InstallDate, "Install date"
        AddListItem OSLangById(.OSLanugage), "System language"
        AddListItem GetProductType(.ProductType), "Product type"
        
        AddListItem "", ""
        AddListItem GetCurrentCountry, "Country", , True
        AddListItem GetCurrentLanguage, "Language", , True
        AddListItem GetGeoFriendlyName, "Location", , True
        'AddListItem .CodeSet, "ANSI енкодинг", , True
        AddListItem GetCurrentTimeZone2, "Time zone", , True
                     
    End With
    
    On Local Error Resume Next
            
    AddListItem "", ""
        
    For cnt = 0 To UBound(SW_LICENSES)
        With SW_LICENSES(cnt)
            tKey = vbNullString
            If Len(.CDKey) > 5 Then _
                tKey = Mid$(.CDKey, 1, Len(.CDKey) - 5) & "*****"
            AddListItem tKey, .Product, , True
        End With
        Next
    
    AddListItem "", ""
    
    With SW_SECURITY_PRODUCTS
            
        For cnt = 0 To UBound(.AntiVirus)
            
            If .AntiVirus(cnt).ProductName <> vbNullString Then
                AddListItem .AntiVirus(cnt).ProductName, "Antivirus", , True
                AddListItem .AntiVirus(cnt).ProductVersion, "Version"
                If .AntiVirus(cnt).RTPStatus > -2 Then
                    AddListItem IIf((.AntiVirus(cnt).RTPStatus = -1), "Enabled", "Disabled"), _
                                "Real-time protection", , , _
                                IIf((.AntiVirus(cnt).RTPStatus = -1), vbGreen, vbRed)
                End If
                
                If .AntiVirus(cnt).UpToDate > -2 Then
                    AddListItem IIf((.AntiVirus(cnt).UpToDate = -1), "Yes", "No"), _
                                "Definition up-to-date", , , _
                                IIf((.AntiVirus(cnt).UpToDate = -1), vbGreen, vbRed)
                
                End If
            End If
        Next
                
        For cnt = 0 To UBound(SW_SECURITY_PRODUCTS.Firewall)
            If .Firewall(cnt).ProductName <> vbNullString Then
                AddListItem .Firewall(cnt).ProductName, "Firewall", , True
                AddListItem .Firewall(cnt).ProductVersion, "Verion"
                
                If .Firewall(cnt).Enabled > -2 Then _
                    AddListItem IIf((.Firewall(cnt).Enabled = -1), "Enabled", "Disabled"), _
                                "Real time protection", , , _
                                IIf((.Firewall(cnt).Enabled = -1), vbGreen, vbRed)
            End If
        Next
        
        For cnt = 0 To UBound(SW_SECURITY_PRODUCTS.Spyware)
            AddListItem .Spyware(cnt).ProductName, "Anti-Spyware", , True
        Next
    End With
End Sub

Private Sub Form_Resize()
On Error Resume Next

    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2

    fraOS.Move 120, _
               700, _
               Me.ScaleWidth - 240, _
               Me.ScaleHeight - 700 - 120
    
    With lvOSDetails
        .Move 480, _
              480, _
              fraOS.Width - 960, _
              fraOS.Height - 960
        
        .ColumnHeaders(1).Width = .Width * 0.4
        .ColumnHeaders(2).Width = .Width * 0.6 - 300
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

Private Sub lblHeaderSharedFolders_Click()
    frmSharedFolders.ZOrder 0
    frmSharedFolders.Show
End Sub

Private Sub lblHeaderStartUp_Click()
    frmStartUp.ZOrder 0
    frmStartUp.Show
End Sub

