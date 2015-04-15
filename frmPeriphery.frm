VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPeriphery 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Printers"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13425
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
   ScaleHeight     =   6180
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPeriphery 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   6255
      Begin MSComctlLib.ListView lvDevices 
         Height          =   2265
         Left            =   360
         TabIndex        =   6
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
         SmallIcons      =   "ilDevices"
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
      ScaleWidth      =   13425
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13425
      Begin VB.PictureBox picMenuHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   700
         Left            =   0
         ScaleHeight     =   705
         ScaleWidth      =   5535
         TabIndex        =   1
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
            MouseIcon       =   "frmPeriphery.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   4
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
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   2160
            MouseIcon       =   "frmPeriphery.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   3
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
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   4065
            MouseIcon       =   "frmPeriphery.frx":02A4
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   120
            Width           =   1530
            WordWrap        =   -1  'True
         End
      End
   End
   Begin MSComctlLib.ImageList ilDevices 
      Left            =   7560
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":03F6
            Key             =   "PRINTER"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":17D8
            Key             =   "ON"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":2BBA
            Key             =   "OFF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":3F9C
            Key             =   "PRINTERD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":537E
            Key             =   "INACTIVE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":6760
            Key             =   "ACTIVE"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":7B42
            Key             =   "NETWORK"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":8F24
            Key             =   "SHARED"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriphery.frx":A306
            Key             =   "DEFAULT"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPeriphery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    With lvDevices
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Path"
        
        .View = lvwReport
        .HideColumnHeaders = True
    End With
    
    FillOSListView
    AutoSizeListViewColumns lvDevices
End Sub

Private Sub FillOSListView()

    Dim tKey    As String
    Dim i       As Integer
    
    modFillListView.MY_ListView = lvDevices
    
    For i = 0 To UBound(HW_PRINTERS)
        With HW_PRINTERS(i)
            If .Model <> vbNullString Then
                AddListItem .Model, "Model", , , &HFFFF00, _
                            IIf(.IsDefault, "PRINTERD", "PRINTER")
                
                AddListItem .Manufacturer, "Manufacturer"
                
                AddListItem IIf(.IsDefault, "Yes", ""), _
                            "Default", , , , _
                            IIf(.IsDefault, "DEFAULT", vbNullString)
                            
                AddListItem IIf(.IsLocal, "Local", "Network"), _
                            "Local" & " / " & "Network", , , , _
                            IIf(.IsLocal, "", "NETWORK")
                            
                AddListItem IIf(.IsOnline, "Yes", "No"), _
                            "Active", , , , _
                            IIf(.IsOnline, "ACTIVE", "INACTIVE")
                
                AddListItem .PortName, "Port name"
                
                If .IsShared Then
                    AddListItem .ShareName, "Shared name", , , , _
                        IIf(.IsShared, "SHARED", vbNullString)
                End If
                                           
                If .IsNetwork Then
                    AddListItem .hostname, "Host name"
                    AddListItem .IPAddress, "IP address"
                    AddListItem .ConnectionStat, "Host connection"
                End If
            End If
        End With
        
        If i < UBound(HW_PRINTERS) Then
            AddListItem "___________________", ""
            AddListItem "", ""
        End If
    Next i
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2

    fraPeriphery.Move 120, _
               700, _
               Me.ScaleWidth - 240, _
               Me.ScaleHeight - 700 - 120
    
    With lvDevices
        .Move 480, _
              480, _
              fraPeriphery.Width - 960, _
              fraPeriphery.Height - 960
        
        .ColumnHeaders(1).Width = .Width * 0.4
        .ColumnHeaders(2).Width = .Width * 0.6
    End With

End Sub

Private Sub lblHeaderDeviceManager_Click()
    frmDeviceManager.ZOrder 0
    frmDeviceManager.Show
End Sub

Private Sub lblHeaderSummary_Click()
    frmHardware.ZOrder 0
    frmHardware.Show
End Sub

Private Sub lblLabel1_Click()

End Sub
