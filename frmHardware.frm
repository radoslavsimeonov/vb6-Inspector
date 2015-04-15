VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHardware 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Hardware summary"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   9855
   Tag             =   "HARDWARE"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraHardware 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   7335
      Begin MSComctlLib.ListView lvHardware 
         Height          =   2265
         Left            =   240
         TabIndex        =   6
         Top             =   360
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
      ScaleWidth      =   9855
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9855
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
            MouseIcon       =   "frmHardware.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   120
            Width           =   1530
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
            MouseIcon       =   "frmHardware.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   120
            Width           =   1500
            WordWrap        =   -1  'True
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
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   60
            MouseIcon       =   "frmHardware.frx":02A4
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   120
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   1800
            X2              =   1800
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   3840
            X2              =   3840
            Y1              =   240
            Y2              =   480
         End
      End
   End
   Begin MSComctlLib.ImageList ilHardware 
      Left            =   8280
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
            Picture         =   "frmHardware.frx":03F6
            Key             =   "SHARE"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHardware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblDevicesPrinters_Click()
    frmPeriphery.ZOrder 0
    frmPeriphery.Show
End Sub

Private Sub lblHeaderDeviceManager_Click()
    frmDeviceManager.ZOrder 0
    frmDeviceManager.Show
End Sub

Private Sub Form_Load()
    
    With lvHardware
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Path"
        
        .View = lvwReport
        .HideColumnHeaders = True
    End With
    
    
    FillHardwareList
End Sub

Private Sub Form_Resize()
On Error Resume Next

    ' Header menu
    picMenuHolder.Left = (Me.Width - picMenuHolder.Width) / 2

    fraHardware.Move 120, _
               700, _
               Me.ScaleWidth - 240, _
               Me.ScaleHeight - 700 - 120
    
    With lvHardware
        .Move 480, _
              480, _
              fraHardware.Width - 960, _
              fraHardware.Height - 960
        
        .ColumnHeaders(1).Width = .Width * 0.3
        .ColumnHeaders(2).Width = .Width * 0.7 - 300
    End With
End Sub

Public Sub FillHardwareList()
'On Error Resume Next
    
    
    Dim ii As Integer
    Dim i  As Integer

    modFillListView.MY_ListView = lvHardware
    lvHardware.ListItems.Clear
    
    With HW_MOTHERBOARD
        AddListItem .SystemModel, "System model"
        AddListItem .SystemMfg, "Manufacturer"
        
        If .SystemModel <> vbNullString Or .SystemMfg <> vbNullString Then _
            AddListItem "", ""
    
        AddListItem .Manufacturer & " " & .Model, "Motherboard"
        AddListItem .BIOS, "BIOS"
        AddListItem FormFactor(.ChassisType), "Chassis type"

        AddListItem "", "", True
        AddListItem HW_CPU.Model & " " & _
            IIf(HW_CPU.Architecture <> vbNullString, HW_CPU.Architecture & "bit", ""), "Processor"
                
        AddListItem "", "", True
        AddListItem FormatBytes(HW_RAM_MEMORY.TotalMemory, 0), "Installed memory"

        For ii = 0 To UBound(HW_RAM_MEMORY.Banks)
            AddListItem HW_RAM_MEMORY.Banks(ii).BankLabel & " " & _
                        FormatBytes(HW_RAM_MEMORY.Banks(ii).Capacity, 0) & "  " & _
                        RAMType(HW_RAM_MEMORY.Banks(ii).Type) & "-" & _
                        HW_RAM_MEMORY.Banks(ii).Speed & " " & _
                        RAMFormFactor(HW_RAM_MEMORY.Banks(ii).FormFactor), ""
        Next ii
                
        AddListItem "", ""
        
        If UBound(.Floppy) > -1 Then

            For ii = 0 To UBound(.Floppy)
                AddListItem .Floppy(ii), IIf(ii > 0, "", "Floppy drive") & " #" & i + 1
            Next ii

        End If
                
        AddListItem "", "", True

        If UBound(.Ports) > -1 Then

            For ii = 0 To UBound(.Ports)
                AddListItem .Ports(ii), IIf(ii > 0, "", "Ports")
            Next ii

        End If
        
        AddListItem "", "", True
        AddListItem HW_HID.Mouse, "Mouse"
        AddListItem HW_HID.Keyboard, "Keyboard"
                
    End With
            
    AddListItem "", ""

    For i = 0 To UBound(HW_SOUND_DEVICES)

        With HW_SOUND_DEVICES(i)
            AddListItem .Model, "Audio device #" & i + 1
        End With

    Next i
            
    AddListItem "", ""

    For i = 0 To UBound(HW_VIDEO_ADAPTERS)

        With HW_VIDEO_ADAPTERS(i)
            AddListItem .Model & _
                        IIf(.VideoRAM <> 0, " ( " & FormatBytes(.VideoRAM, 0) & " )", ""), _
                        "Video adapter"
        End With

    Next i

    AddListItem "", ""
    
    For i = 0 To UBound(HW_MONITORS)

        With HW_MONITORS(i)
            AddListItem .Manufacturer & " " & _
                        .Model & " " & _
                        .Size & IIf(.Size <> vbNullString, "''", "") & " " & _
                        .VideoInput, _
                        "Monitor #" & i + 1
        End With

    Next
            
    AddListItem "", ""
    For i = 0 To UBound(HW_HARDDISKS)

        With HW_HARDDISKS(i)
            If .Model <> vbNullString And .Family <> vbNullString Then
                If .Removable = False Then
                    Dim thdd As String
                    
                    If (Trim$(.Model) <> Trim$(.Family)) And .Family <> "" Then
                        thdd = .Family
                    Else: thdd = .Model: End If
        
                    AddListItem thdd & " ( " & _
                                Trim$(str$(.Size)) & " GB, " & _
                                .InterfaceType & ", s/n " & _
                                .SerialNumber & " )" _
                                , "Disk drive #" & i + 1
                End If
            End If
        End With

    Next i
            
    AddListItem "", ""

    For i = 0 To UBound(HW_CDROMS)

        With HW_CDROMS(i)
            If .Virtual = False Then
                AddListItem .Description & " " & _
                            .Manufacturer & " " & _
                            .Model, _
                            "Optical device #" & i + 1
            End If
        End With
    Next i
                        
    AddListItem "", ""

    For i = 0 To UBound(HW_NETWORK_ADAPTERS)

        With HW_NETWORK_ADAPTERS(i)
            Dim tNet As String
    
            If Len(Join$(.Configuration.IP)) > 0 Then _
                tNet = " (" & .Configuration.IP(0) & ")" Else: tNet = ""
            
            
            AddListItem .Model & tNet, "LAN card # " & i + 1
        End With
    Next
    
    AddListItem "", ""
    For i = 0 To UBound(HW_PRINTERS)
        With HW_PRINTERS(i)
            If .IsLocal = True Then
                AddListItem .Model, "Printer #" & i + 1
            End If
        End With
    Next i

    AutoSizeListViewColumns lvHardware
End Sub
