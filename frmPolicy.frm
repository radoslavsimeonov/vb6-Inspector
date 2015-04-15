VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPolicy 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Security policy"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   ControlBox      =   0   'False
   Icon            =   "frmPolicy.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   12090
   Tag             =   "SECPOL"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Local security policies"
      ForeColor       =   &H000080FF&
      Height          =   3375
      Left            =   3600
      TabIndex        =   31
      Top             =   2760
      Width           =   8295
      Begin VB.CheckBox chkInstallPolicy 
         BackColor       =   &H00404040&
         Caption         =   "Apply policies"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvPolices 
         Height          =   2535
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList ilDevices 
      Left            =   11160
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   4210752
      MaskColor       =   4210752
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Devices         "
      ForeColor       =   &H000080FF&
      Height          =   2655
      Left            =   3600
      TabIndex        =   35
      Top             =   0
      Width           =   8295
      Begin MSComctlLib.TreeView TV 
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4048
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "ilDevices"
         Appearance      =   0
      End
      Begin VB.Image Image3 
         Height          =   225
         Left            =   840
         MouseIcon       =   "frmPolicy.frx":0594
         MousePointer    =   99  'Custom
         Picture         =   "frmPolicy.frx":06E6
         Stretch         =   -1  'True
         ToolTipText     =   "Refresh"
         Top             =   -15
         Width           =   225
      End
      Begin VB.Image imgSelectNone 
         Height          =   240
         Left            =   7800
         MouseIcon       =   "frmPolicy.frx":0919
         MousePointer    =   99  'Custom
         Picture         =   "frmPolicy.frx":0A6B
         ToolTipText     =   "Disable all devices"
         Top             =   -15
         Width           =   240
      End
      Begin VB.Image imgSelectAll 
         Height          =   240
         Left            =   7440
         MouseIcon       =   "frmPolicy.frx":0B1B
         MousePointer    =   99  'Custom
         Picture         =   "frmPolicy.frx":0C6D
         ToolTipText     =   "Enable all devices"
         Top             =   -15
         Width           =   240
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Log"
      ForeColor       =   &H000080FF&
      Height          =   2895
      Left            =   3600
      TabIndex        =   34
      Top             =   6240
      Width           =   4935
      Begin VB.TextBox txtLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2535
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4560
         MouseIcon       =   "frmPolicy.frx":0D2A
         MousePointer    =   99  'Custom
         Picture         =   "frmPolicy.frx":0E7C
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Startup"
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   32
      Top             =   6240
      Width           =   3135
      Begin VB.CheckBox chkDisableWelcome 
         BackColor       =   &H00404040&
         Caption         =   "Disable interactive logon"
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   360
         MouseIcon       =   "frmPolicy.frx":0F13
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   240
         Value           =   1  'Checked
         Width           =   2625
      End
   End
   Begin VB.CommandButton cmdInstall 
      BackColor       =   &H00808080&
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8520
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Event logs"
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   240
      TabIndex        =   26
      Top             =   2760
      Width           =   3135
      Begin VB.TextBox txtEventsDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   2
         Text            =   "512"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtEventsSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   1
         Text            =   "512"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "days"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2280
         TabIndex        =   30
         Top             =   720
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "MB"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2280
         TabIndex        =   29
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Overwrite after"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Maximum log size"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Desctop"
      ForeColor       =   &H000080FF&
      Height          =   2295
      Left            =   240
      TabIndex        =   25
      Top             =   4080
      Width           =   3135
      Begin VB.CheckBox chkIconsHideAll 
         BackColor       =   &H00404040&
         Caption         =   "Hide all user icons"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":1065
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox chkIconNetwork 
         BackColor       =   &H00404040&
         Caption         =   "Hide My Network Places"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":11B7
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1770
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkIconIE 
         BackColor       =   &H00404040&
         Caption         =   "Hide Internet Explorer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":1309
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1410
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkIconMyDocuments 
         BackColor       =   &H00404040&
         Caption         =   "Hide My Documents"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":145B
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1050
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkIconsMyComputer 
         BackColor       =   &H00404040&
         Caption         =   "Hide My Computer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":15AD
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   690
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Screen saver"
      ForeColor       =   &H000080FF&
      Height          =   1575
      Left            =   360
      TabIndex        =   22
      Top             =   1080
      Width           =   3135
      Begin VB.CheckBox chkDefaultScreenSaver 
         BackColor       =   &H00404040&
         Caption         =   "Default screen saver"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":16FF
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1125
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkSSSecure 
         BackColor       =   &H00404040&
         Caption         =   "Require password"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":1851
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   780
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.TextBox txtScrSaverTimeout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2000
         TabIndex        =   3
         Text            =   "10"
         Top             =   320
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "minutes"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2475
         TabIndex        =   24
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Start screen saver after"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   345
         Width           =   1650
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Security clearance"
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   240
      TabIndex        =   21
      Top             =   0
      Width           =   3135
      Begin VB.ComboBox cmbClassification 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmPolicy.frx":19A3
         Left            =   240
         List            =   "frmPolicy.frx":19B6
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "User folders"
      ForeColor       =   &H000080FF&
      Height          =   2655
      Left            =   240
      TabIndex        =   20
      Top             =   6480
      Width           =   3135
      Begin VB.CheckBox chkMakePrivate 
         BackColor       =   &H00404040&
         Caption         =   "Make 'Private'"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   795
         MouseIcon       =   "frmPolicy.frx":1A09
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Папката ще е достъпна само за собственика й и администратора"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbDrives 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdMoveNow 
         BackColor       =   &H00808080&
         Caption         =   "Move now"
         Height          =   375
         Left            =   1320
         MouseIcon       =   "frmPolicy.frx":1B5B
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Премест папките на всички потребители"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox chkDesktop 
         BackColor       =   &H00404040&
         Caption         =   "Move Desktop"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":1CAD
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkMyDocuments 
         BackColor       =   &H00404040&
         Caption         =   "Move 'My Documents '"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         MouseIcon       =   "frmPolicy.frx":1DFF
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "in local drive"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1485
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ClassImageList As SP_CLASSIMAGELIST_DATA
Private IsOK As Boolean
Private Classification As String
Private DATAPATH As String

Private Sub cmbClassification_Click()
    Classification = cmbClassification.List(cmbClassification.ListIndex)
End Sub

Public Sub cmdInstall_Click()
    
    IsOK = True
    
    cmdInstall.Enabled = False
    txtLog.Text = ""

    Screen.MousePointer = vbHourglass
    
    Logme "Десктоп икони..." & SetDesktopIcons
    Logme "Скрийн сейвър..." & SetScreenSaver
    'Logme "Потребителски папки..." & SetPersonalFolders
    
    If IS_ADMIN = True And Command$ = "" Then
        Logme "Одитини записи..." & SetEventViewer
        Logme "Стартиране..." & SetStartUp
        Logme "Устройства..." & DeviceStatus
        Logme "Политики за сигурност..." & CreatePolicy
        Logme "Запазване на настойките..." & SaveSettings
    End If
        
    Screen.MousePointer = vbNormal
    
    cmdInstall.Enabled = True
    
    If IsOK = True Then
        Logme "Операцията приключи УСПЕШНО!"
    Else
        Logme "Операцията приключи с ГРЕШКИ!"
    End If
    
End Sub

Private Function SetScreenSaver() As String
On Error GoTo Err
Dim Result As Long
   
    SetScreenSaver = "OK"
        
    SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaverIsSecure", _
        Trim$(str$(chkSSSecure.Value)), REG_SZ
        
    SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveTimeOut", _
        Trim$(str$(Int(txtScrSaverTimeout.Text) * 60)), REG_SZ
    
    If chkDefaultScreenSaver.Value = 1 Or _
        QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE") = "" Then
            SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", _
                "C:\WINDOWS\system32\logon.scr", REG_SZ
    End If
    
    If Int(txtScrSaverTimeout.Text) > 0 Then
        SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveActive", "1", REG_SZ
    Else
        DeleteKeys HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", Result
    End If
    
    Exit Function
Err:
    SetScreenSaver = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Function SetStartUp() As String
On Error GoTo Err
Dim Result As Long
   
    SetStartUp = "OK"
      
    If chkDisableWelcome.Value = 1 Then
        SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "LogonType", "0", REG_DWORD
    Else
        SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "LogonType", "1", REG_DWORD
    End If
    Exit Function
Err:
    SetStartUp = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Function SetDesktopIcons() As String
'On Error GoTo Err
Dim Result As Long
    
    SetDesktopIcons = "OK"
    
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu", HKEY_CURRENT_USER
    
    If chkIconsHideAll.Value = 1 Then
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu", _
            "{00000000-0000-0000-0000-000000000000}", Trim$(str$(chkIconsHideAll.Value)), REG_DWORD
    Else
        DeleteKeys HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu", _
            "{00000000-0000-0000-0000-000000000000}", Result
    End If
    
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu", _
        "{450D8FBA-AD25-11D0-98A8-0800361B1103}", Trim$(str$(chkIconMyDocuments.Value)), REG_DWORD 'My Documents
    
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu", _
        "{20D04FE0-3AEA-1069-A2D8-08002B30309D}", Trim$(str$(chkIconsMyComputer.Value)), REG_DWORD 'My Computer
    
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu", _
        "{208D2C60-3AEA-1069-A2D7-08002B30309D}", Trim$(str$(chkIconNetwork.Value)), REG_DWORD 'Network Places
    
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu", _
        "{871C5380-42A0-1069-A2EA-08002B30309D}", Trim$(str$(chkIconIE.Value)), REG_DWORD 'Internet Explorer
    
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", HKEY_CURRENT_USER
    
    If chkIconsHideAll.Value = 1 Then
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", _
            "{00000000-0000-0000-0000-000000000000}", Trim$(str$(chkIconsHideAll.Value)), REG_DWORD
    Else
        DeleteKeys HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", _
            "{00000000-0000-0000-0000-000000000000}", Result
    End If
    
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", _
        "{450D8FBA-AD25-11D0-98A8-0800361B1103}", Trim$(str$(chkIconMyDocuments.Value)), REG_DWORD 'My Documents
    
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", _
        "{20D04FE0-3AEA-1069-A2D8-08002B30309D}", Trim$(str$(chkIconsMyComputer.Value)), REG_DWORD 'My Computer
    
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", _
        "{208D2C60-3AEA-1069-A2D7-08002B30309D}", Trim$(str$(chkIconNetwork.Value)), REG_DWORD 'Network Places
    
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", _
        "{871C5380-42A0-1069-A2EA-08002B30309D}", Trim$(str$(chkIconIE.Value)), REG_DWORD 'Internet Explorer
    Exit Function

Err:
    SetDesktopIcons = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Function SetEventViewer() As String
On Error GoTo Err
    
    SetEventViewer = "OK"
    
    SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Eventlog\Application", _
        "MaxSize", Int(txtEventsSize.Text) * 1048576, REG_DWORD 'Application Max Size
    
    SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Eventlog\Application", _
        "Retention", Int(txtEventsDays.Text) * 86400, REG_DWORD 'Application Days
    
    SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Eventlog\Security", _
        "MaxSize", Int(txtEventsSize.Text) * 1048576, REG_DWORD 'Security Max Size
    
    SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Eventlog\Security", _
        "Retention", Int(txtEventsDays.Text) * 86400, REG_DWORD 'Security Days
    
    SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Eventlog\System", _
        "MaxSize", Int(txtEventsSize.Text) * 1048576, REG_DWORD 'System Max Size
    
    SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Eventlog\System", _
        "Retention", Int(txtEventsDays.Text) * 86400, REG_DWORD 'System Days

    Exit Function
Err:
    SetEventViewer = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Function SaveSettings() As String
On Error GoTo Err
  
    SaveSettings = "OK"
    
    'Classification
    writeINI INIPATH, "Classification", "Classification", cmbClassification.ListIndex

    'Screen Saver
    writeINI INIPATH, "Screen Saver", "ScreenSaverIsSecure", Trim$(str$(chkSSSecure.Value))
    writeINI INIPATH, "Screen Saver", "ScreenSaveTimeOut", Trim$(txtScrSaverTimeout.Text)
    writeINI INIPATH, "Screen Saver", "DefaultScreenSaver", Trim$(str$(chkDefaultScreenSaver.Value))
    
    'Personal Folders
    writeINI INIPATH, "Personal Folders", "MoveMyDocuments", Trim$(str$(chkMyDocuments.Value))
    writeINI INIPATH, "Personal Folders", "MoveDesktop", Trim$(str$(chkDesktop.Value))
    writeINI INIPATH, "Personal Folders", "MoveTo", cmbDrives.ListIndex
    
    'Desktop Icons
    writeINI INIPATH, "Desktop Icons", "HideUserIcons", Trim$(str$(chkIconsHideAll.Value))
    writeINI INIPATH, "Desktop Icons", "HideMyDocuments", Trim$(str$(chkIconMyDocuments.Value))
    writeINI INIPATH, "Desktop Icons", "HideMyComputer", Trim$(str$(chkIconsMyComputer.Value))
    writeINI INIPATH, "Desktop Icons", "HideNetwork", Trim$(str$(chkIconNetwork.Value))
    writeINI INIPATH, "Desktop Icons", "HideIE", Trim$(str$(chkIconIE.Value))
    
    'Event Viewer
    writeINI INIPATH, "Event Viewer", "MaxSize", Int(txtEventsSize.Text)
    writeINI INIPATH, "Event Viewer", "Retention", Int(txtEventsDays.Text)
    
    'Office Macros
    'writeINI INIPATH, "Office Macros", "Install", Trim(Str(chkInstallMacros.Value))
    
    'Security Policy
    writeINI INIPATH, "Security Policy", "Install", Trim$(str$(chkInstallPolicy.Value))
    
    'StartUp
    writeINI INIPATH, "StartUp", "WelcomeScreen", Trim$(str$(chkDisableWelcome.Value))
    
    Exit Function
Err:
    SaveSettings = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Function ReadSettings() As String
On Error GoTo Err
    
    ReadSettings = "OK"
    
    'Classification
    cmbClassification.ListIndex = sGetINI(INIPATH, "Classification", "Classification", "0")
    
    'Screen Saver
    chkSSSecure.Value = sGetINI(INIPATH, "Screen Saver", "ScreenSaverIsSecure", "1")
    txtScrSaverTimeout.Text = sGetINI(INIPATH, "Screen Saver", "ScreenSaveTimeOut", "10")
    chkDefaultScreenSaver.Value = sGetINI(INIPATH, "Screen Saver", "DefaultScreenSaver", "1")
      
    'Personal Folders
    chkMyDocuments.Value = sGetINI(INIPATH, "Personal Folders", "MoveMyDocuments", "1")
    chkDesktop.Value = sGetINI(INIPATH, "Personal Folders", "MoveDesktop", "1")
    cmbDrives.ListIndex = sGetINI(INIPATH, "Personal Folders", "MoveTo", cmbDrives.ListCount - 1)
      
    'Desktop Icons
    chkIconsHideAll.Value = sGetINI(INIPATH, "Desktop Icons", "HideUserIcons", "0")
    chkIconMyDocuments.Value = sGetINI(INIPATH, "Desktop Icons", "HideMyDocuments", "0")
    chkIconsMyComputer.Value = sGetINI(INIPATH, "Desktop Icons", "HideMyComputer", "0")
    chkIconNetwork.Value = sGetINI(INIPATH, "Desktop Icons", "HideNetwork", "1")
    chkIconIE.Value = sGetINI(INIPATH, "Desktop Icons", "HideIE", "1")
    
    'Event Viewer
    txtEventsSize.Text = sGetINI(INIPATH, "Event Viewer", "MaxSize", "512")
    txtEventsDays.Text = sGetINI(INIPATH, "Event Viewer", "Retention", "180")
    
    'Office Macros
    'chkInstallMacros.Value = sGetINI(INIPATH, "Office Macros", "Install", "1")
    
    'Security Policy
    chkInstallPolicy.Value = sGetINI(INIPATH, "Security Policy", "Install", "1")
        
    'StartUp
    chkDisableWelcome.Value = sGetINI(INIPATH, "StartUp", "WelcomeScreen", "1")
      
    Exit Function
Err:
    ReadSettings = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Function CreatePolicy() As String
On Error GoTo Err

Dim sPath As String
Dim sLegalNoticeText As String
    
    CreatePolicy = "OK"
    sPath = DATAPATH & "\policy.inf"
   
    writeINI sPath, "Unicode", "Unicode", "yes"
    writeINI sPath, "System Access", "MinimumPasswordAge", "0"
    writeINI sPath, "System Access", "MaximumPasswordAge", "90"
    writeINI sPath, "System Access", "MinimumPasswordLength", "8"
    writeINI sPath, "System Access", "PasswordComplexity", "1"
    writeINI sPath, "System Access", "PasswordHistorySize", "10"
    writeINI sPath, "System Access", "LockoutBadCount", "5"
    writeINI sPath, "System Access", "ResetLockoutCount", "10"
    writeINI sPath, "System Access", "LockoutDuration", "10"
    writeINI sPath, "System Access", "EnableAdminAccount", "1"
    writeINI sPath, "System Access", "EnableGuestAccount", "0"
    
    writeINI sPath, "Event Audit", "AuditSystemEvents", "3"
    writeINI sPath, "Event Audit", "AuditLogonEvents", "3"
    writeINI sPath, "Event Audit", "AuditObjectAccess", "3"
    writeINI sPath, "Event Audit", "AuditPrivilegeUse", "3"
    writeINI sPath, "Event Audit", "AuditPolicyChange", "3"
    writeINI sPath, "Event Audit", "AuditAccountManage", "3"
    writeINI sPath, "Event Audit", "AuditProcessTracking", "3"
    writeINI sPath, "Event Audit", "AuditDSAccess", "3"
    writeINI sPath, "Event Audit", "AuditAccountLogon", "3"
 
    sLegalNoticeText = "Автоматизираната информационна система" & Chr$(34) & "," & Chr$(34) & _
        " която използвате е с ниво на класификация " & UCase$(cmbClassification.Text) & _
            ",Всички ваши последващи действия трябва да отговарят на изискванията на " & _
            "Закона за защита на класифицираната информация" & Chr$(34) & "," & Chr$(34) & _
                " поднормативните актове за неговото прилагане и вътрешноведомствените документи."
                
    writeINI sPath, "Registry Values", "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableCAD", "4,0"
    writeINI sPath, "Registry Values", "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DontDisplayLastUserName", "4,1"
    writeINI sPath, "Registry Values", "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeCaption", "1,ВНИМАНИЕ !!!"
    writeINI sPath, "Registry Values", "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeText", _
        "7," & sLegalNoticeText
    writeINI sPath, "Registry Values", "MACHINE\System\CurrentControlSet\Control\Session Manager\Memory Management\ClearPageFileAtShutdown", "4,1"
    
    writeINI sPath, "Privilege Rights", "SeNetworkLogonRight", "*S-1-5-32-544"
    writeINI sPath, "Privilege Rights", "SeSystemtimePrivilege ", "*S-1-5-32-544"
    writeINI sPath, "Privilege Rights", "SeShutdownPrivilege ", "*S-1-5-32-544,*S-1-5-32-545"
    
    writeINI sPath, "Version", "signature", Chr$(34) & "$CHICAGO$" & Chr$(34)
    writeINI sPath, "Version", "Revision", "1"
    
    If chkInstallPolicy.Value = 1 Then
        ShellWait "cmd /c secedit /configure /CFG " & Chr$(34) & sPath & Chr$(34) & " /DB " & _
             Chr$(34) & DATAPATH & "\database.sdb" & Chr$(34) & ">>" & Chr$(34) & DATAPATH & _
                "\SecurityPolicy.log" & Chr$(34), vbHide
    End If
       
    If FExists(DATAPATH & "\policy.inf") Then Kill DATAPATH & "\policy.inf"
    If FExists(DATAPATH & "\database.sdb") Then Kill DATAPATH & "\database.sdb"
    Exit Function
Err:
    CreatePolicy = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Function SetPersonalFolders() As String
On Error GoTo Err
       
    SetPersonalFolders = "OK"
    
    If chkMyDocuments.Value = 1 Then
        
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", _
            "Personal", cmbDrives.List(cmbDrives.ListIndex) & "My Documents\%USERNAME%", REG_EXPAND_SZ
            
        If DirExists("c:\Documents and Settings\" & USER_NAME & "\My Documents") = True Then
            MakeSureDirectoryPathExists cmbDrives.List(cmbDrives.ListIndex) & "My Documents\" & USER_NAME & "\"
            If CopySource("c:\Documents and Settings\" & USER_NAME & "\My Documents", _
                cmbDrives.List(cmbDrives.ListIndex) & "My Documents\" & USER_NAME, True) = True Then
                    Logme USER_NAME & "'s 'My Documents'...OK"
            Else
                Logme USER_NAME & "'s 'My Documents'...ГРЕШКА"
                SetPersonalFolders = "ГРЕШКА"
                IsOK = False
            End If
        End If
    End If
    
    If chkDesktop.Value = 1 Then
        
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", _
            "Desktop", cmbDrives.List(cmbDrives.ListIndex) & "My Documents\%USERNAME%\Desktop", REG_EXPAND_SZ
        
        If DirExists("c:\Documents and Settings\" & USER_NAME & "\Desktop") = True Then
            
            MakeSureDirectoryPathExists cmbDrives.List(cmbDrives.ListIndex) & "My Documents\" & USER_NAME & "\"
            
            If CopySource("c:\Documents and Settings\" & USER_NAME & "\Desktop", _
                cmbDrives.List(cmbDrives.ListIndex) & "My Documents\" & USER_NAME & "\Desktop", True) = True Then
                    Logme USER_NAME & "'s 'Desktop'...OK"
            Else
                Logme USER_NAME & "'s 'Desktop'...ГРЕШКА"
                SetPersonalFolders = "ГРЕШКА"
                IsOK = False
            End If
        End If
    End If
    Exit Function
Err:
    SetPersonalFolders = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Sub cmdMoveNow_Click()
    MoveFilderContent
End Sub

Private Function MoveFilderContent() As String
On Error GoTo Err

    Dim Mypath As String, MyName As String, iCount As Integer
    
    MoveFilderContent = "OK"
    
    iCount = 0
    Mypath = "c:\Documents and Settings\"    ' Set the path.
    MyName = Dir$(Mypath, vbDirectory)   ' Retrieve the first entry.
    
    Do While MyName <> ""   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If MyName <> "." And MyName <> ".." Then
          ' Use bitwise comparison to make sure MyName is a directory.
          If (GetAttr(Mypath & MyName) And vbDirectory) = vbDirectory Then
             If MyName <> "All Users" Then
             
                If DirExists("c:\Documents and Settings\" & MyName & "\My Documents") = True Then
                    MakeSureDirectoryPathExists cmbDrives.List(cmbDrives.ListIndex) & "My Documents\" & MyName
                    If CopySource("c:\Documents and Settings\" & MyName & "\My Documents", _
                        cmbDrives.List(cmbDrives.ListIndex) & "My Documents\" & MyName, True) = True Then
                        Logme MyName & "'s 'My Documents'...OK"
                    Else
                        Logme MyName & "'s 'My Documents'...ГРЕШКА"
                    End If
                End If
                 
                If DirExists("c:\Documents and Settings\" & MyName & "\Desktop") = True Then
                    MakeSureDirectoryPathExists cmbDrives.List(cmbDrives.ListIndex) & "My Documents\" & MyName & "\Desktop\"
                    If CopySource("c:\Documents and Settings\" & MyName & "\Desktop", _
                        cmbDrives.List(cmbDrives.ListIndex) & "My Documents\" & MyName, True) = True Then
                        Logme MyName & "'s 'Desktop'...OK"
                    Else
                        Logme MyName & "'s 'Desktop'...ГРЕШКА"
                    End If
                End If
             
             End If
             iCount = iCount + 1
          End If   ' it represents a directory.
       End If
       MyName = Dir   ' Get next entry.
    Loop
    MsgBox "Потребителските папки са преместени", vbInformation, "Администраторски панел"
    Exit Function
Err:
    MoveFilderContent = "Грешка"
    Resume Next
End Function

Private Sub FillTV()
Dim nodX As Node
Dim nodR As String
Dim x As Long

    TV.Nodes.Clear

    On Local Error GoTo Fin
    
    For x = 0 To UBound(HW_DEVICES)
        With HW_DEVICES(x)
            Select Case LCase$(.Class)
                Case "ports"
                    nodR = "Портове"
                Case "fdc"
                    nodR = "Флопи контролери"
                Case "cdrom"
                    nodR = "Оптични устройства"
                Case "modem"
                    nodR = "Модеми"
                Case "usb"
                    nodR = "USB контролери"
                Case Else
                    GoTo DoNext
            End Select
            
            If Not NodeExists(TV, nodR) Then
                Set nodX = TV.Nodes.Add(, , nodR, nodR, GetClassImageListIndex(ClassImageList, .ClassGuid))
                nodX.Checked = True
                nodX.BackColor = &H404040
                nodX.ForeColor = vbWhite
                nodX.Bold = True
                Set nodX = Nothing
            End If
            
            If LCase$(.Class) = "usb" And InStr(LCase$(.DeviceDesc), "controller") = 0 Then GoTo DoNext
            
            Set nodX = TV.Nodes.Add(nodR, tvwChild, "H" & .devInst, IIf(.FriendlyName <> "", .FriendlyName, .DeviceDesc), GetClassImageListIndex(ClassImageList, .ClassGuid))
            nodX.Tag = .Index
            nodX.Checked = .Enabled
            nodX.BackColor = &H404040
            nodX.ForeColor = vbWhite
            
            If .Enabled = False Then
                nodX.ForeColor = vbRed
                nodX.Parent.ForeColor = vbRed
                nodX.Parent.Checked = False
                nodX.Parent.Expanded = True
            End If
DoNext:
            Set nodX = Nothing
        End With
        
    Next x
    
Fin:
End Sub

Private Function NodeExists(TV As TreeView, ByVal sKey As String) As Boolean
   Dim nd As Node
   On Error Resume Next
   Set nd = TV.Nodes(sKey)
   NodeExists = (Err = 0)
   Set nd = Nothing
End Function

Private Sub Form_Activate()
    DATAPATH = App.Path
    SetTVBackColor TV, &H404040
    FillTV
End Sub

Private Sub FillPolicyListView()
    
    Dim itmX As ListItem
    
    modFillListView.MY_ListView = lvPolices
    
    AddListItem "10 password remembered", "Password History Size"
    AddListItem "0 days", "Minimum Password Age"
    AddListItem "90 days", "Maximum Password Age"
    AddListItem "8 characters", "Minimum Password Length"
    AddListItem "Enabled", "Password Complexity"
    
    AddListItem "", ""
    AddListItem "10 minutes", "Account lockout duration"
    AddListItem "5 invalid logon attempts", "Account lockout threshold"
    AddListItem "10 minutes", "Reset account lockout counter after"
    
    AddListItem "", ""
    AddListItem "Success, Failure", "Audit System Events"
    AddListItem "Success, Failure", "Audit Logon Events"
    AddListItem "Success, Failure", "Audit Object Access"
    AddListItem "Success, Failure", "Audit Privilege Use"
    AddListItem "Success, Failure", "Audit Policy Change"
    AddListItem "Success, Failure", "Audit Account Management"
    AddListItem "Success, Failure", "Audit Process Tracking"
    AddListItem "Success, Failure", "Audit Directory Service Access"
    AddListItem "Success, Failure", "Audit Account Logon Events"
    
    AddListItem "", ""
    AddListItem "Administrators", "Access this computer from the network"
    AddListItem "Administrators, Users", "Shut down the sistem"
    AddListItem "Administrators", "Change the system time"
    
    AddListItem "", ""
    AddListItem "Enabled", "Account: Administrator account status"
    AddListItem "Disabled", "Account: Guest account status"
    AddListItem "Enabled", "Interactive login:Do not display last user name"
    AddListItem "Disabled", "Interactive login: Do not require CTRL + ALD +DEL"
    AddListItem "ВНИМАНИЕ !!!", "Interactive logon: Message title for users attempting to log on"
    AddListItem "Предупредителен текст", "Interactive logon: Message text for users attempting to log on"
    AddListItem "Enabled", "Shutdown: Clear virtual memory pagefile"
    AutoSizeListViewColumns lvPolices
End Sub

Private Sub Form_Load()
        
    
    GetDrives
    ReadSettings
    Classification = cmbClassification.List(cmbClassification.ListIndex)
    IsSecondDrive
    
    DoEvents

    ClassImageList = GetClassImageList
    
    FillImageListWithClassImageList ClassImageList, ilDevices
    
    With lvPolices
        .ColumnHeaders.Add , , "Policy"
        .ColumnHeaders.Add , , "Value"
    End With
    
    FillPolicyListView
End Sub

Private Sub IsSecondDrive()
    If cmbDrives.ListCount = 0 Then
        chkDesktop.Enabled = False
        chkDesktop.Value = 0
        
        chkMyDocuments.Enabled = False
        chkMyDocuments.Value = 0
        
        frmPolicy!cmdMoveNow.Enabled = False
        cmbDrives.Enabled = False
       
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyClassImageList ClassImageList
    SaveSettings
    Unload Me
End Sub

Private Sub Image1_Click()
    Shell ("explorer " & DATAPATH)
End Sub

Private Sub Image3_Click()
    GetDevProp "", HW_DEVICES
    FillTV
End Sub

Private Sub imgSelectAll_Click()
Dim oNode As Node
    For Each oNode In TV.Nodes
        oNode.Checked = True
        Call TV_NodeCheck(oNode)
    Next
End Sub

Private Sub imgSelectNone_Click()
Dim oNode As Node
    For Each oNode In TV.Nodes
        oNode.Checked = False
        Call TV_NodeCheck(oNode)
    Next
End Sub

Private Sub Logme(sAction As String)
    If Trim$(txtLog.Text) <> "" Then txtLog.Text = vbCrLf & txtLog.Text

    txtLog.Text = ">> " & sAction & txtLog.Text
    txtLog.Refresh
End Sub

Private Sub TV_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim oChild As Node
    
    If Node.Parent Is Nothing Then
        Set oChild = Node.Child
        Do Until oChild Is Nothing
            oChild.Checked = Node.Checked
            oChild.ForeColor = IIf(oChild.Checked, vbWhite, vbRed)
            Set oChild = oChild.Next
        Loop
    Else
        Node.Parent.Checked = True
        Set oChild = Node.FirstSibling
        Do Until oChild Is Nothing
            Node.Parent.Checked = Node.Parent.Checked And oChild.Checked
            Set oChild = oChild.Next
        Loop
        Node.Parent.ForeColor = IIf(Node.Parent.Checked, vbWhite, vbRed)
    End If
    
    Node.ForeColor = IIf(Node.Checked, vbWhite, vbRed)

End Sub

Private Function DeviceStatus() As String
On Error GoTo Err
Dim oParent As Node, oChild As Node

    DeviceStatus = "OK"
    
    GetDevProp "", HW_DEVICES
    
    For Each oParent In TV.Nodes
        Set oChild = oParent.Child
        Do Until oChild Is Nothing
            If HW_DEVICES(oChild.Tag).Enabled <> oChild.Checked Then
                EnableDevice oChild.Tag, oChild.Checked
                Logme oChild & "..." & IIf(oChild.Checked, "разрешено", "забранено")
            End If
            Set oChild = oChild.Next
            DoEvents
        Loop
        DoEvents
    Next
Exit Function
Err:
    DeviceStatus = "ГРЕШКА"
    IsOK = False
    Resume Next
End Function

Private Sub txtEventsDays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtEventsDays = Int(txtEventsDays) + 1
            KeyCode = 0
        Case 40
            If Int(txtEventsDays) > 0 Then
                txtEventsDays = Int(txtEventsDays) - 1
                KeyCode = 0
            End If
    End Select
End Sub

Private Sub txtEventsDays_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtEventsDays_LostFocus()
    If Trim$(txtEventsDays.Text) = "" Then txtEventsDays.Text = "0"
End Sub

Private Sub txtEventsSize_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtEventsSize = Int(txtEventsSize) + 1
            KeyCode = 0
        Case 40
            If Int(txtEventsSize) > 0 Then
                txtEventsSize = Int(txtEventsSize) - 1
                KeyCode = 0
            End If
    End Select
End Sub

Private Sub txtEventsSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtEventsSize_LostFocus()
    If Trim$(txtEventsSize.Text) = "" Then txtEventsSize.Text = "0"
End Sub

Private Sub txtScrSaverTimeout_Change()
    
    Dim a As Integer
    
    a = Int(Int("0" & txtScrSaverTimeout) > 0)
    
    chkSSSecure.Enabled = Int("0" & txtScrSaverTimeout) > 0
    chkDefaultScreenSaver.Enabled = CInt("0" & txtScrSaverTimeout > 0)
End Sub

Private Sub txtScrSaverTimeout_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtScrSaverTimeout = Int(txtScrSaverTimeout) + 1
            KeyCode = 0
        Case 40
            If Int(txtScrSaverTimeout) > 0 Then
                txtScrSaverTimeout = Int(txtScrSaverTimeout) - 1
                KeyCode = 0
            End If
    End Select
End Sub

Private Sub txtScrSaverTimeout_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtScrSaverTimeout_LostFocus()
    If Trim$(txtScrSaverTimeout.Text) = "" Then txtScrSaverTimeout.Text = "0"
End Sub
