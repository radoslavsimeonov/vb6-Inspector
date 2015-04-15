Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    
    Dim x As Integer
    
    COMPUTER_NAME = CreateObject("WScript.Network").ComputerName
    DOMAIN_NAME = CreateObject("WScript.Network").UserDomain
    USER_NAME = CreateObject("WScript.Network").username
    IS_ADMIN = CBool(IsNTAdmin(ByVal 0&, ByVal 0&))
    OS_VERSION = GetOSVersion
    INIPATH = App.Path & "\SecurityPolicy.ini"
    
    ReDim REG_USERS(0)
    
    ReDim SW_SERVICES(0)
    ReDim SW_APPLICATIONS(0)
    ReDim SW_LOCAL_USERS(0)
    ReDim SW_SHARED_FOLDERS(0)
    ReDim SW_START_COMMANDS(0)
    ReDim SW_LICENSES(0)
    ReDim SW_SECURITY_PRODUCTS.AntiVirus(0)
    ReDim SW_SECURITY_PRODUCTS.Firewall(0)
    ReDim SW_SECURITY_PRODUCTS.Spyware(0)
    
    ReDim HW_MOTHERBOARD.Floppy(0)
    ReDim HW_MOTHERBOARD.Ports(0)
    ReDim HW_RAM_MEMORY.Banks(0)
    ReDim HW_DEVICES(0)
    ReDim HW_CDROMS(0)
    ReDim HW_HARDDISKS(0)
    ReDim HW_LOGICAL_DRIVES(0)
    ReDim HW_MONITORS(0)
    ReDim HW_NETWORK_ADAPTERS(0)
    ReDim HW_PRINTERS(0)
    ReDim HW_VIDEO_ADAPTERS(0)
    ReDim HW_SOUND_DEVICES(0)
    
    SW_OPERATING_SYSTEM.ActivationStatus = -1
    
    'GoTo DontLoad
        
    Splash "Operating system..."
    SW_OPERATING_SYSTEM = GetOperatingSystem
    
    If Not NotAdminAlert Then
        If IsFormLoaded("frmSplash") Then Unload frmSplash
        Exit Sub
    End If

    With frmSplash
        .lblAction.BorderStyle = 0
        .ProgressBar1.Min = 1
        .ProgressBar1.max = 21
        .Show
    End With

    Splash "Windows services..."
    SW_SERVICES = EnumServices
    
    Splash "Security products..."
    SW_SECURITY_PRODUCTS = GetSecuityProducts
    
    Splash "Installed applications..."
    SW_APPLICATIONS = EnumApplications
    
    Splash "Software licenses..."
    SW_LICENSES = GetSoftwareLicenses
            
    Splash "Local user accounts..."
    SW_LOCAL_USERS = EnumAccounts
    
    Splash "Shared folders..."
    EnumSharedFolders2 SW_SHARED_FOLDERS
    
    Splash "Startup programs..."
    SW_START_COMMANDS = EnumStartUpCommands

    Splash "CD-ROMS / DVD-ROMS..."
    HW_CDROMS = EnumCDROMs

    Splash "Processors..."
    HW_CPU = GetCPU
       
    Splash "Hard disk drives..."
    HW_HARDDISKS = EnumHardDrives
    
    Splash "RAM Memory..."
    HW_RAM_MEMORY = EnumRAMMemory
     
    Splash "Monitors..."
    HW_MONITORS = EnumMonitors
    
    Splash "Motherboard..."
    HW_MOTHERBOARD = GetMotherBoard
    
    Splash "Network adapters..."
    HW_NETWORK_ADAPTERS = EnumNetworkAdapters

    Splash "Printers..."
    HW_PRINTERS = EnumPrinters
    
    Splash "Video adapters..."
    HW_VIDEO_ADAPTERS = EnumVideoAdapters

    Splash "Human Interface Devices..."
    HW_HID = GetHID
    
    Splash "Devices..."
    GetDevProp "", HW_DEVICES
    
    Splash "Sound devices..."
    HW_SOUND_DEVICES = EnumSoundDevices
    
    Splash "Finish :)"
    DoEvents

    If IsFormLoaded("frmSplash") Then Unload frmSplash

DontLoad:

    Load MDIMain
    
    MDIMain.Show
End Sub

Public Function CheckRegionalSetting()
    If SW_OPERATING_SYSTEM.CodeSet <> "1251" Then MsgBox "NE E BG"
End Function

Private Sub Splash(sAction As String)
    
    Static cnt As Integer
    
    If Not IsFormLoaded("frmSplash") Then Exit Sub
    
    frmSplash.lblAction.Caption = sAction
    cnt = cnt + 1
    frmSplash.ProgressBar1.Value = cnt
    DoEvents
    
End Sub
