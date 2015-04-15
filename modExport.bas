Attribute VB_Name = "modExport"
Option Explicit

Public Enum eExportFile
    AsXML
    AsJSON
    AsHTML
End Enum

Public Function ExportData(ByVal exportAs As eExportFile) As Boolean
          
    Select Case exportAs
    
        Case eExportFile.AsXML
            modXML.ClearXMLFile
            GenerateReport
            ExportData = modXML.SaveFileAsXML
        Case eExportFile.AsJSON
        
        Case eExportFile.AsHTML
    
    End Select

End Function

Private Sub GenerateReport()
    
    Dim i As Integer
    Dim j As Integer
    
    AddRow Root, COMPUTER_NAME & _
           " ReportDate='" & Now & "' CurrentUser='" & CurrentUser & _
           "' Domain='" & DOMAIN_NAME & "'", , False
        
        ' REG_WORKSTATION
        AddRow Element, "REG_WORKSTATION", , False
            With REG_WORKSTATION
                    AddRow Entry, "CaseStickers", Join(.CaseStickers, ",")
                    AddRow Entry, "InventaryNo", .InventaryNo
                    AddRow Entry, "InventaryDate", .InventaryDate
                    AddRow Entry, "BookNo", .BookNo
                    AddRow Entry, "BookDate", .BookDate
                    AddRow Entry, "Classification", .Classification
                    AddRow Entry, "SocketSKS", .SocketSKS
            End With
        AddRow Element, "REG_WORKSTATION"
        
        ' REG_USERS
        AddRow Element, "REG_USERS", , False
            For i = 0 To UBound(REG_USERS)
                If REG_USERS(i).username <> vbNullString Then
                    AddRow SubElement, "User", , False
                    With REG_USERS(i)
                        AddRow Entry, "username", .username
                        AddRow Entry, "Rank", .Rank
                        AddRow Entry, "FirstName", .FirstName
                        AddRow Entry, "MiddleName", .MiddleName
                        AddRow Entry, "SurName", .SurName
                        AddRow Entry, "IsOwner", .IsOwner
                        AddRow Entry, "IsNavy", .IsNavy
                        AddRow Entry, "Email", .Email
                        AddRow Entry, "Function", .Function
                        AddRow Entry, "Department", .Department
                        AddRow Entry, "Building", .Building
                        AddRow Entry, "Room", .Room
                        AddRow Entry, "Phone", .Phone
                        AddRow Entry, "UserServices", , False
                            For j = 0 To UBound(.UserServices)
                                If .UserServices(j).ServiceName <> vbNullString Then
                                    AddRow Entry1, "Service", , False
                                    With .UserServices(j)
                                        AddRow Entry2, "ServiceName", .ServiceName
                                        AddRow Entry2, "ServiceAddress", .ServiceAddress
                                        AddRow Entry2, "ServiceValue", .ServiceValue
                                        AddRow Entry2, "ServicePeriod", .ServicePeriod
                                    End With
                                    AddRow SubElement, "Service"
                                End If
                            Next j
                        AddRow Entry, "UserServices"
                    End With
                    AddRow SubElement, "User"
                End If
            Next i
        AddRow Element, "REG_USERS"
        
        ' SW_OPERATING_SYSTEM
        AddRow Element, "SW_OPERATING_SYSTEM", , False
            With SW_OPERATING_SYSTEM
                If .Caption <> vbNullString Then
                    AddRow SubElement, "Caption", .Caption
                    AddRow SubElement, "Architecture", .Architecture
                    AddRow SubElement, "BuildNumber", .BuildNumber
                    AddRow SubElement, "CodeSet", .CodeSet
                    AddRow SubElement, "CSDVersion", .CSDVersion
                    AddRow SubElement, "CountryCode", .CountryCode
                    AddRow SubElement, "CSName", .CSName
                    AddRow SubElement, "Domain", .Domain
                    AddRow SubElement, "CurrentTimeZone", .CurrentTimeZone
                    AddRow SubElement, "InstallDate", .InstallDate
                    AddRow SubElement, "Locale", .Locale
                    AddRow SubElement, "Organization", .Organization
                    AddRow SubElement, "OSLanugage", .OSLanugage
                    AddRow SubElement, "Primary", .Primary
                    AddRow SubElement, "ProductType", GetProductType(.ProductType)
                    AddRow SubElement, "RegisteredUser", .RegisteredUser
                    AddRow SubElement, "SPMajorVersion", .SPMajorVersion
                    AddRow SubElement, "SPMinorVersion", .SPMinorVersion
                    AddRow SubElement, "SystemDrive", .SystemDrive
                    AddRow SubElement, "SystemDirectory", .SystemDirectory
                    AddRow SubElement, "WindowsDirecorty", .WindowsDirecorty
                    AddRow SubElement, "Version", .Version
                    AddRow SubElement, "ActivationStatus", .ActivationStatus
                End If
            End With
        AddRow Element, "SW_OPERATING_SYSTEM"
        
        ' SW_LICENSES
        AddRow Element, "SW_LICENSES", , False
            For i = 0 To UBound(SW_LICENSES)
                If SW_LICENSES(i).Product <> vbNullString Then
                    AddRow SubElement, "License", , False
                    With SW_LICENSES(i)
                        AddRow Entry, "Product", .Product
                        AddRow Entry, "CDKey", .CDKey
                    End With
                    AddRow SubElement, "License"
                End If
            Next i
        AddRow Element, "SW_LICENSES"
        
        ' SW_APPLICATIONS
        AddRow Element, "SW_APPLICATIONS", , False
            For i = 0 To UBound(SW_APPLICATIONS)
                With SW_APPLICATIONS(i)
                    If .AppName <> vbNullString Then
                        AddRow SubElement, "Application", , False
                        
                        AddRow Entry, "AppName", .AppName
                        AddRow Entry, "Publisher", .Publisher
                        AddRow Entry, "InstallLocation", .InstallLocation
                        AddRow Entry, "Version", .Version
                        AddRow Entry, "UninstallString", .UninstallString
                        AddRow Entry, "ModifyString", .ModifyString
                        AddRow Entry, "InstalledOn", .InstalledOn
                        AddRow Entry, "AppBits", .AppBits
                        
                        AddRow SubElement, "Application"
                    End If
                End With
                
            Next i
        AddRow Element, "SW_APPLICATIONS"
                
        ' SW_START_COMMANDS
        AddRow Element, "SW_START_COMMANDS", , False
            For i = 0 To UBound(SW_START_COMMANDS)
                With SW_START_COMMANDS(i)
                    If .Command <> vbNullString Then
                        AddRow SubElement, "StartCommand", , False
                        
                        AddRow Entry, "CommandName", .CommandName
                        AddRow Entry, "CommandNameShort", .CommandNameShort
                        AddRow Entry, "Command", .Command
                        AddRow Entry, "Location", .Location
                        AddRow Entry, "Architecture", IIf(.Architecture = x64, "x64", "x86")
                        AddRow Entry, "UserRange", .UserRange
                        AddRow Entry, "Vendor", .Vendor
                    
                        AddRow SubElement, "StartCommand"
                    End If
                End With
            Next i
        AddRow Element, "SW_START_COMMANDS"
        
        ' SW_SHARED_FOLDERS
        AddRow Element, "SW_SHARED_FOLDERS", , False
            For i = 0 To UBound(SW_SHARED_FOLDERS)
                With SW_SHARED_FOLDERS(i)
                    If .FolderPath <> vbNullString Then
                        AddRow SubElement, "SharedFolder", , False
                    
                        AddRow Entry, "ShareName", .ShareName
                        AddRow Entry, "FolderPath", .FolderPath
                        AddRow Entry, "Description", .Description
                        AddRow Entry, "MaximumAllowed", .MaximumAllowed
                    
                        AddRow SubElement, "SharedFolder"
                    End If
                End With
            Next i
        AddRow Element, "SW_SHARED_FOLDERS"
        
        ' SW_SECURITY_PRODUCTS
        AddRow Element, "SW_SECURITY_PRODUCTS", , False
            
            For i = 0 To UBound(SW_SECURITY_PRODUCTS.AntiVirus)
                
                If SW_SECURITY_PRODUCTS.AntiVirus(i).ProductName <> vbNullString Then
                    AddRow SubElement, "Antivirus", , False
                    With SW_SECURITY_PRODUCTS.AntiVirus(i)
                        AddRow Entry, "ProductName", .ProductName
                        AddRow Entry, "Publisher", .Publisher
                        AddRow Entry, "ProductVersion", .ProductVersion
                        AddRow Entry, "ProductGUID", .ProductGUID
                        AddRow Entry, "RTPStatus", Abs(.RTPStatus)
                        AddRow Entry, "UpToDate", Abs(.UpToDate)
                        AddRow Entry, "Enabled", Abs(.Enabled)
                        AddRow Entry, "PathToProduct", .PathToProduct
                        AddRow Entry, "ProductState", .ProductState
                    End With
                    AddRow SubElement, "Antivirus"
                End If
            Next i
                        
            For i = 0 To UBound(SW_SECURITY_PRODUCTS.Firewall)
                If SW_SECURITY_PRODUCTS.Firewall(i).ProductName <> vbNullString Then
                    AddRow SubElement, "Firewall", , False
                    With SW_SECURITY_PRODUCTS.Firewall(i)
                        AddRow Entry, "ProductName", .ProductName
                        AddRow Entry, "Publisher", .Publisher
                        AddRow Entry, "ProductVersion", .ProductVersion
                        AddRow Entry, "ProductGUID", .ProductGUID
                        AddRow Entry, "RTPStatus", Abs(.RTPStatus)
                        AddRow Entry, "UpToDate", Abs(.UpToDate)
                        AddRow Entry, "Enabled", Abs(.Enabled)
                        AddRow Entry, "PathToProduct", .PathToProduct
                        AddRow Entry, "ProductState", .ProductState
                    End With
                    AddRow SubElement, "Firewall"
                End If
            Next i
            
            For i = 0 To UBound(SW_SECURITY_PRODUCTS.Spyware)
                If SW_SECURITY_PRODUCTS.Spyware(i).ProductName <> vbNullString Then
                    AddRow SubElement, "Spyware", , False
                    With SW_SECURITY_PRODUCTS.Spyware(i)
                        AddRow Entry, "ProductName", .ProductName
                        AddRow Entry, "Publisher", .Publisher
                        AddRow Entry, "ProductVersion", .ProductVersion
                        AddRow Entry, "ProductGUID", .ProductGUID
                        AddRow Entry, "RTPStatus", Abs(.RTPStatus)
                        AddRow Entry, "UpToDate", Abs(.UpToDate)
                        AddRow Entry, "Enabled", Abs(.Enabled)
                        AddRow Entry, "PathToProduct", .PathToProduct
                        AddRow Entry, "ProductState", .ProductState
                    End With
                    AddRow SubElement, "Spyware"
                End If
            Next i
            
        AddRow Element, "SW_SECURITY_PRODUCTS"
        
        ' SW_LOCAL_USERS
        AddRow Element, "SW_LOCAL_USERS", , False
            For i = 0 To UBound(SW_LOCAL_USERS)
                With SW_LOCAL_USERS(i)
                    If .Name <> vbNullString Then
                        AddRow SubElement, "LocalUser", , False
                        
                        AddRow Entry, "Name", .Name
                        AddRow Entry, "FullName", .FullName
                        AddRow Entry, "Description", .Description
                        AddRow Entry, "UserComment", .UserComment
                        AddRow Entry, "CannotChangePassword", .CannotChangePassword = 1
                        AddRow Entry, "PasswordNeverExpires", .PasswordNeverExpires = 1
                        AddRow Entry, "PasswordExpired", .PasswordExpired = 1
                        AddRow Entry, "AccountDisabled", .AccountDisabled = 1
                        AddRow Entry, "AccountLocked", .AccountLocked = 1
                        AddRow Entry, "LastLogin", .LastLogin
                        AddRow Entry, "Group", .Group
                        AddRow Entry, "Groups", .Groups
                        AddRow Entry, "MaxPasswordLen", .MaxPasswordLen
                        AddRow Entry, "Password", .Password
                        
                        AddRow SubElement, "LocalUser"
                    End If
                End With
                
            Next i
        AddRow Element, "SW_LOCAL_USERS"
        
        ' HW_MOTHERBOARD
        AddRow Element, "HW_MOTHERBOARD", , False
            With HW_MOTHERBOARD
                If .Model <> vbNullString Or .Manufacturer <> vbNullString Or _
                   .BIOS <> vbNullString Or .SystemModel <> vbNullString Then
                    AddRow SubElement, "SystemModel", .SystemModel
                    AddRow SubElement, "SystemMfg", .SystemMfg
                    AddRow SubElement, "Model", .Model
                    AddRow SubElement, "Manufacturer", .Manufacturer
                    AddRow SubElement, "SerialNumber", .SerialNumber
                    AddRow SubElement, "MemorySlots", .MemorySlots
                    AddRow SubElement, "MemoryMax", .MemoryMax
                    AddRow SubElement, "CPUSocket", .CPUSocket
                    AddRow SubElement, "ChassisType", FormFactor(.ChassisType)
                    AddRow SubElement, "ChassisSN", .ChassisSN
                    AddRow SubElement, "BIOS", .BIOS
                    
                    AddRow SubElement, "FloppyDrives", , False
                    For i = 0 To UBound(.Floppy)
                        AddRow Entry, "Floppy", .Floppy(i)
                    Next i
                    AddRow SubElement, "FloppyDrives"
                    
                    AddRow SubElement, "Ports", , False
                    For i = 0 To UBound(.Ports)
                        AddRow Entry, "Port", .Ports(i)
                    Next i
                    AddRow SubElement, "Ports"
                End If
            End With
        AddRow Element, "HW_MOTHERBOARD"
        
       ' HW_HARDDISKS
        AddRow Element, "HW_HARDDISKS", , False
            For i = 0 To UBound(HW_HARDDISKS)
                With HW_HARDDISKS(i)
                    If .Model <> vbNullString And .Removable = False Then
                        AddRow SubElement, "HardDisk", , False
                        
                        AddRow Entry, "SerialNumber", .SerialNumber
                        AddRow Entry, "Model", .Model
                        AddRow Entry, "InterfaceType", .InterfaceType
                        AddRow Entry, "Family", .Family
                        'AddRow Entry, "RealSize", .RealSize
                        AddRow Entry, "Size", .Size
                        AddRow Entry, "Mode", .Mode
                        AddRow Entry, "PNPDeviceId", .PNPDeviceId
                        
                        AddRow Entry, "Partitions", , False
                        For j = 0 To UBound(.Paritions)
                            With .Paritions(j)
                                AddRow Entry1, "Partition", , False
                                    AddRow Entry2, "Caption", .Caption
                                    AddRow Entry2, "FileSystem", .FileSystem
                                    AddRow Entry2, "VolumeName", .VolumeName
                                    AddRow Entry2, "Size", .Size
                                AddRow Entry, "Partition"
                            End With
                        Next j
                        AddRow Entry, "Partitions"
                        
                        AddRow Entry, "RegisterHDD", , False
                            With .Registry
                                AddRow Entry1, "RegistryNo", .RegistryNo
                                AddRow Entry1, "InventaryNo", .InventaryNo
                                AddRow Entry1, "InventaryDate", .InventaryDate
                                AddRow Entry1, "InventarySerialNum", .InventarySerialNum
                                AddRow Entry1, "AdminName", .AdminName
                                AddRow Entry1, "AdminDate", .AdminDate
                                AddRow Entry1, "AdminSticker", .AdminSticker
                                AddRow Entry1, "RegistryName", .RegistryName
                                AddRow Entry1, "RegistryDate", .RegistryDate
                                AddRow Entry1, "RegistrySticker", .RegistrySticker
                            End With
                        AddRow Entry, "RegisterHDD"
                        
                        AddRow SubElement, "HardDisk"
                    End If
                End With
                
            Next i
        AddRow Element, "HW_HARDDISKS"
        
        ' HW_MONITORS
        AddRow Element, "HW_MONITORS", , False
            For i = 0 To UBound(HW_MONITORS)
                With HW_MONITORS(i)
                    If .Model <> vbNullString Or .Manufacturer <> vbNullString Or .ModelId <> vbNullString Then
                        AddRow SubElement, "Monitor", , False
                        
                        AddRow Entry, "Manufacturer", .Manufacturer
                        AddRow Entry, "Model", .Model
                        AddRow Entry, "PNPDeviceId", .PNPDeviceId
                        AddRow Entry, "Size", .Size
                        AddRow Entry, "HSize", .HSize
                        AddRow Entry, "VSize", .VSize
                        AddRow Entry, "SerialNumber", .SerialNumber
                        AddRow Entry, "ModelId", .ModelId
                        AddRow Entry, "ManufacureDate", .ManufacureDate
                        AddRow Entry, "VideoInput", .VideoInput
                        
                        AddRow SubElement, "Monitor"
                    End If
                End With
                
            Next i
        AddRow Element, "HW_MONITORS"
        
        ' HW_CDROMS
        AddRow Element, "HW_CDROMS", , False
            For i = 0 To UBound(HW_CDROMS)
                With HW_CDROMS(i)
                    If (.Model <> vbNullString Or .Manufacturer <> vbNullString Or .PNPDeviceId <> vbNullString) And _
                        .Virtual = False Then
                        AddRow SubElement, "CD-ROM", , False
                        
                        AddRow Entry, "Model", .Model
                        AddRow Entry, "Manufacturer", .Manufacturer
                        AddRow Entry, "ReadMedia", .ReadMedia
                        AddRow Entry, "WriteMedia", .WriteMedia
                        'AddRow Entry, "SerialNumber", .SerialNumber
                        AddRow Entry, "Description", .Description
                        AddRow Entry, "FirmWare", .FirmWare
                        AddRow Entry, "PNPDeviceId", .PNPDeviceId
                        
                        AddRow SubElement, "CD-ROM"
                    End If
                End With
                
            Next i
        AddRow Element, "HW_CDROMS"
        
        ' HW_CPU
        AddRow Element, "HW_CPU", , False
            With HW_CPU
                If .Model <> vbNullString Then
                    AddRow Entry, "Manufacturer", .Manufacturer
                    AddRow Entry, "Model", .Model
                    AddRow Entry, "Architecture", .Architecture
                    AddRow Entry, "ClockSpeed", .ClockSpeed
                    AddRow Entry, "Cores", .Cores
                    AddRow Entry, "LogicalProcessors", .LogicalProcessors
                    AddRow Entry, "Socket", CPU_Socket(.Socket)
                    AddRow Entry, "PNPDeviceId", .PNPDeviceId
                End If
            End With
        AddRow Element, "HW_CPU"
        
        ' HW_RAM_MEMORY
        AddRow Element, "HW_RAM_MEMORY", , False
            With HW_RAM_MEMORY
                If .TotalMemory > 0 Then
                    AddRow SubElement, "TotalMemory", .TotalMemory
                    AddRow SubElement, "RAMBanks", , False
                    
                    For i = 0 To UBound(.Banks)
                        With .Banks(i)
                            If .Capacity > 0 Or .BankLabel <> vbNullString Then
                                AddRow SubElement, "RAMBank", , False
                                
                                AddRow Entry, "BankLabel", .BankLabel
                                AddRow Entry, "Capacity", .Capacity
                                AddRow Entry, "Location", .Location
                                AddRow Entry, "Speed", .Speed
                                AddRow Entry, "Type", RAMType(.Type)
                                AddRow Entry, "FormFactor", RAMFormFactor(.FormFactor)
                                
                                AddRow SubElement, "RAMBank"
                            End If
                        End With
                    Next i
                    AddRow SubElement, "RAMBanks"
                End If
                

            End With
        AddRow Element, "HW_RAM_MEMORY"
        
        ' HW_NETWORK_ADAPTERS
        AddRow Element, "HW_NETWORK_ADAPTERS", , False
            For i = 0 To UBound(HW_NETWORK_ADAPTERS)
                With HW_NETWORK_ADAPTERS(i)
                    If .MACAddress <> vbNullString Then
                        AddRow SubElement, "NetworkAdapter", , False
                        
                        AddRow Entry, "AdapterType", AdapterTypeToStr(.AdapterType)
                        AddRow Entry, "Model", .Model
                        AddRow Entry, "Manufacturer", .Manufacturer
                        AddRow Entry, "MACAddress", .MACAddress
                        AddRow Entry, "PNPDeviceId", .PNPDeviceId
                        AddRow Entry, "Speed", .Speed
                        'AddRow Entry, "Configuration", .Configuration
                        
                        AddRow Entry, "Configuration", , False
                        With .Configuration
                            AddRow Entry1, "IP", Join(.IP, ",")
                            AddRow Entry1, "Mask", Join(.mask, ",")
                            AddRow Entry1, "GateWay", Join(.GateWay, ",")
                            AddRow Entry1, "DNS", Join(.DNS, ",")
                            AddRow Entry1, "DHCPEnabled", .DHCPEnabled
                            AddRow Entry1, "DHCPServer", .DHCPServer
                            AddRow Entry1, "NetConnectionID", .NetConnectionID
                        End With
                        AddRow Entry, "Configuration"
                        
                        AddRow SubElement, "NetworkAdapter"
                    End If
                End With
                
            Next i
        AddRow Element, "HW_NETWORK_ADAPTERS"
        
        ' HW_SOUND_DEVICES
        AddRow Element, "HW_SOUND_DEVICES", , False
            For i = 0 To UBound(HW_SOUND_DEVICES)
                With HW_SOUND_DEVICES(i)
                    If (.Model <> vbNullString Or .PNPDeviceId <> vbNullString) Then
                        AddRow SubElement, "SoundDevice", , False
                        
                        AddRow Entry, "Model", .Model
                        AddRow Entry, "Manufacturer", .Manufacturer
                        AddRow Entry, "PNPDeviceId", .PNPDeviceId
                        
                        AddRow SubElement, "SoundDevice"
                    End If
                End With
            Next i
        AddRow Element, "HW_SOUND_DEVICES"
        
        ' HW_VIDEO_ADAPTERS
        AddRow Element, "HW_VIDEO_ADAPTERS", , False
            For i = 0 To UBound(HW_VIDEO_ADAPTERS)
                With HW_VIDEO_ADAPTERS(i)
                    If (.Model <> vbNullString Or .PNPDeviceId <> vbNullString) Then
                        AddRow SubElement, "VideoAdapter", , False
                        
                        AddRow Entry, "Model", .Model
                        AddRow Entry, "Manufacturer", .Manufacturer
                        AddRow Entry, "PNPDeviceId", .PNPDeviceId
                        AddRow Entry, "VideoRAM", .VideoRAM
                        
                        AddRow SubElement, "VideoAdapter"
                    End If
                End With
            Next i
        AddRow Element, "HW_VIDEO_ADAPTERS"
        
        ' HW_HID
        AddRow Element, "HW_HID", , False
            With HW_HID
                AddRow SubElement, "Mouse", .Mouse
                AddRow SubElement, "Keyboard", .Keyboard
            End With
        AddRow Element, "HW_HID"
        
        ' HW_PRINTERS
        AddRow Element, "HW_PRINTERS", , False
            For i = 0 To UBound(HW_PRINTERS)
                With HW_PRINTERS(i)
                    If (.Model <> vbNullString Or .PortName <> vbNullString) Then
                        AddRow SubElement, "Printer", , False
                        
                        AddRow Entry, "Model", .Model
                        AddRow Entry, "Manufacturer", .Manufacturer
                        AddRow Entry, "DeviceID", .DeviceID
                        AddRow Entry, "ErrorState", .ErrorState
                        AddRow Entry, "DriverName", .DriverName
                        AddRow Entry, "PortName", .PortName
                        AddRow Entry, "IPAddress", .IPAddress
                        AddRow Entry, "hostname", .hostname
                        AddRow Entry, "ShareName", .ShareName
                        AddRow Entry, "ConnectionStat", .ConnectionStat
                        AddRow Entry, "IsOnline", .IsOnline
                        AddRow Entry, "IsDefault", .IsDefault
                        AddRow Entry, "IsLocal", .IsLocal
                        AddRow Entry, "IsNetwork", .IsNetwork
                        AddRow Entry, "IsShared", .IsShared
                        AddRow Entry, "InventaryNo", .InventaryNo
                        
                        AddRow SubElement, "Printer"
                    End If
                End With
            Next i
        AddRow Element, "HW_PRINTERS"
        
    AddRow Root, COMPUTER_NAME
    
End Sub


