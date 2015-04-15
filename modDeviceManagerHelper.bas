Attribute VB_Name = "modDeviceManagerHelper"
Option Explicit

Public Sub FillDeviceDetails(dV As HardwareDevice, lv As ListView)
    
    Dim i, ii           As Integer
    Dim arrDNS()        As String
    Dim arrPntCap()     As String
    Dim parentDev       As HardwareDevice
    
    lv.ListItems.Clear
    modFillListView.MY_ListView = lv
    
    Select Case UCase$(dV.Class)
        
        Case "CDROM"
            For i = 0 To UBound(HW_CDROMS)
                With HW_CDROMS(i)
                    If .PNPDeviceId = dV.PNPDeviceId Then
                        AddListItem .Model, "Model"
                        AddListItem .Manufacturer, "Manufacturer"
                        AddListItem .Description, "Description"
                        AddListItem .Interface, "Interface"
                        AddListItem .FirmWare, "Firmware"
                        AddListItem .ReadMedia, "Read media"
                        AddListItem .WriteMedia, "Write media"
                        AddListItem .SerialNumber, "Serial number"
                        AddListItem IIf(.Virtual, "Yes", ""), "Virtual drive"
                        AddListItem .DriveLetter, "Drive letter"
                    End If
                End With
            Next i
        Case "PROCESSOR"
            With HW_CPU
                AddListItem .Model, "Model"
                AddListItem .Manufacturer, "Manufacturer"
                AddListItem .Architecture & "bit", "Architecture"
                AddListItem .ClockSpeed & "MHz", "Speed"
                AddListItem .Cores, "Cores"
                AddListItem .LogicalProcessors, "Processors"
                AddListItem CPU_Socket(.Socket), "Socket"
            End With
        Case "DISKDRIVE"
            For i = 0 To UBound(HW_HARDDISKS)
                With HW_HARDDISKS(i)
                    If .PNPDeviceId = dV.PNPDeviceId Then
                        AddListItem .Model, "Model"
                        If (Trim$(.Model) <> Trim$(.Family)) And .Family <> "" Then _
                            AddListItem .Family, "Family"
                        AddListItem .Size & " GB", "Size"
                        AddListItem .InterfaceType, "Interface"
                        AddListItem .Mode, "Mode"
                        AddListItem .SerialNumber, "Serial number"
                        
                        AddListItem "", ""
                        
                        If UBound(.Paritions) > -1 Then
                            For ii = 0 To UBound(.Paritions)
                                AddListItem .Paritions(ii).Caption & _
                                     " (" & _
                                    .Paritions(ii).FileSystem & ") " & _
                                    FormatBytes(.Paritions(ii).Size) & " (" & _
                                    FormatBytes(.Paritions(ii).FreeSpace) & " free)", _
                                    IIf(ii > 0, "", "Partitions")
                                    
                            Next ii
                        End If
                    End If
                End With
            Next i
        Case "MONITOR"
            For i = 0 To UBound(HW_MONITORS)
                With HW_MONITORS(i)
                    AddListItem .Model & " " & .Size & "''", "Model"
                    AddListItem .Manufacturer, "Manufacturer"
                    AddListItem .SerialNumber, "Serial number"
                    AddListItem .AspectRatio, "Aspect ratio"
                    AddListItem .HSize & "cm x " & .VSize & "cm", "Dimensions"
                    AddListItem .VideoInput, "Video input"
                    AddListItem .ManufacureDate, "Manufature date"
                End With
            Next
    
        Case "COMPUTER" 'MOTHERBOARD 'MEMORY
            FillComputerDetails lv
        Case "NET"
            For i = 0 To UBound(HW_NETWORK_ADAPTERS)
                If HW_NETWORK_ADAPTERS(i).PNPDeviceId = dV.PNPDeviceId Then
                    With HW_NETWORK_ADAPTERS(i)
                        
                        AddListItem .Model, "Model"
                        AddListItem .Manufacturer, "Manufacturer"
                        AddListItem AdapterTypeToStr(.AdapterType), "Device type"
                        AddListItem .MACAddress, "MAC Address"
                        If val(.Speed) > 0 Then _
                            AddListItem .Speed & " Mbps", "Speed"
                        
                        If .Configuration.DHCPEnabled = True Then
                            AddListItem "", ""
                            AddListItem IIf(.Configuration.DHCPEnabled, "Yes", "No"), "DHCP enabled"
                            AddListItem .Configuration.DHCPServer, "DHCP Server"
                        End If
                        
                        AddListItem "", ""
                                                
                        If Len(Join(.Configuration.IP)) <> 0 Then AddListItem Join(.Configuration.IP, ", "), "IP address"
                        If Len(Join(.Configuration.mask)) <> 0 Then AddListItem Join(.Configuration.mask, ", "), "Subnet mask"
                        If Len(Join(.Configuration.GateWay)) <> 0 Then AddListItem Join(.Configuration.GateWay, ", "), "Default gateway"
                        If Len(Join(.Configuration.DNS)) <> 0 Then
                            For ii = 0 To UBound(.Configuration.DNS)
                                AddListItem .Configuration.DNS(ii), IIf(ii > 0, "", "DNS servers")
                            Next ii
                            AddListItem "", ""
                        End If
                        
                        'If .Configuration.ConnectionStatus <> "" Then
                        AddListItem WMINetConnectorStatus(.Configuration.ConnectionStatus), "Connection status"
                        AddListItem .Configuration.NetConnectionID, "Connection name"
                        
                    End With
                End If
            Next
        Case "PRINTER", "PRINTQUEUE" ' ????
            For i = 0 To UBound(HW_PRINTERS)
                With HW_PRINTERS(i)
                    If .DeviceID = dV.DeviceDesc Or _
                        .DeviceID = dV.FriendlyName Then
                            AddListItem .Model, "Model"
                            AddListItem .Manufacturer, "Manufacturer"
                            AddListItem IIf(.IsDefault, "Yes", "No"), _
                                        "Default"
                            AddListItem IIf(.IsOnline, "Yes", "No"), _
                                        "Active"
                            AddListItem IIf(.IsLocal, "Yes", "No"), _
                                        "Local"
                            AddListItem .PortName, "Port name"
                            AddListItem .DriverName, "Driver"
                            
                            If .IsShared Then
                                AddListItem IIf(.IsShared, "Yes", "No"), _
                                            "Shared"
                                AddListItem .ShareName, "Shared name"
                            End If
                                                       
                            If .IsNetwork Then
                                AddListItem "", ""
                                AddListItem IIf(.IsNetwork, "Yes", "No"), _
                                            "Network"
                                AddListItem .hostname, "Host name"
                                AddListItem .IPAddress, "IP address"
                                AddListItem .ConnectionStat, "Host connection"
                        End If
                    End If
                End With
            Next i
        Case "DISPLAY"
            For i = 0 To UBound(HW_VIDEO_ADAPTERS)
                With HW_VIDEO_ADAPTERS(i)
                    If .PNPDeviceId = dV.PNPDeviceId Then
                        AddListItem .Model, "Model"
                        AddListItem .Manufacturer, "Manufacturer"
                        If val(.VideoRAM) <> 0 Then _
                            AddListItem FormatBytes(.VideoRAM, 0), "Memory"
                    End If
                End With
            Next i
        Case "MOUSE", "KEYBOARD"
        
        Case "MEDIA"
        
        Case "USB" ' TODO:
        
        Case "HIDCLASS"
        
        Case "HDC" 'IDE ATA/ATAPI controllers
        
        Case "SCSIADAPTER" ' Storage controller
        
        Case "SYSTEM MANAGEMENT" ' TODO
    End Select
    
    If lv.ListItems.count > 0 Then AddListItem "", ""
    
    AddListItem dV.ClassDesc, "Device type"
    
    If Left$(dV.Manufacturer, 1) <> "(" And Right$(dV.Manufacturer, 1) <> ")" _
        And InStr(LCase$(dV.Manufacturer), "microsoft") = 0 And dV.Manufacturer <> "" Then _
        AddListItem dV.Manufacturer, "Manufacturer"
        
    If GetDeviceByDevInst(dV.DevParent, parentDev) Then _
        AddListItem IIf(parentDev.FriendlyName <> "", _
                    parentDev.FriendlyName, parentDev.DeviceDesc), _
                    "Parent device"
    
    AddListItem dV.DeviceStatus, "Status"
       
    If Len(dV.VenDev.VEN) > 0 Or Len(dV.VenDev.dev) > 0 Then
    AddListItem "", ""
    AddListItem "Chip details", "", True
    
    AddListItem _
                IIf(Len(dV.VenDev.VEN) > 0, dV.VenDev.VEN, "") & _
                IIf(Len(dV.VenDev.dev) > 0, " - " & dV.VenDev.dev, ""), _
                "Instance Id"
    End If
    
    If Len(dV.VenDev.Type) > 0 Then _
        AddListItem dV.VenDev.Type, "Enumerator"
    If Len(dV.VenDevInfo.Chip) > 0 Then _
        AddListItem dV.VenDevInfo.Chip, "Device Id"
    If Len(dV.VenDevInfo.Vendor) > 0 Then _
        AddListItem dV.VenDevInfo.Vendor, "Vendor Id"

    AutoSizeListViewColumns lv
End Sub

Public Sub FillComputerDetails(Optional lv As ListView)
On Error Resume Next

    Dim ii As Integer

    lv.ListItems.Clear
    
            With HW_MOTHERBOARD
            
                AddListItem .SystemModel, "System model"
                AddListItem .SystemMfg, "Manufacturer"
                
                If .SystemModel <> vbNullString Or .SystemMfg <> vbNullString Then _
                    AddListItem "", ""
                
                AddListItem .Manufacturer & " " & .Model, "Motherboard"
                AddListItem .SerialNumber, "Serial number"
                AddListItem .BIOS, "BIOS"
                
                AddListItem "", "", True
                AddListItem FormFactor(.ChassisType), "Chassis type"
                AddListItem .ChassisSN, "Chassis S/N"
                
                AddListItem "", "", True
                AddListItem HW_CPU.Model & ", " & HW_CPU.Architecture & "bit", "Processor"
                
                AddListItem "", "", True
                AddListItem FormatBytes(HW_RAM_MEMORY.TotalMemory, 0), "Installed memory"
                For ii = 0 To UBound(HW_RAM_MEMORY.Banks)
                    AddListItem HW_RAM_MEMORY.Banks(ii).BankLabel & _
                                " " & FormatBytes(HW_RAM_MEMORY.Banks(ii).Capacity, 0) & _
                                "  " & RAMType(HW_RAM_MEMORY.Banks(ii).Type) & "-" & _
                                HW_RAM_MEMORY.Banks(ii).Speed & " " & _
                                RAMFormFactor(HW_RAM_MEMORY.Banks(ii).FormFactor), _
                                ""
                Next ii
                AddListItem FormatBytes(.MemoryMax * 1024, 0), "Max memory"
                AddListItem .MemorySlots, "Memory slots"
                
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
    AutoSizeListViewColumns lv
End Sub


'Private Sub AddListItem(sField As String, sValue As String, Optional bBold As Boolean = False)
'
'    Dim itmX As ListItem
'
'    Set itmX = lstV.ListItems.Add(, , sField)
'    itmX.ForeColor = &H80FF&        '&HC0C0C0
'    itmX.Bold = bBold
'    itmX.SubItems(1) = sValue
'    itmX.ListSubItems(1).Bold = bBold
'    itmX.ListSubItems(1).ForeColor = vbWhite
'End Sub
