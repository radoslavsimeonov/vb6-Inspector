Attribute VB_Name = "hwMotherBoard"
Option Explicit

Public Function GetMotherBoard() As HardwareMotherBoard
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim mobo     As HardwareMotherBoard

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

    With mobo
        
        Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_BaseBoard", , 48)

        For Each objItem In colItems
            .Manufacturer = GetProp(objItem.Manufacturer)
            .Model = GetProp(objItem.Product)
            .SerialNumber = GetProp(objItem.SerialNumber)
        Next

        Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem", , 48)

        For Each objItem In colItems
            If LCase$(.Manufacturer) <> LCase$(GetProp(objItem.Manufacturer)) Then _
                .SystemMfg = GetProp(objItem.Manufacturer)
            If LCase$(.Model) <> LCase$(GetProp(objItem.Model)) Then _
                .SystemModel = GetProp(objItem.Model)
        Next

        Set colItems = objWMIService.ExecQuery("SELECT MemoryDevices,MaxCapacity FROM Win32_PhysicalMemoryArray", , 48)

        For Each objItem In colItems
            .MemorySlots = GetProp(objItem.MemoryDevices)
            .MemoryMax = GetProp(objItem.MaxCapacity)
        Next

        Set colItems = objWMIService.ExecQuery("SELECT ChassisTypes,SerialNumber FROM Win32_SystemEnclosure", , 48)

        For Each objItem In colItems
            .ChassisType = Join(objItem.ChassisTypes, ",")
            .ChassisSN = GetProp(objItem.SerialNumber)
        Next

        Set colItems = objWMIService.ExecQuery("SELECT Manufacturer  FROM Win32_BIOS", , 48)

        For Each objItem In colItems
            .BIOS = GetProp(objItem.Manufacturer)
        Next

        .Ports = GetDevicesList("{4d36e978-e325-11ce-bfc1-08002be10318}")
        .Floppy = GetDevicesList("{4d36e980-e325-11ce-bfc1-08002be10318}")
        '.PCIDevices = GetOnboardDevices
        '.USB = GetDevicesList("{36fc9e60-c465-11cf-8056-444553540000}")
    End With

    GetMotherBoard = mobo
    Set objWMIService = Nothing
    Set colItems = Nothing
End Function

Public Function FormFactor(system_form_factor) As String

    Select Case system_form_factor
        Case "1"
            FormFactor = "Other"
        Case "2"
            FormFactor = "Unknown"
        Case "3"
            FormFactor = "Desktop"
        Case "4"
            FormFactor = "Low Profile Desktop"
        Case "5"
            FormFactor = "Pizza Box"
        Case "6"
            FormFactor = "Mini Tower"
        Case "7"
            FormFactor = "Tower"
        Case "8"
            FormFactor = "Portable"
        Case "9"
            FormFactor = "Laptop"
        Case "10"
            FormFactor = "Notebook"
        Case "11"
            FormFactor = "Hand Held"
        Case "12"
            FormFactor = "Docking Station"
        Case "13"
            FormFactor = "All in One"
        Case "14"
            FormFactor = "sub Notebook"
        Case "15"
            FormFactor = "Space-Saving"
        Case "16"
            FormFactor = "Lunch Box"
        Case "17"
            FormFactor = "Main System Chassis"
        Case "18"
            FormFactor = "Expansion Chassis"
        Case "19"
            FormFactor = "SubChassis"
        Case "20"
            FormFactor = "Bus Expansion Chassis"
        Case "21"
            FormFactor = "Peripheral Chassis"
        Case "22"
            FormFactor = "Storage Chassis"
        Case "23"
            FormFactor = "Rack Mount Chassis"
        Case "24"
            FormFactor = "Sealed-case PC"
    End Select

End Function

Public Function GetOnboardDevices() As HardwareDevice()
On Error Resume Next

    Dim objBuses, objBus, objDevices, objDevice
    Dim dev()  As HardwareDevice
    Dim iDev   As Integer
    Dim strMsg As String

    iDev = 0
    ReDim dev(iDev)
    Set objBuses = GetObject("winmgmts:").InstancesOf("Win32_Bus")

    For Each objBus In objBuses

        If (objBus.BusType = 5) Then
            Set objDevices = GetObject("winmgmts:").ExecQuery("Associators of {Win32_Bus.DeviceID=""" & objBus.DeviceID & """} WHERE AssocClass = Win32_DeviceBus")

            For Each objDevice In objDevices

                If GetProp(objDevice.ClassGuid) <> "" Then
                    If LCase$(objDevice.ClassGuid) <> "{4d36e97d-e325-11ce-bfc1-08002be10318}" Then ' Don't show system devices
                        ReDim Preserve dev(iDev)

                        With dev(iDev)

                            Dim tmpDev() As HardwareDevice

                            If GetDevProp(objDevice.DeviceID, tmpDev) Then dev(iDev) = tmpDev(0)
                            .BusName = GetProp(objBus.DeviceID)
                            .DeviceStatus = GetProp(objDevice.status)
                            .ErrorCode = GetProp(objDevice.ConfigManagerErrorCode)

                            If Len(dev(iDev).ClassGuid) = 0 Then
                                .HardwareIDs = GetProp(objDevice.PNPDeviceId)
                                .DeviceDesc = GetProp(objDevice.Description)
                            End If

                        End With

                        iDev = iDev + 1
                    End If
                End If

            Next

        End If

    Next

    GetOnboardDevices = dev
    Set objBuses = Nothing
    Set objDevices = Nothing
End Function

Private Function GetDevicesList(strGUID As String) As String()

    Dim tmpDev() As HardwareDevice
    Dim tmpStr() As String
    Dim i        As Integer

    ReDim tmpStr(0)

    If GetDevProp(strGUID, tmpDev) Then
        ReDim tmpStr(UBound(tmpDev))

        For i = 0 To UBound(tmpDev)

            If Trim$(Len(tmpDev(i).FriendlyName)) > 0 Then
                tmpStr(i) = tmpDev(i).FriendlyName
            Else
                tmpStr(i) = tmpDev(i).DeviceDesc
            End If

        Next

    End If

    GetDevicesList = tmpStr
End Function

Private Function GetUSBControllers() As String()

    Dim tmpDev() As HardwareDevice
    Dim tmpStr() As String
    Dim i        As Integer

    If GetDevProp("{9d7debbc-c85d-11d1-9eb4-006008c3a19a}", tmpDev) Then
        ReDim tmpStr(UBound(tmpDev))

        For i = 0 To UBound(tmpDev)
            tmpStr(i) = tmpDev(i).DeviceDesc
        Next

    End If

    GetUSBControllers = tmpStr
End Function

Public Function GetDeviceErrorCodeMsg(errCode As Integer) As String

    Dim Msg As String

    Select Case errCode
        Case 0
            Msg = "This device is working properly."
        Case 1
            Msg = "This device is not configured correctly."
        Case 2
            Msg = "Windows cannot load the driver for this device."
        Case 3
            Msg = "The driver might be corrupted, or your system " & "may be running low on memory or other resources."
        Case 4
            Msg = "This device is not working properly. One of its " & "drivers or your registry might be corrupted."
        Case 5
            Msg = "The driver for this device needs a resource " & "that Windows cannot manage."
        Case 6
            Msg = "The boot configuration for this device " & "conflicts with other devices."
        Case 7
            Msg = "Cannot filter."
        Case 8
            Msg = "The driver loader for the device is missing."
        Case 9
            Msg = "This device is not working properly because" & "the controlling firmware is reporting the " & "resources for the device incorrectly."
        Case 10
            Msg = "This device cannot start."
        Case 11
            Msg = "This device failed."
        Case 12
            Msg = "This device cannot find enough free " & "resources that it can use."
        Case 13
            Msg = "Windows cannot verify this device's resources."
        Case 14
            Msg = "This device cannot work properly until " & "you restart your computer."
        Case 15
            Msg = "This device is not working properly because " & "there is probably a re-enumeration problem."
        Case 16
            Msg = "Windows cannot identify all the resources this device uses."
        Case 17
            Msg = "This device is asking for an unknown resource type."
        Case 18
            Msg = "Reinstall the drivers for this device."
        Case 19
            Msg = "Failure using the VXD loader."
        Case 20
            Msg = "Your registry might be corrupted."
        Case 21
            Msg = "System failure: Try changing the driver for this device. " & "If that does not work, see your hardware " & "documentation. Windows is removing this device."
        Case 22
            Msg = "This device is disabled."
        Case 23
            Msg = "System failure: Try changing the driver for " & "this device. If that doesn't work, see your " & "hardware documentation."
        Case 24
            Msg = "This device is not present, is not working " & "properly, or does not have all its drivers installed."
        Case 25
            Msg = "Windows is still setting up this device."
        Case 26
            Msg = "Windows is still setting up this device."
        Case 27
            Msg = "This device does not have valid log configuration."
        Case 28
            Msg = "The drivers for this device are not installed."
        Case 29
            Msg = "This device is disabled because the firmware of " & "the device did not give it the required resources."
        Case 30
            Msg = "This device is using an Interrupt Request (IRQ) " & "resource that another device is using."
        Case 31
            Msg = "This device is not working properly because Windows " & "cannot load the drivers required for this device."
    End Select

    GetDeviceErrorCodeMsg = Msg
End Function
