Attribute VB_Name = "modMessages"
Option Explicit

Public Const MSG_VALIDATE_FIELDS = "Enter value"

Public Function DeviceStatusMessage(Value As Long) As String

    Select Case Value
        Case 0
            DeviceStatusMessage = "This device is working properly."
        Case 2
            DeviceStatusMessage = "Windows could not load the driver for this device because the computer is reporting two <bustype> bus types. (Code 2)"
        Case 3
            DeviceStatusMessage = "The driver for this device might be corrupted, or your system may be running low on memory or other resources. (Code 3)"
        Case 4
            DeviceStatusMessage = "This device is not working properly because one of its drivers may be bad, or your registry may be bad. (Code 4)"
        Case 5
            DeviceStatusMessage = "The driver for this device requested a resource that Windows does not know how to handle. (Code 5)"
        Case 6
            DeviceStatusMessage = "Another device is using the resources this device needs. (Code 6)"
        Case 7
            DeviceStatusMessage = "The drivers for this device need to be reinstalled. (Code 7)"
        Case 8
            DeviceStatusMessage = "Device failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation. (Code 8)"
        Case 9
            DeviceStatusMessage = "Windows cannot identify this hardware because it does not have a valid hardware identification number. (Code 9)"
        Case 10
            DeviceStatusMessage = "This device cannot start. (Code 10)"
        Case 11
            DeviceStatusMessage = "Windows stopped responding while attempting to start this device, and therefore will never attempt to start this device again. (Code 11)"
        Case 12
            DeviceStatusMessage = "This device cannot find enough free resources that it can use. (Code 12)"
        Case 13
            DeviceStatusMessage = "This device is either not present, not working properly, or does not have all the drivers installed. (Code 13)"
        Case 14
            DeviceStatusMessage = "This device cannot work properly until you restart your computer. (Code 14)"
        Case 15
            DeviceStatusMessage = "This device is causing a resource conflict. (Code 15)"
        Case 16
            DeviceStatusMessage = "Windows cannot identify all the resources this device uses. (Code 16)"
        Case 17
            DeviceStatusMessage = "The driver information file (INF-file-name) is telling this child device to use a resource that the parent device does not have or recognize. (Code 17)"
        Case 18
            DeviceStatusMessage = "Reinstall the drivers for this device. (Code 18)"
        Case 19
            DeviceStatusMessage = "Windows cannot start this hardware device because its configuration information (in the registry) is incomplete or damaged. (Code 19)"
        Case 20
            DeviceStatusMessage = "Windows could not load one of the drivers for this device. (Code 20)"
        Case 21
            DeviceStatusMessage = "Windows is removing this device. (Code 21)"
        Case 22
            DeviceStatusMessage = "This device is disabled. (Code 22)"
        Case 23
            DeviceStatusMessage = "This display adapter is functioning correctly. (Code 23)"
        Case 24
            DeviceStatusMessage = "This device is not present, is not working properly, or does not have all its drivers installed. (Code 24)"
        Case 25
            DeviceStatusMessage = "Windows is in the process of setting up this device. (Code 25)"
        Case 26
            DeviceStatusMessage = "Windows is in the process of setting up this device. (Code 26)"
        Case 27
            DeviceStatusMessage = "Windows can't specify the resources for this device. (Code 27)"
        Case 28
            DeviceStatusMessage = "The drivers for this device are not installed. (Code 28)"
        Case 29
            DeviceStatusMessage = "This device is disabled because the firmware of the device did not give it the required resources. (Code 29)"
        Case 30
            DeviceStatusMessage = "This device is using an Interrupt Request (IRQ) resource that is in use by another device and cannot be shared. You must change the conflicting setting or remove the real-mode driver causing the conflict. (Code 30)"
        Case 31
            DeviceStatusMessage = "This device is not working properly because Windows cannot load the drivers required for this device. (Code 31)"
        Case 32
            DeviceStatusMessage = "A driver (service) for this device has been disabled. An alternate driver may be providing this functionality. (Code 32)"
        Case 33
            DeviceStatusMessage = "Windows cannot determine which resources are required for this device. (Code 33)"
        Case 34
            DeviceStatusMessage = "Windows cannot determine the settings for this device. Consult the documentation that came with this device and use the Resource tab to set the configuration. (Code 34)"
        Case 35
            DeviceStatusMessage = "Your computer's system firmware does not include enough information to properly configure and use this device. To use this device, contact your computer manufacturer to obtain a firmware or BIOS update. (Code 35)"
        Case 36
            DeviceStatusMessage = "This device is requesting a PCI interrupt but is configured for an ISA interrupt (or vice versa). Please use the computer's system setup program to reconfigure the interrupt for this device. (Code 36)"
        Case 37
            DeviceStatusMessage = "Windows cannot initialize the device driver for this hardware. (Code 37)"
        Case 38
            DeviceStatusMessage = "Windows cannot load the device driver for this hardware because a previous instance of the device driver is still in memory. (Code 38)"
        Case 39
            DeviceStatusMessage = "Windows cannot load the device driver for this hardware. The driver may be corrupted or missing. (Code 39)"
        Case 40
            DeviceStatusMessage = "Windows cannot access this hardware because its service key information in the registry is missing or recorded incorrectly (Code 40)"
        Case 41
            DeviceStatusMessage = "Windows successfully loaded the device driver for this hardware but cannot find the hardware device. (Code 41)"
        Case 42
            DeviceStatusMessage = "Windows cannot load the device driver for this hardware becuse there is a duplicate device already running in the system. (Code 42)"
        Case 43
            DeviceStatusMessage = "Windows has stopped this device because it has reported problems. (Code 43)"
        Case 44
            DeviceStatusMessage = "An application or service has shut down this hardware device. (Code 44)"
        Case 45
            DeviceStatusMessage = "Currently, this hardware device is not connected to the computer. (Code 45)"
        Case 46
            DeviceStatusMessage = "Windows cannot gain access to this hardware device because the operating system is in the process of shutting down. (Code 46)"
        Case 47
            DeviceStatusMessage = "Windows cannot use this hardware device because it has been prepared for 'safe removal', but it has not been removed from the computer. (Code 47)"
        Case 48
            DeviceStatusMessage = "The software for this device has been blocked from starting because it is known to have problems with Windows. Contact the hardware vendor for a new driver. (Code 48)"
        Case 49
            DeviceStatusMessage = "Windows cannot start new hardware devices because the system hive is too large (exceeds the Registry Size Limit). (Code 49)"
        Case 50
            DeviceStatusMessage = "Windows cannot apply all of the properties for this device. Device properties may include information that describes the device's capabilities and settings (such as security settings for example). (Code 50)"
        Case Else
            DeviceStatusMessage = "Unknown device status. ( Code " & Value & ")"
    End Select

End Function

Public Function NetErrorMsg(iMsg As Integer) As String
    Select Case iMsg
        Case 0: NetErrorMsg = "Successful completion, no reboot required."
        Case 1: NetErrorMsg = "Successful completion, reboot required."
        Case 64: NetErrorMsg = "Method not supported on this platform."
        Case 65: NetErrorMsg = "Unknown failure."
        Case 66: NetErrorMsg = "Invalid subnet mask."
        Case 67: NetErrorMsg = "An error occurred while processing an instance that was returned."
        Case 68: NetErrorMsg = "Invalid input parameter."
        Case 69: NetErrorMsg = "More than five gateways specified."
        Case 70: NetErrorMsg = "Invalid IP address."
        Case 71: NetErrorMsg = "Invalid gateway IP address."
        Case 72: NetErrorMsg = "An error occurred while accessing the registry for the requested information."
        Case 73: NetErrorMsg = "Invalid domain name."
        Case 74: NetErrorMsg = "Invalid host name."
        Case 75: NetErrorMsg = "No primary or secondary WINS server defined."
        Case 76: NetErrorMsg = "Invalid file."
        Case 77: NetErrorMsg = "Invalid system path."
        Case 78: NetErrorMsg = "File copy failed."
        Case 79: NetErrorMsg = "Invalid security parameter."
        Case 80: NetErrorMsg = "Unable to configure TCP/IP service."
        Case 81: NetErrorMsg = "Unable to configure DHCP service."
        Case 82: NetErrorMsg = "Unable to renew DHCP lease."
        Case 83: NetErrorMsg = "Unable to release DHCP lease."
        Case 84: NetErrorMsg = "IP not enabled on adapter."
        Case 85: NetErrorMsg = "IPX not enabled on adapter."
        Case 86: NetErrorMsg = "Frame or network number bounds error."
        Case 87: NetErrorMsg = "Invalid frame type."
        Case 88: NetErrorMsg = "Invalid network number."
        Case 89: NetErrorMsg = "Duplicate network number."
        Case 90: NetErrorMsg = "Parameter out of bounds."
        Case 91: NetErrorMsg = "Access denied."
        Case 92: NetErrorMsg = "Out of memory."
        Case 93: NetErrorMsg = "Already exists."
        Case 94: NetErrorMsg = "Path, file, or object not found."
        Case 95: NetErrorMsg = "Unable to notify service."
        Case 96: NetErrorMsg = "Unable to notify DNS service."
        Case 97: NetErrorMsg = "Interface not configurable."
        Case 98: NetErrorMsg = "Not all DHCP leases could be released or renewed."
        Case 100: NetErrorMsg = "DHCP not enabled on the adapter."
    End Select
End Function

Public Function WMINetConnectorStatus(status) As String

    Select Case status
        Case 0: WMINetConnectorStatus = "Disconnected"
        Case 1: WMINetConnectorStatus = "Connecting"
        Case 2: WMINetConnectorStatus = "Connected"
        Case 3: WMINetConnectorStatus = "Disconnecting"
        Case 4: WMINetConnectorStatus = "Hardware not present"
        Case 5: WMINetConnectorStatus = "Hardware disabled"
        Case 6: WMINetConnectorStatus = "Hardware malfunction"
        Case 7: WMINetConnectorStatus = "Media disconnected"
        Case 8: WMINetConnectorStatus = "Authenticating"
        Case 9: WMINetConnectorStatus = "Authentication succeeded (Connected)"
        Case 10: WMINetConnectorStatus = "Authentication failed"
        Case 11: WMINetConnectorStatus = "Invalid address"
        Case 12: WMINetConnectorStatus = "Credentials required"
    End Select

End Function

Public Function GetWindowsActivationStatus(iStatus As Integer) As String
    
    Dim Result As String
    
    Select Case iStatus
        Case 0, 2 To 4, 6
            Result = "Windows IS NOT activated"
        Case 1
            Result = "Windows is activated"
        Case 5
            Result = "Windows IS NOT activated (notifications)"
    End Select
    
    GetWindowsActivationStatus = Result
    
End Function
