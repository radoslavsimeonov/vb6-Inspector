Attribute VB_Name = "hwNetwork"
Option Explicit

Private Const NET_ADAPTER_REG As String = "SYSTEM\CurrentControlSet\Services\tcpip\Parameters\Interfaces\"

Private gGUID As String

Public Function EnumNetworkAdapters() As HardwareNetworkAdapter()
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim colNicConfigs, objItemConfig
    Dim tmpArr() As HardwareNetworkAdapter
    Dim idx      As Integer
    Dim tGUID    As String
    Dim tDNS     As String

    idx = 0
    ReDim tmpArr(idx)
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = _
        objWMIService.ExecQuery( _
        "SELECT * FROM Win32_NetworkAdapter WHERE NOT Manufacturer='Microsoft' " & _
        "AND NOT ProductName LIKE '%mini%po%rt%' AND NOT MACAddress='' AND " & _
        "MACAddress IS NOT NULL AND NOT Name LIKE '%virtual%'", , 48)

    For Each objItem In colItems
        ReDim Preserve tmpArr(idx)
        
        ' Get Network adapter properties
        With tmpArr(idx)
            .AdapterType = GetProp(objItem.AdapterTypeID)
            .MACAddress = GetProp(objItem.MACAddress)
            .Manufacturer = GetProp(objItem.Manufacturer)
            .Model = GetProp(objItem.Name)
            .PNPDeviceId = GetProp(objItem.PNPDeviceId)
            .Speed = GetProp(objItem.Speed)

            If val(.Speed) > 0 Then
                .Speed = Int(.Speed) / 1000000
                Else:
                .Speed = 0
            End If
            Set colNicConfigs = _
                objWMIService.ExecQuery("ASSOCIATORS OF {Win32_NetworkAdapter.DeviceID='" & _
                objItem.DeviceID & "'} WHERE AssocClass=Win32_NetworkAdapterSetting")

            For Each objItemConfig In colNicConfigs
                .GUID = GetProp(objItemConfig.SettingID)
            Next

            tGUID = .GUID
        End With
        
        ' Get Network adapter configuration
        With tmpArr(idx).Configuration
            .ConnectionStatus = GetProp(objItem.NetConnectionStatus)
            .NetConnectionID = GetProp(objItem.NetConnectionID)
            .DHCPEnabled = (QueryValue(HKEY_LOCAL_MACHINE, NET_ADAPTER_REG & tGUID, "EnableDHCP") = 1)

            If .DHCPEnabled Then
                ReDim .IP(0)
                ReDim .mask(0)
                
                .DHCPServer = QueryValue(HKEY_LOCAL_MACHINE, _
                                         NET_ADAPTER_REG & tGUID, _
                                         "DhcpServer")
                
                .IP(0) = QueryValue(HKEY_LOCAL_MACHINE, _
                                    NET_ADAPTER_REG & tGUID, _
                                    "DhcpIPAddress")

                If .IP(0) = "0.0.0.0" Then .IP(0) = ""
                
                .mask(0) = QueryValue(HKEY_LOCAL_MACHINE, _
                                      NET_ADAPTER_REG & tGUID, _
                                      "DhcpSubnetMask")

                If .mask(0) = "255.0.0.0" Then .mask(0) = ""

                On Error Resume Next

                .GateWay = QueryValue(HKEY_LOCAL_MACHINE, _
                                      NET_ADAPTER_REG & tGUID, _
                                      "DhcpDefaultGateway")

                On Error GoTo 0

                tDNS = QueryValue(HKEY_LOCAL_MACHINE, _
                                  NET_ADAPTER_REG & tGUID, _
                                  "DhcpNameServer")
                
                .DNS = Split(tDNS, " ")
            Else
                On Local Error Resume Next
                .IP = QueryValue(HKEY_LOCAL_MACHINE, NET_ADAPTER_REG & tGUID, "IPAddress")
                .mask = QueryValue(HKEY_LOCAL_MACHINE, NET_ADAPTER_REG & tGUID, "SubnetMask")
                .GateWay = QueryValue(HKEY_LOCAL_MACHINE, NET_ADAPTER_REG & tGUID, "DefaultGateway")
                tDNS = QueryValue(HKEY_LOCAL_MACHINE, NET_ADAPTER_REG & tGUID, "NameServer")
                .DNS = Split(tDNS, ",")
                On Error GoTo 0
            End If

        End With

        idx = idx + 1
    Next

    Set objWMIService = Nothing
    Set colItems = Nothing
    
    EnumNetworkAdapters = tmpArr
    
End Function

Public Function SetNetConfig(ByVal tGUID As String, _
                            ByVal tPNP As String, _
                            Optional sIP As String = vbNullString, _
                            Optional sMASK As String = vbNullString, _
                            Optional sGATEWAY As String = vbNullString, _
                            Optional sDNS1 As String = vbNullString, _
                            Optional sDNS2 As String = vbNullString _
                             ) As Boolean
                        
    Dim objWMIService
    Dim colNetAdapters
    Dim objNetAdapter
    Dim colNicConfigs
    Dim objItemConfig
    
    Dim bOK             As Boolean
      
    Dim tDNS            As String
    Dim tDev            As HardwareDevice
    
    If tGUID = vbNullString Or tPNP = vbNullString Then
        SetNetConfig = False
        Exit Function
    End If
    
    gGUID = tGUID
    
    ' Write the network car settings in the regisrty
    tDNS = sDNS1 & IIf(sDNS2 <> vbNullString, "," & sDNS2, "")

    bOK = SetRegIPAddress("IPAddress", _
                          sIP & vbNullChar, _
                          REG_MULTI_SZ)
                          
    bOK = SetRegIPAddress("SubnetMask", _
                          sMASK & vbNullChar, _
                          REG_MULTI_SZ)
                          
    bOK = SetRegIPAddress("DefaultGateway", _
                          sGATEWAY & vbNullChar, _
                          REG_MULTI_SZ)
                          
    bOK = SetRegIPAddress("DefaultGatewayMetric", _
                          IIf(sIP = vbNullString, vbNullString, 0), _
                          REG_MULTI_SZ)
                          
    bOK = SetRegIPAddress("NameServer", _
                          tDNS & vbNullChar, _
                          REG_SZ)
                          
    bOK = SetRegIPAddress("EnableDHCP", _
                          IIf(sIP = vbNullString, 1, 0), _
                          REG_DWORD)
    
    ' Delete value to enable DHCP
    DeleteKeys HKEY_LOCAL_MACHINE, _
               NET_ADAPTER_REG & tGUID, _
               "DisableDhcpOnConnect", _
                 REG_DWORD

    If GetDeviceByPNPDevice(tPNP, tDev) Then
    
        If EnableDevice(tDev.Index, False) Then
            
            bOK = EnableDevice(tDev.Index, True)
        
        End If
    
    End If

    
    SetNetConfig = True

End Function

Private Function SetRegIPAddress(sValueName As String, _
                                 sValueData As Variant, _
                                 lValueType As Long) As Boolean
    
    Dim lError As Long
    
    SetRegIPAddress = True
    
    lError = SetKeyValue(HKEY_LOCAL_MACHINE, _
                NET_ADAPTER_REG & gGUID, _
                sValueName, _
                sValueData, _
                lValueType)
                
    If lError <> 0 Then SetRegIPAddress = False
End Function

Public Function GetDomainName() As String
Dim objWMIService, colItems, objItem

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_ComputerSystem", , 48)
    
    For Each objItem In colItems
        GetDomainName = GetProp(objItem.Domain)
    Next
End Function

Public Function AdapterTypeToStr(iAdapterType As Integer) As String

    Select Case iAdapterType
        Case 0: AdapterTypeToStr = "Ethernet 802.3"
        Case 1: AdapterTypeToStr = "Token Ring 802.5"
        Case 2: AdapterTypeToStr = "Fiber Distributed Data Interface (FDDI)"
        Case 3: AdapterTypeToStr = "Wide Area Network (WAN)"
        Case 4: AdapterTypeToStr = "LocalTalk"
        Case 5: AdapterTypeToStr = "Ethernet using DIX header format"
        Case 6: AdapterTypeToStr = "ARCNET"
        Case 7: AdapterTypeToStr = "ARCNET (878.2)"
        Case 8: AdapterTypeToStr = "ATM"
        Case 9: AdapterTypeToStr = "Wireless"
        Case 10: AdapterTypeToStr = "Infrared Wireless"
        Case 11: AdapterTypeToStr = "Bpc"
        Case 12: AdapterTypeToStr = "CoWan"
        Case 13: AdapterTypeToStr = "1394"
    End Select
End Function
