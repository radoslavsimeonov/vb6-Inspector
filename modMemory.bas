Attribute VB_Name = "hwMemory"
Option Explicit

Public Function EnumRAMMemory() As HardwareRAMMemory
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim tmpMem      As HardwareRAMMemory
    Dim tmpBanks()  As HardwareRAMBank
    Dim tmpCapacity As Double
    Dim idx         As Integer

    idx = 0
    ReDim tmpBanks(idx)
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT BankLabel,Capacity,DeviceLocator,FormFactor,MemoryType,Speed FROM Win32_PhysicalMemory", , 48)

    For Each objItem In colItems
        ReDim Preserve tmpBanks(idx)

        With tmpBanks(idx)
            .BankLabel = GetProp(objItem.BankLabel)
            .Capacity = GetProp(objItem.Capacity) ' / 1048576
            .Location = GetProp(objItem.DeviceLocator)
            .Speed = GetProp(objItem.Speed)
            .Type = GetProp(objItem.MemoryType)
            .FormFactor = GetProp(objItem.FormFactor)
            
            tmpCapacity = tmpCapacity + .Capacity
            
            idx = idx + 1
        End With

    Next

    Set objWMIService = Nothing
    Set colItems = Nothing
    tmpMem.Banks = tmpBanks
    tmpMem.TotalMemory = tmpCapacity
    EnumRAMMemory = tmpMem
End Function

Public Function RAMType(Value As Integer) As String

    Select Case Value
        Case 0
            RAMType = "Unknown"
        Case 1
            RAMType = "Other"
        Case 2
            RAMType = "DRAM"
        Case 3
            RAMType = "SynchronousDRAM"
        Case 4
            RAMType = "CacheDRAM"
        Case 5
            RAMType = "EDO"
        Case 6
            RAMType = "EDRAM"
        Case 7
            RAMType = "VRAM"
        Case 8
            RAMType = "SRAM"
        Case 9
            RAMType = "ram"
        Case 10
            RAMType = "ROM"
        Case 11
            RAMType = "Flash"
        Case 12
            RAMType = "EEPROM"
        Case 13
            RAMType = "FEPROM"
        Case 14
            RAMType = "EPROM"
        Case 15
            RAMType = "CDRAM"
        Case 16
            RAMType = "D3RAM"
        Case 17
            RAMType = "SDRAM"
        Case 18
            RAMType = "SGRAM"
        Case 19
            RAMType = "RDRAM"
        Case 20
            RAMType = "DDR"
        Case 21
            RAMType = "DDR2"
        Case Is >= 22
            RAMType = "DDR?"
    End Select

End Function

Public Function RAMFormFactor(Value As Integer) As String

    Select Case Value
        Case "1"
            RAMFormFactor = "Other"
        Case "2"
            RAMFormFactor = "SIP"
        Case "3"
            RAMFormFactor = "DIP"
        Case "4"
            RAMFormFactor = "ZIP"
        Case "5"
            RAMFormFactor = "SOJ"
        Case "6"
            RAMFormFactor = "Proprietary"
        Case "7"
            RAMFormFactor = "SIMM"
        Case "8"
            RAMFormFactor = "DIMM"
        Case "9"
            RAMFormFactor = "TSOP"
        Case "10"
            RAMFormFactor = "PGA"
        Case "11"
            RAMFormFactor = "RIMM"
        Case "12"
            RAMFormFactor = "SODIMM"
        Case "13"
            RAMFormFactor = "SRIMM"
        Case "14"
            RAMFormFactor = "SMD"
        Case "15"
            RAMFormFactor = "SSMP"
        Case "16"
            RAMFormFactor = "QFP"
        Case "17"
            RAMFormFactor = "TQFP"
        Case "18"
            RAMFormFactor = "SOIC"
        Case "19"
            RAMFormFactor = "LCC"
        Case "20"
            RAMFormFactor = "PLCC"
        Case "21"
            RAMFormFactor = "BGA"
        Case "22"
            RAMFormFactor = "FPBGA"
        Case "23"
            RAMFormFactor = "LGA"
        Case Else
            RAMFormFactor = "Unknown"
    End Select

End Function
