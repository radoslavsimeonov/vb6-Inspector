Attribute VB_Name = "hwCPU"
Option Explicit

Public Function GetCPU() As HardwareCPU
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim tmpArr() As String

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor ", , 48)

    For Each objItem In colItems

        Dim cp As HardwareCPU

        If Len(Trim$(objItem.Name)) > 0 Then

            With cp
                .Architecture = GetProp(objItem.DataWidth)
                .ClockSpeed = GetProp(objItem.MaxClockSpeed)
                .Cores = GetProp(objItem.NumberOfCores)
                .LogicalProcessors = GetProp(objItem.NumberOfLogicalProcessors)
                .Model = CPUNameClear(GetProp(GetProcessorName))
                .Manufacturer = CPUNameClear(GetProp(objItem.Manufacturer))
                .Socket = GetProp(objItem.upgradeMethod)
            End With

        End If

    Next

    Set objWMIService = Nothing
    Set colItems = Nothing
    GetCPU = cp
End Function

Public Function CPU_Socket(upgradeMethod) As String

    Select Case upgradeMethod
        Case "1"
            CPU_Socket = "Other"
        Case "2"
            CPU_Socket = "Unknown"
        Case "3"
            CPU_Socket = "Daughter Board"
        Case "4"
            CPU_Socket = "ZIF Socket"
        Case "5"
            CPU_Socket = "Replaceable Piggy Back"
        Case "6"
            CPU_Socket = "None"
        Case "7"
            CPU_Socket = "LIF Socket"
        Case "8"
            CPU_Socket = "Slot 1"
        Case "9"
            CPU_Socket = "Slot 2"
        Case "10"
            CPU_Socket = "370 Pin Socket"
        Case "11"
            CPU_Socket = "Slot A"
        Case "12"
            CPU_Socket = "Slot M"
        Case "13"
            CPU_Socket = "Socket 423"
        Case "14"
            CPU_Socket = "Socket A (462)"
        Case "15"
            CPU_Socket = "Socket 478"
        Case "16"
            CPU_Socket = "Socket 754"
        Case "17"
            CPU_Socket = "Socket 940"
        Case "18"
            CPU_Socket = "Socket 939"
        Case "19"
            CPU_Socket = "Socket mPGA 604"
        Case "20"
            CPU_Socket = "Socket LGA 771"
        Case "21"
            CPU_Socket = "Socket LGA 775"
        Case "22"
            CPU_Socket = "Socket S1"
        Case "23"
            CPU_Socket = "Socket AM2"
        Case "24"
            CPU_Socket = "Socket F (1207)"
        Case "25"
            CPU_Socket = "Socket LGA 1366"
        Case "26"
            CPU_Socket = "Socket G34"
        Case "27"
            CPU_Socket = "Socket AM3"
        Case "28"
            CPU_Socket = "Socket C32"
        Case "29"
            CPU_Socket = "Socket LGA 1156"
        Case "30"
            CPU_Socket = "Socket LGA 1567"
        Case "31"
            CPU_Socket = "Socket PGA 988A"
        Case "32"
            CPU_Socket = "Socket BGA 1288"
        Case "33"
            CPU_Socket = "Socket rPGA 988B"
        Case "34"
            CPU_Socket = "Socket BGA 1023"
        Case "35"
            CPU_Socket = "Socket BGA 1224"
        Case "36"
            CPU_Socket = "Socket LGA 1155"
        Case "37"
            CPU_Socket = "Socket LGA 1356"
        Case "38"
            CPU_Socket = "Socket LGA 2011"
        Case "39"
            CPU_Socket = "Socket FS1"
        Case "40"
            CPU_Socket = "Socket FS2"
        Case "41"
            CPU_Socket = "Socket FM1"
        Case "42"
            CPU_Socket = "Socket FM2"
        Case "43"
            CPU_Socket = "Socket LGA 2011-3"
        Case "44"
            CPU_Socket = "Socket LGA 1356-3"
        Case "185"
            CPU_Socket = "Socket P (478)"
        Case Default
            CPU_Socket = "Unknown"
    End Select

End Function

Private Function CPUNameClear(sCPU As String) As String
    CPUNameClear = Trim$(sCPU)
    CPUNameClear = Replace(CPUNameClear, "Genuine", "")
    CPUNameClear = Replace(CPUNameClear, "Authentic", "")
    CPUNameClear = Replace(CPUNameClear, "(R)", "")
    CPUNameClear = Replace(CPUNameClear, "(TM)", "")
    CPUNameClear = Replace(CPUNameClear, "(r)", "")
    CPUNameClear = Replace(CPUNameClear, "(tm)", "")
    CPUNameClear = Replace(CPUNameClear, "  ", " ")
    CPUNameClear = Replace(CPUNameClear, "  ", " ")
    CPUNameClear = Replace(CPUNameClear, "  ", " ")
    CPUNameClear = Replace(CPUNameClear, "  ", " ")
    CPUNameClear = Replace(CPUNameClear, "  ", " ")
    CPUNameClear = Replace(CPUNameClear, "  ", " ")
    CPUNameClear = Replace(CPUNameClear, "  ", " ")
End Function
