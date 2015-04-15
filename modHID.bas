Attribute VB_Name = "hwHID"
Option Explicit

Public Function GetHID() As HardwareHID
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim tmpArr() As String

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT Description FROM Win32_Keyboard", , 48)

    For Each objItem In colItems
        GetHID.Keyboard = GetProp(objItem.Description)
    Next

    Set colItems = objWMIService.ExecQuery("SELECT DeviceInterface FROM Win32_PointingDevice", , 48)

    For Each objItem In colItems

        Dim Msg As String

        Select Case objItem.DeviceInterface
            Case 1
                Msg = "Other"
            Case 2
                Msg = "Unknown"
            Case 3
                Msg = "Serial"
            Case 4
                Msg = "PS/2"
            Case 5
                Msg = "Infrared"
            Case 6
                Msg = "HP-HIL"
            Case 7
                Msg = "Bus mouse"
            Case 8
                Msg = "ADB (Apple Desktop Bus)"
            Case 160
                Msg = "Bus mouse DB-9"
            Case 161
                Msg = "Bus mouse micro-DIN"
            Case 162
                Msg = "USB"
        End Select

        GetHID.Mouse = Msg & " Mouse"
    Next

    Set objWMIService = Nothing
    Set colItems = Nothing
End Function
