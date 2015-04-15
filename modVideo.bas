Attribute VB_Name = "hwVideo"
Option Explicit

Public Function EnumVideoAdapters() As HardwareVideoAdapter()
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim tmpArr() As HardwareVideoAdapter
    Dim idx      As Integer

    idx = 0
    ReDim tmpArr(idx)
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController", , 48)

    For Each objItem In colItems
        ReDim Preserve tmpArr(idx)

        With tmpArr(idx)
            .Model = GetProp(objItem.Name)
            .Manufacturer = GetProp(objItem.AdapterCompatibility)
            .PNPDeviceId = GetProp(objItem.PNPDeviceId)

            If Not IsNull(objItem.AdapterRAM) Then
                .VideoRAM = GetProp(objItem.AdapterRAM) '/ 1024 / 1024
            End If

        End With

        idx = idx + 1
    Next

    Set objWMIService = Nothing
    Set colItems = Nothing
    EnumVideoAdapters = tmpArr
End Function
