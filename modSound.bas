Attribute VB_Name = "hwSound"
Option Explicit

Public Function EnumSoundDevices() As HardwareSoundDevice()

    Dim objWMIService, colItems, objItem
    Dim tSD()   As HardwareSoundDevice
    Dim cnt     As Integer
    
    cnt = 0
    ReDim tSD(cnt)
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_SoundDevice", , 48)
        
    For Each objItem In colItems
        ReDim Preserve tSD(cnt)
        
        With tSD(cnt)
            .Manufacturer = GetProp(objItem.Manufacturer)
            .Model = GetProp(objItem.Name)
            .PNPDeviceId = GetProp(objItem.PNPDeviceId)
        End With
        
        cnt = cnt + 1
    Next
    
    EnumSoundDevices = tSD
End Function
