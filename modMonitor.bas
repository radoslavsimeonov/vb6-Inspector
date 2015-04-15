Attribute VB_Name = "hwMonitor"
Option Explicit

Private Type Dimensions
    HSize As String
    VSize As String
    Size  As String
End Type

Public Function EnumMonitors() As HardwareDesktopMonitor()

    Dim strEDID          As String
    Dim arrWMIMonitors() As String
    Dim pnpDevice        As Variant
    Dim idx              As Integer
    Dim tmpDims          As Dimensions
    Dim tmpArr()         As HardwareDesktopMonitor

    idx = 0
    ReDim tmpArr(idx)
    arrWMIMonitors = GetMonitorPNPDevices

    If Len(Join$(arrWMIMonitors)) = 0 Then
        EnumMonitors = tmpArr
        Exit Function
    End If

    If UBound(arrWMIMonitors) > -1 Then

        For Each pnpDevice In arrWMIMonitors
            strEDID = ReadEDID(pnpDevice)

            If strEDID <> "{ERROR}" Then
                ReDim Preserve tmpArr(idx)

                With tmpArr(idx)
                    .PNPDeviceId = pnpDevice
                    .SerialNumber = GetSerialFromEDID(strEDID)
                    .Manufacturer = GetMfgFromEDID(strEDID)
                    .Model = GetModelFromEDID(strEDID)
                    .ModelId = GetDevFromEDID(strEDID)
                    .ManufacureDate = GetMfgDateFromEDID(strEDID)
                    .VideoInput = GetVideoInputFromEDID(strEDID)
                    .AspectRatio = GetAspectRatioFromEDID(strEDID)
                    tmpDims = GetMonitorDimensionsFromEDID(strEDID)
                    .HSize = tmpDims.HSize
                    .VSize = tmpDims.VSize
                    .Size = tmpDims.Size
                End With

                idx = idx + 1
            End If

        Next

    End If

    EnumMonitors = tmpArr
End Function

Private Function ReadEDID(ByVal strPNPDevice As String) As String
On Error Resume Next

    Dim strRegKey As String
    Dim objReg, bVal, bEDID

    strRegKey = "SYSTEM\CurrentControlSet\Enum\" & strPNPDevice & "\Device Parameters"
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!//./root/default:StdRegProv")
    objReg.GetBinaryValue HKEY_LOCAL_MACHINE, strRegKey, "EDID", bEDID

    If VarType(bEDID) <> 8204 Then
        ReadEDID = "{ERROR}"
    Else

        For Each bVal In bEDID
            ReadEDID = ReadEDID + Chr$(bVal)
        Next

    End If

End Function

Private Function GetMonitorPNPDevices() As String()
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim tmpArr() As String
    Dim x        As Integer

    x = 0
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DesktopMonitor", , 48)

    If Not IsNull(colItems) Then

        For Each objItem In colItems

            If Not IsNull(objItem.PNPDeviceId) Then
                ReDim Preserve tmpArr(x)
                tmpArr(x) = objItem.PNPDeviceId
                x = x + 1
            End If

        Next

    End If

    GetMonitorPNPDevices = tmpArr
    Set objWMIService = Nothing
    Set colItems = Nothing
End Function

Private Function GetDescriptorBlockFromEDID(strEDID As String, strTag As String) As String

    Dim strFoundBlock As String
    Dim strResult     As String
    Dim arrDescriptorBlock(3)

    arrDescriptorBlock(0) = Mid$(strEDID, &H36 + 1, 18)
    arrDescriptorBlock(1) = Mid$(strEDID, &H48 + 1, 18)
    arrDescriptorBlock(2) = Mid$(strEDID, &H5A + 1, 18)
    arrDescriptorBlock(3) = Mid$(strEDID, &H6C + 1, 18)

    If InStr(arrDescriptorBlock(0), strTag) > 0 Then
        strFoundBlock = arrDescriptorBlock(0)
    ElseIf InStr(arrDescriptorBlock(1), strTag) > 0 Then
        strFoundBlock = arrDescriptorBlock(1)
    ElseIf InStr(arrDescriptorBlock(2), strTag) > 0 Then
        strFoundBlock = arrDescriptorBlock(2)
    ElseIf InStr(arrDescriptorBlock(3), strTag) > 0 Then
        strFoundBlock = arrDescriptorBlock(3)
    Else
        GetDescriptorBlockFromEDID = ""
        Exit Function
    End If

    strResult = Right$(strFoundBlock, 14)

    If InStr(strResult, Chr$(&HA)) > 0 Then
        strResult = Trim$(Left$(strResult, InStr(strResult, Chr$(&HA)) - 1))
    Else
        strResult = Trim$(strResult)
    End If

    If Left$(strResult, 1) = Chr$(0) Then strResult = Right$(strResult, Len(strResult) - 1)
    GetDescriptorBlockFromEDID = strResult
End Function

Private Function GetSerialFromEDID(strEDID As String) As String

    Dim strTag As String

    strTag = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF)
    GetSerialFromEDID = GetDescriptorBlockFromEDID(strEDID, strTag)
End Function

Function GetModelFromEDID(strEDID As String) As String

    Dim strTag As String

    strTag = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFC)
    GetModelFromEDID = GetDescriptorBlockFromEDID(strEDID, strTag)
End Function

Private Function GetMfgFromEDID(strEDID As String) As String
On Error Resume Next

    Dim tmpEDIDMfg, tmpMfg As String
    Dim Char1, Char2, Char3 As Byte
    Dim Byte1, Byte2 As Byte

    tmpEDIDMfg = Mid$(strEDID, &H8 + 1, 2)
    Char1 = 0
    Char2 = 0
    Char3 = 0
    Byte1 = Asc(Left$(tmpEDIDMfg, 1))
    Byte2 = Asc(Right$(tmpEDIDMfg, 1))

    If (Byte1 And 64) > 0 Then Char1 = Char1 + 16
    If (Byte1 And 32) > 0 Then Char1 = Char1 + 8
    If (Byte1 And 16) > 0 Then Char1 = Char1 + 4
    If (Byte1 And 8) > 0 Then Char1 = Char1 + 2
    If (Byte1 And 4) > 0 Then Char1 = Char1 + 1
    If (Byte1 And 2) > 0 Then Char2 = Char2 + 16
    If (Byte1 And 1) > 0 Then Char2 = Char2 + 8
    If (Byte2 And 128) > 0 Then Char2 = Char2 + 4
    If (Byte2 And 64) > 0 Then Char2 = Char2 + 2
    If (Byte2 And 32) > 0 Then Char2 = Char2 + 1
    Char3 = Char3 + (Byte2 And 16)
    Char3 = Char3 + (Byte2 And 8)
    Char3 = Char3 + (Byte2 And 4)
    Char3 = Char3 + (Byte2 And 2)
    Char3 = Char3 + (Byte2 And 1)
    tmpMfg = Chr$(Char1 + 64) & Chr$(Char2 + 64) & Chr$(Char3 + 64)
    GetMfgFromEDID = ConvertMfgAbbr(tmpMfg)
End Function

Private Function ConvertMfgAbbr(strMfg As String) As String

    Dim tmpMfg As String

    Select Case strMfg
        Case "ACR":            tmpMfg = "Acer"
        Case "ACT":            tmpMfg = "Targa"
        Case "ADI":            tmpMfg = "ADI"
        Case "AOC":            tmpMfg = "AOC"
        Case "API":            tmpMfg = "Acer"
        Case "APP":            tmpMfg = "Apple"
        Case "ART":            tmpMfg = "ArtMedia"
        Case "AST":            tmpMfg = "AST Research"
        Case "CPL":            tmpMfg = "Compal"
        Case "CPQ":            tmpMfg = "Compaq"
        Case "CTX":            tmpMfg = "Chuntex"
        Case "DEC":            tmpMfg = "Digital Equipment Corporation"
        Case "DEL":            tmpMfg = "Dell"
        Case "DPC":            tmpMfg = "Delta"
        Case "DWE":            tmpMfg = "Daewoo"
        Case "ECS":            tmpMfg = "Elitegroup Computer Systems"
        Case "EIZ":            tmpMfg = "EIZO"
        Case "EPI":            tmpMfg = "Envision"
        Case "FCM":            tmpMfg = "Funai"
        Case "FUS":            tmpMfg = "Fujitsu Siemens"
        Case "GSM":            tmpMfg = "LG Electronics"
        Case "GWY":            tmpMfg = "Gateway 2000"
        Case "HEI":            tmpMfg = "Hyundai"
        Case "HIT":            tmpMfg = "Hitachi"
        Case "HSD":            tmpMfg = "Hanns.G"
        Case "HSL":            tmpMfg = "Hansol Electronics"
        Case "HTC":            tmpMfg = "Hitachi"
        Case "HWP":            tmpMfg = "Hewlett Packard"
        Case "IBM":            tmpMfg = "IBM"
        Case "ICL":            tmpMfg = "Fujitsu"
        Case "IVM":            tmpMfg = "Idek Iiyama"
        Case "KFC":            tmpMfg = "KFC Computek"
        Case "LEN":            tmpMfg = "Lenovo"
        Case "LGD":            tmpMfg = "LG Display"
        Case "LKM":            tmpMfg = "ADLAS / AZALEA"
        Case "LNK":            tmpMfg = "LINK Technologies"
        Case "LTN":            tmpMfg = "Lite-On"
        Case "MAG":            tmpMfg = "MAG InnoVision"
        Case "MAX":            tmpMfg = "Maxdata Computer"
        Case "MEI":            tmpMfg = "Panasonic"
        Case "MEL":            tmpMfg = "Mitsubishi Electronics"
        Case "MIR":            tmpMfg = "Miro"
        Case "MTC":            tmpMfg = "MITAC"
        Case "NAN":            tmpMfg = "NANAO"
        Case "NEC":            tmpMfg = "NEC"
        Case "NOK":            tmpMfg = "Nokia"
        Case "OQI":            tmpMfg = "Optiquest"
        Case "PBN":            tmpMfg = "Packard Bell"
        Case "PGS":            tmpMfg = "Princeton Graphic Systems"
        Case "PHL":            tmpMfg = "Philips"
        Case "PNP":            tmpMfg = "Plug n Play (Microsoft"
        Case "REL":            tmpMfg = "Relisys"
        Case "SAM":            tmpMfg = "Samsung"
        Case "SEC":            tmpMfg = "Samsung"
        Case "SMI":            tmpMfg = "Smile"
        Case "SMC":            tmpMfg = "Samtron"
        Case "SNI":            tmpMfg = "Siemens Nixdorf"
        Case "SNY":            tmpMfg = "Sony Corporation"
        Case "SPT":            tmpMfg = "Sceptre"
        Case "SRC":            tmpMfg = "Shamrock Technology"
        Case "STN":            tmpMfg = "Samtron"
        Case "STP":            tmpMfg = "Sceptre"
        Case "TAT":            tmpMfg = "Tatung"
        Case "TRL":            tmpMfg = "Royal Information Company"
        Case "TOS":            tmpMfg = "Toshiba"
        Case "TSB":            tmpMfg = "Toshiba"
        Case "UNM":            tmpMfg = "Unisys"
        Case "VSC":            tmpMfg = "ViewSonic"
        Case "WTC":            tmpMfg = "Wen Technology"
        Case "ZCM":            tmpMfg = "Zenith Data Systems"
        Case "___":            tmpMfg = "Targa"
        Case Else:             tmpMfg = strMfg
    End Select

    tmpMfg = Replace(tmpMfg, "@", "")
    tmpMfg = Replace(tmpMfg, "%", "")
    tmpMfg = Replace(tmpMfg, ";", "")
    ConvertMfgAbbr = tmpMfg
End Function

Private Function GetDevFromEDID(strEDID As String) As String
On Error Resume Next

    Dim tmpEDIDDev1, tmpEDIDDev2 As Byte
    Dim tmpDev As String

    tmpEDIDDev1 = Hex$(Asc(Mid$(strEDID, &HA + 1, 1)))
    tmpEDIDDev2 = Hex$(Asc(Mid$(strEDID, &HB + 1, 1)))

    If Len(tmpEDIDDev1) = 1 Then tmpEDIDDev1 = "0" & tmpEDIDDev1
    If Len(tmpEDIDDev2) = 1 Then tmpEDIDDev2 = "0" & tmpEDIDDev2
    tmpDev = tmpEDIDDev2 & tmpEDIDDev1
    GetDevFromEDID = tmpDev
End Function

Private Function GetMfgDateFromEDID(strEDID As String) As String
On Error Resume Next

    Dim tmpmfgweek, tmpmfgyear, tmpmdt As String

    tmpmfgweek = Asc(Mid$(strEDID, &H10 + 1, 1))
    tmpmfgyear = (Asc(Mid$(strEDID, &H11 + 1, 1))) + 1990
    tmpmdt = Month(DateAdd("ww", tmpmfgweek, DateValue("1/1/" & tmpmfgyear))) & "/" & tmpmfgyear
    GetMfgDateFromEDID = tmpmdt
End Function

Private Function GetVideoInputFromEDID(strEDID As String) As String
On Error Resume Next

    Dim tmpEDIDInp As Byte

    tmpEDIDInp = Asc(Mid$(strEDID, &H14 + 1, 1))
    GetVideoInputFromEDID = IIf((tmpEDIDInp And 128) = 0, "Analog (D-Sub)", "Digital (DVI)")
End Function

Private Function GetMonitorDimensionsFromEDID(strEDID As String) As Dimensions
On Error Resume Next

    Dim tmpDims As Dimensions
    Dim dDiag   As Double

    With tmpDims
        .HSize = Asc(Mid$(strEDID, &H15 + 1, 1))
        .VSize = Asc(Mid$(strEDID, &H16 + 1, 1))
        dDiag = Sqr((.HSize * .HSize) + (.VSize * .VSize)) * 0.393700787401575
        .Size = Round(dDiag, 1)
    End With

    GetMonitorDimensionsFromEDID = tmpDims
End Function

Private Function GetAspectRatioFromEDID(strEDID As String) As String

    Dim ratio As String
    Dim dV, eV As Double
    Dim tmpEDIDMajorVer, tmpEDIDRev, edid_version

    If (Asc(Mid$(strEDID, &H37 + 1, 1)) And 128) Then ratio = "1:" Else ratio = "0:"
    If (Asc(Mid$(strEDID, &H37 + 1, 1)) And 64) Then ratio = ratio & "1" Else ratio = ratio & "0"
    tmpEDIDMajorVer = Asc(Mid$(strEDID, &H12 + 1, 1))
    tmpEDIDRev = Asc(Mid$(strEDID, &H13 + 1, 1))
    edid_version = Chr$(48 + tmpEDIDMajorVer) & "." & Chr$(48 + tmpEDIDRev)

    Select Case ratio
        Case "0:0"

            If CDbl(val(edid_version)) > 1.3 Then ratio = "16:10"
            If CDbl(val(edid_version)) <= 1.3 Then ratio = "16:9"
        Case "0:1": ratio = "4:3"
        Case "1:0": ratio = "5:4"
        Case "1:1": ratio = "16:9"
    End Select

    'GetAspectRatioFromEDID = ratio
    GetAspectRatioFromEDID = ""
End Function
