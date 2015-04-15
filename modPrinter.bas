Attribute VB_Name = "hwPrinter"
Option Explicit

Private Type PrinterDetails
    DriverName As String
    Caption As String
    PortName As String
    DeviceID As String
End Type

Public Function EnumPrinters() As HardwarePrintDevice()
On Error Resume Next

    Dim objWMIService
    Dim objItem
    Dim oReg
    Dim colItems
    Dim tmpArr() As HardwarePrintDevice
    Dim idx, i              As Integer
    Dim printer_ip_address As String
    Dim printer_hostname   As String
    Dim printer_mfg        As String
    Dim printer_model      As String
    Dim connection_status  As String
    Dim print_color()      As String
    Dim print_duplex()     As String
    Dim printer_share_name As String
    Dim printer_reg_driver As String
    Dim printer_shared     As Boolean

    idx = 0
    ReDim tmpArr(idx)
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Printer", , 48)

    For Each objItem In colItems

        Dim pnt As PrinterDetails

        With pnt
            .Caption = GetProp(objItem.Caption)
            .DeviceID = GetProp(objItem.DeviceID)
            .DriverName = GetProp(objItem.DriverName)
            .PortName = GetProp(objItem.PortName)
        End With

        If IsSoftwarePrinter(pnt) = False Then
            printer_ip_address = ""
            printer_hostname = ""
            printer_model = ""
            printer_mfg = ""
            printer_share_name = ""
            connection_status = "Disconnected"
            ReDim Preserve tmpArr(idx)

            If ((InStr(objItem.PortName, "DOT")) = 0) And ((InStr(objItem.PortName, "COM")) = 0) Then
                RegExIsMatch GetProp(objItem.Name), REG_EX_IP_ADDRESS, printer_ip_address
            End If

            If ((printer_ip_address = "") And ((InStr(objItem.PortName, "DOT")) = 0) And ((InStr(objItem.PortName, "COM")) = 0)) Then
                printer_hostname = Replace(GetProp(objItem.servername), "\\", "")
            End If

            If ((printer_hostname > "") And (printer_ip_address = "") And ((InStr(objItem.PortName, "DOT")) = 0) And ((InStr(objItem.PortName, "COM")) = 0)) Then
                'printer_ip_address = DNSResolver(printer_hostname)
                printer_ip_address = HostNameToIP(printer_hostname)
            End If

            If (printer_ip_address > "") Then
                'printer_hostname = DNSResolver(printer_ip_address)
                IP2HostName printer_ip_address, printer_hostname
            End If

            If (printer_ip_address > "") Then

                On Error Resume Next

                Dim colPingResults, objPingResult, error_returned

                Set colPingResults = objWMIService.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '" & printer_ip_address & "'", , 32)
                error_returned = Err.Number

                For Each objPingResult In colPingResults

                    If Not IsObject(objPingResult) Then
                        connection_status = "No"
                    ElseIf objPingResult.StatusCode = 0 Then
                        connection_status = "Yes"
                    Else
                        connection_status = "No"
                    End If

                Next

                Set colPingResults = Nothing

                On Error GoTo 0

            End If

            If (InStr(objItem.PortName, "USB") = 1) Then

                Dim colItems2, objitem2

                On Error Resume Next

                Set colItems2 = objWMIService.ExecQuery("Select * FROM Win32_PnPEntity where Name = '" & objItem.DriverName & "'", , 32)

                On Error GoTo 0

                If (error_returned = 0) Then

                    For Each objitem2 In colItems2
                        connection_status = GetProp(objitem2.status)
                    Next

                Else
                    connection_status = "OK"
                End If

                Set colItems2 = Nothing
            End If

            If (InStr(objItem.PortName, "DOT") = 1) Then
                connection_status = "OK"
            End If

            If (InStr(objItem.PortName, "LPT1") = 1) Then
                connection_status = "OK"
            End If

            printer_model = Replace(GetProp(objItem.DriverName), " PCL 5e", "")
            printer_model = Replace(printer_model, " PCL 5", "")
            printer_model = Replace(printer_model, " PCL5", "")
            printer_model = Replace(printer_model, " PCL 6e", "")
            printer_model = Replace(printer_model, " PCL 6", "")
            printer_model = Replace(printer_model, " PCL6", "")
            printer_model = Replace(printer_model, " PCL", "")
            printer_model = Replace(printer_model, " PS", "")
            printer_mfg = GetPrinterMfg(printer_model)

            If (GetProp(objItem.ShareName) > "") Then
                printer_shared = True
                printer_share_name = GetProp(objItem.ShareName) & _
                    IIf(Len(printer_hostname) > 0, " on " & printer_hostname, "")
            Else
                printer_shared = False
                printer_share_name = ""
            End If

            With tmpArr(idx)
                .DeviceID = GetProp(objItem.DeviceID)
                .IPAddress = printer_ip_address
                .hostname = printer_hostname
                .ConnectionStat = connection_status
                .Model = printer_model
                .Manufacturer = printer_mfg
                .IsDefault = objItem.Default
                .IsLocal = objItem.Local
                .IsNetwork = objItem.Network
                .IsShared = objItem.Shared
                .DriverName = GetProp(objItem.DriverName)
                .PortName = GetProp(objItem.PortName)
                .ShareName = printer_share_name
                .IsOnline = Not objItem.WorkOffline
            End With

            idx = idx + 1
        End If

    Next

    EnumPrinters = tmpArr
    Set objWMIService = Nothing
    Set colItems = Nothing
End Function

Private Function GetPrinterMfg(printer_model As String) As String

    If (InStr(1, printer_model, "Aficio", vbTextCompare) = 1) Then GetPrinterMfg = "Ricoh"
    If (InStr(1, printer_model, "AGFA", vbTextCompare) = 1) Then GetPrinterMfg = "Agfa"
    If (InStr(1, printer_model, "Apple Laser", vbTextCompare) = 1) Then GetPrinterMfg = "Apple"
    If (InStr(1, printer_model, "Brother", vbTextCompare) = 1) Then GetPrinterMfg = "Brother"
    If (InStr(1, printer_model, "Canon", vbTextCompare) = 1) Then GetPrinterMfg = "Canon"
    If (InStr(1, printer_model, "Color-MFPe", vbTextCompare) = 1) Then GetPrinterMfg = "Toshiba"
    If (InStr(1, printer_model, "Datamax", vbTextCompare) = 1) Then GetPrinterMfg = "Datamax"
    If (InStr(1, printer_model, "Dell", vbTextCompare) = 1) Then GetPrinterMfg = "Dell"
    If (InStr(1, printer_model, "DYMO", vbTextCompare) = 1) Then GetPrinterMfg = "Dymo"
    If (InStr(1, printer_model, "EasyCoder", vbTextCompare) = 1) Then GetPrinterMfg = "Intermec"
    If (InStr(1, printer_model, "Epson", vbTextCompare) = 1) Then GetPrinterMfg = "Epson"
    If (InStr(1, printer_model, "Fiery", vbTextCompare) = 1) Then GetPrinterMfg = "Konica Minolta"
    If (InStr(1, printer_model, "Fuji", vbTextCompare) = 1) Then GetPrinterMfg = "Fuji"
    If (InStr(1, printer_model, "FX ApeosPort", vbTextCompare) = 1) Then GetPrinterMfg = "Fuji"
    If (InStr(1, printer_model, "FX DocuCentre", vbTextCompare) = 1) Then GetPrinterMfg = "Fuji"
    If (InStr(1, printer_model, "FX DocuPrint", vbTextCompare) = 1) Then GetPrinterMfg = "Fuji"
    If (InStr(1, printer_model, "FX DocuWide", vbTextCompare) = 1) Then GetPrinterMfg = "Fuji"
    If (InStr(1, printer_model, "FX Document", vbTextCompare) = 1) Then GetPrinterMfg = "Xerox"
    If (InStr(1, printer_model, "GelSprinter", vbTextCompare) = 1) Then GetPrinterMfg = "Ricoh"
    If (InStr(1, printer_model, "HP ", vbTextCompare) = 1) Then GetPrinterMfg = "Hewlett Packard"
    If (InStr(1, printer_model, "Konica", vbTextCompare) = 1) Then GetPrinterMfg = "Konica Minolta"
    If (InStr(1, printer_model, "Kyocera", vbTextCompare) = 1) Then GetPrinterMfg = "Kyocera Mita"
    If (InStr(1, printer_model, "LAN-Fax", vbTextCompare) = 1) Then GetPrinterMfg = "Ricoh"
    If (InStr(1, printer_model, "Lexmark", vbTextCompare) = 1) Then GetPrinterMfg = "Lexmark"
    If (InStr(1, printer_model, "Mita", vbTextCompare) = 1) Then GetPrinterMfg = "Kyocera-Mita"
    If (InStr(1, printer_model, "Muratec", vbTextCompare) = 1) Then GetPrinterMfg = "Muratec"
    If (InStr(1, printer_model, "Oce", vbTextCompare) = 1) Then GetPrinterMfg = "Oce"
    If (InStr(1, printer_model, "Oki", vbTextCompare) = 1) Then GetPrinterMfg = "Oki"
    If (InStr(1, printer_model, "Panaboard", vbTextCompare) = 1) Then GetPrinterMfg = "Panasonic"
    If (InStr(1, printer_model, "Ricoh", vbTextCompare) = 1) Then GetPrinterMfg = "Ricoh"
    If (InStr(1, printer_model, "Samsung", vbTextCompare) = 1) Then GetPrinterMfg = "Samsung"
    If (InStr(1, printer_model, "Sharp", vbTextCompare) = 1) Then GetPrinterMfg = "Sharp"
    If (InStr(1, printer_model, "SP 3", vbTextCompare) = 1) Then GetPrinterMfg = "Ricoh"
    If (InStr(1, printer_model, "Tektronix", vbTextCompare) = 1) Then GetPrinterMfg = "Tektronix"
    If (InStr(1, printer_model, "Toshiba", vbTextCompare) = 1) Then GetPrinterMfg = "Toshiba"
    If (InStr(1, printer_model, "Xerox", vbTextCompare) = 1) Then GetPrinterMfg = "Xerox"
    If (InStr(1, printer_model, "ZDesigner", vbTextCompare) = 1) Then GetPrinterMfg = "Zebra"
    If (InStr(1, printer_model, "Zebra", vbTextCompare) = 1) Then GetPrinterMfg = "Zebra"
    If GetPrinterMfg = "" Then GetPrinterMfg = printer_model
End Function

Private Function IsSoftwarePrinter(prnt As PrinterDetails) As Boolean

    Dim blnSoftware As Boolean

    blnSoftware = False

    With prnt

        If (InStr(1, .DriverName, "ActiveFax", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "AdobePS Acrobat Distiller", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Amyuni Document Converter", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Autodesk DWFWriter", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Black Ice", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Canon DM Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Canon iW Image Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Citrix Universal Printer", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Document Publisher Plus Printer Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "DocuWorks Printer Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "e-BRIDGE Viewer", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Fax", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "FRx Document Image Writer Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Generic / Text Only", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "HP Universal Printing", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Journal Note Writer Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "MacromediaFlashPaper", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Microsoft", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "mimio Print Capture Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Nitro Reader", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Nuance Image Printer Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "PaperPort", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "PCL6 Driver fif Universal Print", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "PS Driver fif Universal Print", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "PDF", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "SnagIt", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Snapshot 70 Driver", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "OneNote", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "Therefore", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "TIFF Image Printer", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "TP Output Gateway", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "WINPRINT-Kyocera", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .DriverName, "WebEx", vbTextCompare)) Then blnSoftware = True
        If (InStr(1, .Caption, "PDF", vbTextCompare)) Then blnSoftware = True
        If (InStr(LCase$(.PortName), "ts0") = 1) Then blnSoftware = True
        If (InStr(LCase$(.PortName), "webex") = 1) Then blnSoftware = True
        If (InStr(LCase$(.PortName), "file") = 1) Then blnSoftware = True
        If (InStr(.PortName, "\\") = 1) Then blnSoftware = True
        If (InStr(LCase$(.PortName), "oletoadi") = 1) Then blnSoftware = True
        If (InStr(LCase$(.PortName), "client/") = 1) Then blnSoftware = True
        If (InStr(LCase$(.PortName), "client:") = 1) Then blnSoftware = True
        If (InStr(LCase$(.PortName), "lpt1:") = 1) Then blnSoftware = True
        If (InStr(1, LCase$(.DeviceID), "(copy ", vbTextCompare)) Then blnSoftware = True
    End With

    IsSoftwarePrinter = blnSoftware
End Function

Private Function ValidateIPAddress1(PortName As String) As Boolean

    Dim aTmp() As String
    Dim field  As Variant

    ValidateIPAddress1 = True
    aTmp = Split(PortName, ".")

    If UBound(aTmp) = 3 Then
        ValidateIPAddress1 = True

        For Each field In aTmp

            If (IsNumeric(field)) Then
                If (CInt(field) > 255) Then ValidateIPAddress1 = False
            Else
                ValidateIPAddress1 = False
            End If

        Next

    Else
        ValidateIPAddress1 = False
    End If

End Function
