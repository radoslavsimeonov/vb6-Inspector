Attribute VB_Name = "modReadPCIDatabase"
Option Explicit

Private Const PCI_DEV As String = "pcidevs.txt"
Private Const MON_DEV As String = "mondevs.txt"
Private Const PCM_DEV As String = "pcmdevs.txt"
Private Const PNP_DEV As String = "pnpdevs.txt"
Private Const USB_DEV As String = "usbdevs.txt"

Public Function ReadTextFile(hw As HWID) As HWDetails

    Dim FileNo         As Integer
    Dim FileName       As String
    Dim FilePath       As String
    Dim tmpData        As String
    Dim tmpHWD         As HWDetails

    Static C           As Collection
    Static tmpFileName As String

    Dim sRow           As Variant

    FileNo = FreeFile

    Select Case hw.Type
        Case "PCI", "HDAUDIO"
            FileName = PCI_DEV
        Case "DISPLAY", "MONITOR"
            FileName = MON_DEV
        Case "PCM"
            FileName = PCI_DEV
        Case "ACPI"
            FileName = PNP_DEV
        Case "USB", "HID"
            FileName = USB_DEV
        Case Else
            Exit Function
    End Select

    FilePath = App.Path & "\siv\" & FileName

    If FExists(FilePath) = False Then Exit Function
    If FileName <> tmpFileName Then
        Set C = Nothing
        tmpFileName = FileName
    End If

    If C Is Nothing Then Set C = New Collection
    If C.count = 0 Then
        Open FilePath For Input As #FileNo

        Do While Not EOF(FileNo)
            Input #FileNo, tmpData
            C.Add tmpData
        Loop

        Close
    End If

    For Each sRow In C

        Select Case hw.Type
            Case "USB", "PCI", "HID", "HDAUDIO"
                ParseLine sRow, hw, tmpHWD.Chip, tmpHWD.Vendor

                If tmpHWD.Vendor <> "" And tmpHWD.Chip <> "" Then Exit For
            Case Else

                If ParseLine(sRow, hw, vbNullString, tmpHWD.Chip) Then
                    Exit For
                End If

        End Select

    Next

    ReadTextFile = tmpHWD
End Function

Private Function ParseLine(ByVal sLine As String, _
                           hw As HWID, _
                           sModel As String, _
                           sManufacturer As String) As Boolean

    Dim Split1() As String
    Dim Split2() As String
    Dim sMod     As String

    Dim sIDs     As String

    ParseLine = False

    If InStr(sLine, "==") > 0 Then
        Split1 = Split(sLine, "==")
    ElseIf InStr(sLine, "=") Then
        Split1 = Split(sLine, "=")
    Else
        Exit Function
    End If

    sIDs = Trim$(UCase$(Split1(0)))

    If Len(sIDs) = 0 Then Exit Function
    sMod = Trim$(Split1(1))

    If InStr(sIDs, ":") > 0 Then
        Split2 = Split(sIDs, ":")
    Else

        If sIDs = UCase$(hw.VEN) Then
            sManufacturer = sMod
            ParseLine = True
            'TODO: return manufacturer only
            Exit Function
        End If
    End If

    If Len(Trim$(Join$(Split2))) = 0 Then Exit Function
    If UBound(Split2) = 1 And InStr(sIDs, UCase$(hw.VEN & ":" & hw.dev)) = 1 Then
        sModel = sMod
    End If

End Function
