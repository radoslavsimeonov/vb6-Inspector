Attribute VB_Name = "hwHardDisks"
Option Explicit

Private Type GETVERSIONOUTPARAMS
    bVersion As Byte 'Binary driver version.
    bRevision As Byte 'Binary driver revision
    bReserved As Byte 'Not used
    bIDEDeviceMap As Byte 'Bit map of IDE devices
    fCapabilities As Long 'Bit mask of driver capabilities
    dwReserved(3) As Long 'For future use
End Type

Private Type IDEREGS
    bFeaturesReg As Byte 'Used for specifying SMART "commands"
    bSectorCountReg As Byte 'IDE sector count register
    bSectorNumberReg As Byte 'IDE sector number register
    bCylLowReg As Byte 'IDE low order cylinder value
    bCylHighReg As Byte 'IDE high order cylinder value
    bDriveHeadReg As Byte 'IDE drive/head register
    bCommandReg As Byte 'Actual IDE command
    bReserved As Byte 'reserved for future use - must be zero
End Type

Private Type SENDCMDINPARAMS
    cBufferSize As Long 'Buffer size in bytes
    irDriveRegs As IDEREGS 'Structure with drive register values.
    bDriveNumber As Byte 'Physical drive number to send command to (0,1,2,3).
    bReserved(2) As Byte 'Bytes reserved
    dwReserved(3) As Long 'DWORDS reserved
    bBuffer() As Byte 'Input buffer.
End Type

Private Const IDE_ID_FUNCTION = &HEC 'Returns ID sector for ATA.
Private Const IDE_EXECUTE_SMART_FUNCTION = &HB0 'Performs SMART cmd.

Private Const SMART_CYL_LOW = &H4F
Private Const SMART_CYL_HI = &HC2

Private Type DRIVERSTATUS
    bDriverError As Byte 'Error code from driver, or 0 if no error
    bIDEStatus As Byte 'Contents of IDE Error register
    bReserved(1) As Byte
    dwReserved(1) As Long
End Type

Private Type DeviceData
    GeneralConfiguration As Integer            ' 0
    LogicalCylinders As Integer                ' 1  Obsolete
    SpecificConfiguration As Integer               ' 2
    LogicalHeads As Integer                    ' 3 Obsolete
    Retired1(1) As Integer                     ' 4-5
    LogicalSectors As Integer                      ' 6 Obsolete
    ReservedForCompactFlash(1) As Integer        ' 7-8
    Retired2 As Integer                        ' 9
    SerialNumber(19) As Byte                  ' 10-19
    Retired3 As Integer                        ' 20
    BufferSize As Integer                          ' 21 Obsolete
    ECCSize As Integer                         ' 22
    FirmwareRev(7) As Byte                       ' 23-26
    ModelNumber(39) As Byte                           ' 27-46
    MaxNumPerInterupt As Integer                   ' 47
    Reserved1 As Integer                           ' 48
    Capabilities1 As Integer                       ' 49
    Capabilities2 As Integer                       ' 50
    Obsolute5(1) As Integer                          ' 51-52
    Field88and7064 As Integer                      ' 53
    Obsolute6(4) As Integer                       ' 54-58
    MultSectorStuff As Integer                 ' 59
    TotalAddressableSectors(1) As Integer        ' 60-61
    Obsolute7 As Integer                           ' 62
    MultiWordDma As Integer                    ' 63
    PioMode As Integer                         ' 64
    MinMultiwordDmaCycleTime As Integer        ' 65
    RecommendedMultiwordDmaCycleTime As Integer   ' 66
    MinPioCycleTimewoFlowCtrl As Integer           ' 67
    MinPioCycleTimeWithFlowCtrl As Integer     ' 68
    Reserved2(5) As Integer                      ' 69-74
    QueueDepth As Integer                          ' 75
    SerialAtaCapabilities As Integer               ' 76
    ReservedForFutureSerialAta As Integer          ' 77
    SerialAtaFeaturesSupported As Integer          ' 78
    SerialAtaFeaturesEnabled As Integer        ' 79
    MajorVersion As Integer                    ' 80
    MinorVersion As Integer                    ' 81
    CommandSetSupported1 As Integer            ' 82
    CommandSetSupported2 As Integer            ' 83
    CommandSetSupported3 As Integer            ' 84
    CommandSetEnabled1 As Integer                  ' 85
    CommandSetEnabled2 As Integer                  ' 86
    CommandSetDefault As Integer                   ' 87
    UltraDmaMode As Integer                    ' 88
    TimeReqForSecurityErase As Integer         ' 89
    TimeReqForEnhancedSecure As Integer        ' 90
    CurrentPowerManagement As Integer              ' 91
    MasterPasswordRevision As Integer              ' 92
    HardwareResetResult As Integer             ' 93
    Acoustricmanagement As Integer             ' 94
    StreamMinRequestSize As Integer            ' 95
    StreamingTimeDma As Integer                ' 96
    StreamingAccessLatency As Integer              ' 97
    StreamingPerformance(1) As Integer               ' 98-99
    MaxUserLba As Long                         ' 100-103
    StremingTimePio As Integer                 ' 104
    Reserved3 As Integer                           ' 105
    SectorSize As Integer                          ' 106
    InterSeekDelay As Integer                      ' 107
    IeeeOui As Integer                         ' 108
    UniqueId3 As Integer                           ' 109
    UniqueId2 As Integer                           ' 110
    UniqueId1 As Integer                           ' 111
    Reserved4(3) As Integer                       ' 112-115
    Reserved5 As Integer                           ' 116
    WordsPerLogicalSector(1) As Integer              ' 117-118
    Reserved6(7) As Integer                       ' 119-126
    RemovableMediaStatus As Integer            ' 127
    SecurityStatus As Integer                      ' 128
    VendorSpecific(30) As Integer                 ' 129-159
    CfaPowerMode1 As Integer                       ' 160
    Reserved7(6) As Integer                       ' 161-167
    DeviceNominalFormFactor As Integer          '   168
    DataSetManagement As Integer                    '   169
    AdditionalProductIdentifier(3) As Integer '  170-173
    ReservedForCompactFlashAssociation(1) As Integer '  174-175
    CurrentMediaSerialNo(59) As Byte          ' 176-205
    SctCommandTransport As Integer             ' 206  254
    ReservedForCeAta1(1) As Integer               ' 207-208
    AlignmentOfLogicalBlocks As Integer        ' 209
    WriteReadVerifySectorCountMode3(1) As Integer   ' 210-211
    WriteReadVerifySectorCountMode2(1) As Integer   ' 212-213
    NvCacheCapabilities As Integer             ' 214
    NvCacheSizeLogicalBlocks(1) As Integer           ' 215-216
    NominalMediaRotationRate As Integer        ' 217
    Reserved8 As Integer                           ' 218
    NvCacheOptions1 As Integer                 ' 219
    NvCacheOptions2 As Integer                 ' 220
    Reserved9 As Integer                           ' 221
    TransportMajorVersionNumber As Integer     ' 222
    TransportMinorVersionNumber As Integer     ' 223
    ReservedForCeAta2(9) As Integer               ' 224-233
    MinimumBlocksPerDownloadMicrocode As Integer   ' 234
    MaximumBlocksPerDownloadMicrocode As Integer   ' 235
    Reserved10(18) As Integer                     ' 236-254
    IntegrityWord As Integer                       ' 255
End Type

Private Type SENDCMDOUTPARAMS
    cBufferSize As Long 'Size of Buffer in bytes
    DRIVERSTATUS As DRIVERSTATUS 'Driver status structure
    bBuffer() As Byte 'Buffer of arbitrary length for data read from drive
End Type

Private Const SMART_ENABLE_SMART_OPERATIONS = &HD8

Private Enum STATUS_FLAGS
    PRE_FAILURE_WARRANTY = &H1
    ON_LINE_COLLECTION = &H2
    PERFORMANCE_ATTRIBUTE = &H4
    ERROR_RATE_ATTRIBUTE = &H8
    EVENT_COUNT_ATTRIBUTE = &H10
    SELF_PRESERVING_ATTRIBUTE = &H20
End Enum

Private Const DFP_GET_VERSION = &H74080
Private Const DFP_SEND_DRIVE_COMMAND = &H7C084
Private Const DFP_RECEIVE_DRIVE_DATA = &H7C088

Private Type ATTR_DATA
    AttrID As Byte
    AttrName As String
    AttrValue As Byte
    ThresholdValue As Byte
    WorstValue As Byte
    StatusFlags As STATUS_FLAGS
End Type

Private Type DRIVE_INFO
    ID                      As Integer
    Index                   As Integer
    bDriveType              As Byte
    SerialNumber            As String
    Model                   As String
    InterfaceType           As String
    Family                  As String
    FirmWare                As String
    RealSize                As Double
    Size                    As Double
    NumAttributes           As Byte
    DeviceID                As String
    PNPDeviceId             As String
    Mode                    As String
    Attributes()            As ATTR_DATA
    Partitions()            As SoftwareHardDrivePartition
    Removable               As Boolean
End Type

Private Enum IDE_DRIVE_NUMBER
    PRIMARY_MASTER
    PRIMARY_SLAVE
    SECONDARY_MASTER
    SECONDARY_SLAVE
    TERTIARY_MASTER
    TERTIARY_SLAVE
    QUARTIARY_MASTER
    QUARTIARY_SLAVE
End Enum

Public Declare Function DeviceIoControl _
               Lib "KERNEL32" (ByVal hDevice As Long, _
                               ByVal dwIoControlCode As Long, _
                               lpInBuffer As Any, _
                               ByVal nInBufferSize As Long, _
                               lpOutBuffer As Any, _
                               ByVal nOutBufferSize As Long, _
                               lpBytesReturned As Long, _
                               lpOverlapped As Any) As Long

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx _
                Lib "KERNEL32" _
                Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Dim dFamilies As Dictionary
Dim arrHardDisks()  As HardwareHardDrive

Private Function GetDriveInfo(drvNumber As IDE_DRIVE_NUMBER) As DRIVE_INFO

    Dim di     As DRIVE_INFO
    Dim hDrive As Long

    hDrive = SmartOpen(drvNumber)

    di = GetWMIDriveInfo(drvNumber)

    If hDrive <> INVALID_HANDLE_VALUE Then
        If SmartGetVersion(hDrive) = True Then

            With di
                .bDriveType = 0
                .NumAttributes = 0
                ReDim .Attributes(0)
                .bDriveType = 1
            End With

            If SmartCheckEnabled(hDrive, drvNumber) Then
                If IdentifyDrive(hDrive, IDE_ID_FUNCTION, drvNumber, di) = True Then
                    GetDriveInfo = di
                End If 'IdentifyDrive
            End If 'SmartCheckEnabled
        End If 'SmartGetVersion
    End If 'hDrive <> INVALID_HANDLE_VALUE
    
    GetDriveInfo = di
    
    CloseHandle hDrive
End Function

Private Function IdentifyDrive(ByVal hDrive As Long, _
                               ByVal IDCmd As Byte, _
                               ByVal drvNumber As IDE_DRIVE_NUMBER, _
                               di As DRIVE_INFO) As Boolean

    Dim SCIP                          As SENDCMDINPARAMS
    Dim IDSEC                         As DeviceData ' IDSECTOR
    Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
    Dim cbBytesReturned               As Long

    With SCIP
        .cBufferSize = IDENTIFY_BUFFER_SIZE
        .bDriveNumber = CByte(drvNumber)

        With .irDriveRegs
            .bFeaturesReg = 0
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = 0
            .bCylHighReg = 0
            .bDriveHeadReg = &HA0 'compute the drive number

            If Not IsWinNT4Plus Then
                .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
            End If

            .bCommandReg = CByte(IDCmd)
        End With
    End With

    If DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, ByVal 0&) Then
        CopyMemory IDSEC, bArrOut(16), Len(IDSEC)
        di.Model = Trim$(StrConv(SwapBytes(IDSEC.ModelNumber), vbUnicode))
        di.SerialNumber = Trim$(StrConv(SwapBytes(IDSEC.SerialNumber), vbUnicode))
        di.Family = GetFamily(di.Model)
        If IDSEC.MaxUserLba > 0 Then
            di.Size = Round((IDSEC.MaxUserLba / 1000000000) * 512, 0)
            di.RealSize = Round((IDSEC.MaxUserLba / 1073741824) * 512, 2)
        End If
        
        If IDSEC.SerialAtaCapabilities <> 0 And IDSEC.SerialAtaCapabilities <> &HFFFF Then
            di.InterfaceType = "Serial ATA"
        Else
            di.InterfaceType = "Parallel ATA"
        End If

        If (IDSEC.UltraDmaMode And 40) Then
            di.Mode = "Ultra DMA 133 (Mode 6)"
        ElseIf (IDSEC.UltraDmaMode And 20) Then
            di.Mode = "Ultra DMA 100 (Mode 5)"
        ElseIf (IDSEC.UltraDmaMode And 10) Then
            di.Mode = "Ultra DMA 66 (Mode 4)"
        ElseIf (IDSEC.UltraDmaMode And 8) Then
            di.Mode = "Ultra DMA 44(Mode 3)"
        ElseIf (IDSEC.UltraDmaMode And 4) Then
            di.Mode = "Ultra DMA 33 (Mode 2)"
        ElseIf (IDSEC.UltraDmaMode And 2) Then
            di.Mode = "Ultra DMA 25 (Mode 1)"
        ElseIf (IDSEC.UltraDmaMode And 1) Then
            di.Mode = "Ultra DMA 16 (Mode 0)"
        End If

        If IDSEC.SerialAtaCapabilities And 8 Then
            di.Mode = "SATA 600"
        ElseIf IDSEC.SerialAtaCapabilities And 4 Then
            di.Mode = "SATA 300"
        ElseIf IDSEC.SerialAtaCapabilities And 2 Then
            di.Mode = "SATA 150"
        End If

        IdentifyDrive = True
    End If

End Function

Private Function IsWinNT4Plus() As Boolean

    Dim osv As OSVERSIONINFO

    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        IsWinNT4Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwVerMajor >= 4)
    End If

End Function

Private Function SmartCheckEnabled(ByVal hDrive As Long, _
                                   drvNumber As IDE_DRIVE_NUMBER) As Boolean

    Dim SCIP            As SENDCMDINPARAMS
    Dim SCOP            As SENDCMDOUTPARAMS
    Dim cbBytesReturned As Long

    With SCIP
        .cBufferSize = 0

        With .irDriveRegs
            .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI
            .bDriveHeadReg = &HA0

            If Not IsWinNT4Plus Then
                .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
            End If

            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
        End With

        .bDriveNumber = drvNumber
    End With

    SmartCheckEnabled = DeviceIoControl(hDrive, DFP_SEND_DRIVE_COMMAND, SCIP, Len(SCIP) - 4, SCOP, Len(SCOP) - 4, cbBytesReturned, ByVal 0&)
End Function

Private Function SmartGetVersion(ByVal hDrive As Long) As Boolean

    Dim cbBytesReturned As Long
    Dim GVOP            As GETVERSIONOUTPARAMS

    SmartGetVersion = DeviceIoControl(hDrive, DFP_GET_VERSION, ByVal 0&, 0, GVOP, Len(GVOP), cbBytesReturned, ByVal 0&)
End Function

Private Function SmartOpen(drvNumber As IDE_DRIVE_NUMBER) As Long

    If IsWinNT4Plus() Then
        SmartOpen = CreateFile("\\.\PhysicalDrive" & CStr(drvNumber), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    Else
        SmartOpen = CreateFile("\\.\SMARTVSD", 0&, 0&, ByVal 0&, CREATE_NEW, 0&, 0&)
    End If

End Function

Private Function SwapBytes(b() As Byte) As Byte()

    Dim bTemp As Byte
    Dim cnt   As Long

    For cnt = LBound(b) To UBound(b) Step 2
        bTemp = b(cnt)
        b(cnt) = b(cnt + 1)
        b(cnt + 1) = bTemp
    Next cnt

    SwapBytes = b()
End Function

Public Function EnumHardDrives() As HardwareHardDrive()

    Dim di        As DRIVE_INFO
    Dim drvNumber As Long

    Dim idx       As Integer

    idx = 0
    ReDim arrHardDisks(idx)
    
    For drvNumber = 0 To 10 'PRIMARY_MASTER ' To QUARTIARY_SLAVE-???????? ? ???? HDD
        
        di = GetDriveInfo(drvNumber)

        With di

            If .bDriveType = 1 And Len(.Model) > 0 Then
                ReDim Preserve arrHardDisks(idx)
                arrHardDisks(idx).Index = idx
                arrHardDisks(idx).DeviceID = .DeviceID
                arrHardDisks(idx).PNPDeviceId = .PNPDeviceId
                arrHardDisks(idx).SerialNumber = .SerialNumber  ' Serial Number
                arrHardDisks(idx).Model = .Model ' Drive Model
                arrHardDisks(idx).Family = .Family ' Drive Family
                arrHardDisks(idx).InterfaceType = .InterfaceType  ' Interface Type
                arrHardDisks(idx).RealSize = .RealSize ' Real Size
                arrHardDisks(idx).Size = .Size ' Theoritical Size
                arrHardDisks(idx).Mode = .Mode
                arrHardDisks(idx).Paritions = GetHardDiskPartitions(idx)
                arrHardDisks(idx).Removable = .Removable
                idx = idx + 1
            End If

        End With

    Next

    EnumHardDrives = arrHardDisks
End Function

Private Function GetFamily(ByVal Model As String) As String

    Dim sModel As Variant
    Dim re     As RegExp

    GetFamily = Model
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    Set dFamilies = LoadHardDiskFamilies

    For Each sModel In dFamilies.Keys
        re.Pattern = sModel

        If re.Test(Model) = True Then
            GetFamily = dFamilies.Item(UCase$(sModel))
            Exit For
        End If

    Next

End Function

Private Function LoadHardDiskFamilies() As Dictionary

    Dim d As New Dictionary

    With d
        .Add "^EXCELSTOR TECHNOLOGY J240$", "ExcelStor J240"
        .Add "^EXCELSTOR TECHNOLOGY J340$", "ExcelStor J340"
        .Add "^EXCELSTOR TECHNOLOGY J360$", "ExcelStor J360"
        .Add "^EXCELSTOR TECHNOLOGY J680$", "ExcelStor J680"
        .Add "^EXCELSTOR TECHNOLOGY J880$", "ExcelStor J880"
        .Add "(IBM-|HITACHI )?IC35L0[12346]0AVER07", "IBM Deskstar 60GXP"
        .Add "(IBM-)?DTLA-30[57]0[123467][05]", "IBM Deskstar 40GV & 75GXP"
        .Add "^FUJITSU M1623TAU$", "Fujitsu M1623TAU"
        .Add "^FUJITSU MHM2200AT", "Fujitsu MHM2200AT"
        .Add "^FUJITSU MHM2150AT", "Fujitsu MHM2150AT"
        .Add "^FUJITSU MHM2100AT", "Fujitsu MHM2100AT"
        .Add "^FUJITSU MHM2060AT", "Fujitsu MHM2060AT"
        .Add "^FUJITSU MHG2...ATU?", "Fujitsu MHG"
        .Add "^FUJITSU MHH2...ATU?", "Fujitsu MHH"
        .Add "^FUJITSU MHJ2...ATU?", "Fujitsu MHJ"
        .Add "^FUJITSU MHK2...ATU?", "Fujitsu MHK"
        .Add "^FUJITSU MHL2300AT$", "Fujitsu MHL2300AT"
        .Add "^FUJITSU MHN2...AT$", "Fujitsu MHN"
        .Add "^FUJITSU MHR2020AT$", "Fujitsu MHR2020AT"
        .Add "^FUJITSU MHR2040AT$", "Fujitsu MHR2040AT"
        .Add "^FUJITSU MHS20[6432]0AT(  .)?$", "Fujitsu MHSxxxxAT"
        .Add "^FUJITSU MHT2...(AH|AS|AT|BH)U?", "Fujitsu MHT"
        .Add "^FUJITSU MHU2...ATU?", "Fujitsu MHU"
        .Add "^FUJITSU MHV2...(AH|AS|AT|BH|BS|BT)", "Fujitsu MHV"
        .Add "^FUJITSU MP[A-G]3...A[HTEV]U?", "Fujitsu MPA..MPG"
        .Add "^FUJITSU MHW2(04|06|08|10|12|16)0BH$", "Fujitsu MHW2 BH"
        .Add "^SAMSUNG SV4012H$", "Samsung SV4012H"
        .Add "^SAMSUNG SV0412H$", "Samsung SV0412H"
        .Add "^SAMSUNG SV1204H$", "Samsung SV1204H"
        .Add "^SAMSUNG SV0322A$", "Samsung SV0322A"
        .Add "^SAMSUNG SP40A2H$", "Samsung SP40A2H with RR100-07 firmware"
        .Add "^SAMSUNG SP8004H$", "Samsung SP8004H with QW100-61 firmware"
        .Add "^SAMSUNG HD(250KD|(30[01]|320|40[01])L[DJ])$", "Samsung SpinPoint T133"
        .Add "^SAMSUNG HD(080G|160H|32[01]K|403L|50[01]L)J$", "Samsung SpinPoint T166"
        .Add "^SAMSUNG SP(16[01]3|2[05][01]4)[CN]$", "Samsung SpinPoint P120"
        .Add "^SAMSUNG HD(080H|120I|160J)J$", "Samsung SpinPoint P80 SD"
        .Add "^SAMSUNG SP(0451|08[0124]2|12[0145]3|16[0145]4)[CN]$", "Samsung SpinPoint P80"
        .Add "^SAMSUNG HD(753L|642J|501I|322H|25[12]H|103U|16[12]G)J$", "Samsung SpinPoint F1"
        .Add "^SAMSUNG HA(751L|501I|321H|251H|101U)J$", "Samsung SpinPoint F1 CE"
        .Add "^SAMSUNG HD(753LI|642JI|083GJ)$", "Samsung SpinPoint F1 DT"
        .Add "^SAMSUNG HE(753L|642J|502I|322H|252H|103U)J$", "Samsung SpinPoint F1 RAID Class"
        .Add "^SAMSUNG HD(754J|502H|103S)J$", "Samsung SpinPoint F3"
        .Add "^SAMSUNG HM(250JI|1(60|21)H[IC])$", "Samsung SpinPoint M5"
        .Add "^SAMSUNG HM([54]00L|(320|251)J)I$", "Samsung SpinPoint M6"
        .Add "^SAMSUNG HM([54]00J|320I|250H|161H|120G)I$", "Samsung SpinPoint M7"
        .Add "^SAMSUNG HM641JI$", "Samsung SpinPoint M7E"
        .Add "^SAMSUNG HM031HC$", "Samsung SpinPoint MC30"
        .Add "^SAMSUNG HM((08|12|10)[01]HJ)|((25|20|16)[01]JJ)$", "Samsung SpinPoint MP2"
        .Add "^SAMSUNG HM((64|50)0J|(32|25)0H)J$", "Samsung SpinPoint MP4"
        .Add "^SAMSUNG HM(320J|251J|160H)X$", "Samsung SpinPoint MU-X"
        .Add "^SAMSUNG HS((08[02]|06[0T]|04[0T])HB)|030GB|122JB$", "Samsung SpinPoint N2B"
        .Add "^SAMSUNG HS(0[68]|12)YHA$", "Samsung SpinPoint N3A"
        .Add "^SAMSUNG HS12UHE$", "Samsung SpinPoint N3B"
        .Add "^SAMSUNG HS(0[68]VH|1[26]VJ)F$", "Samsung SpinPoint N3C"
        .Add "^SAMSUNG HA(753L|642J|502J|322H|252H|103U)I$", "Samsung EcoGreen F1 CE "
        .Add "^SAMSUNG HD(753L|642J|502J|322H|252H|103U)I$", "Samsung EcoGreen F1 DT"
        .Add "^SAMSUNG HD(502HI|10[23]SI|15[34]UI)$", "Samsung EcoGreen F2"
        .Add "^SAMSUNG HD(754J|503H|203W|253G|105S|153W)I$", "Samsung EcoGreen F3"
        .Add "^MAXTOR 2B0(0[468]|1[05]|20)H1$", "Maxtor Fireball 541DX"
        .Add "^MAXTOR 2F0[234]0[JL]0$", "Maxtor Fireball 3"
        .Add "^MAXTOR 8(1750[AD]2|2560[AD]3|3240[AD]4|3500[AD]4|4320[AD]5|5250[AD]6|6480[AD]8|7000[AD]8)$", "Maxtor DiamondMax 1750 Ultra ATA"
        .Add "^MAXTOR 8(2160D2|3228D3|3240D3|4320D4|6480D6|8400D8|8455D8)$", "Maxtor DiamondMax 2160 Ultra ATA"
        .Add "^MAXTOR 9(0256D2|0288D2|0432D3|0510D4|0576D4|0648D5|0720D5|084[05]D6|0864D6|1008D[78]|1152D8)$", "Maxtor DiamondMax 2880 Ultra ATA"
        .Add "^MAXTOR 9(1(360|350|202)D8|1190D7|10[12]0D6|0840D5|06[48]0D4|0510D3|0340D2|1(350|202)E8|1010E6|0840E5|0640E4)$", "Maxtor DiamondMax 3400 Ultra ATA"
        .Add "^MAXTOR 6L0(20[JL]1|40[JL]2|60[JL]3|80[JL]4)$", "Maxtor DiamondMax Plus D740X"
        .Add "^MAXTOR 9(0500D4|0625D5|0750D6|0840D7|0875D7|0910D8|1000D8)$", "Maxtor DiamondMax Plus 2500 Ultra ATA"
        .Add "^MAXTOR 9(0512D2|0680D3|0750D3|0913D4|1024D4|1360D6|1536D6|1792D7|2048D8)$", "Maxtor DiamondMax Plus 5120 Ultra ATA 33"
        .Add "^MAXTOR 9(2732U8|2390U7|2049U6|1707U5|1366U4|1024U3|0845U3|0683U2)$", "Maxtor DiamondMax Plus 6800 Ultra ATA 66"
        .Add "^MAXTOR 9(2720U8|2040U6|1360U4|1020U3|0845U3|0650U2)$", "Maxtor DiamondMax 6800 Ultra ATA 66"
        .Add "^MAXTOR 4D0(20H1|40H2|60H3|80H4)$", "Maxtor DiamondMax D540X-4D"
        .Add "^MAXTOR (91728D8|91512D7|91303D6|91080D5|90845D4|90645D3|90648D[34]|90432D2)$", "Maxtor DiamondMax 4320 Ultra ATA"
        .Add "^MAXTOR 9(0431U1|0641U2|0871U2|1301U3|1741U4)$", "Maxtor DiamondMax 17 VL"
        .Add "^MAXTOR (94091U8|93071U6|92561U5|92041U4|917[34]1U4|91531U3|913[06]1U3|91021U2|908[47]1U2|90651U2|90431U1)$", "Maxtor DiamondMax 20 VL"
        .Add "^MAXTOR (33073U4|32049U3|31536U2|30768U1)$", "Maxtor DiamondMax 30 VL"
        .Add "^MAXTOR (93652U8|92739U6|91826U4|91369U3|90913U2|90845U2|90435U1)$", "Maxtor DiamondMax 36"
        .Add "^MAXTOR 9(0684U2|1024U2|136[29]U3|1536U3|1826U4|2049U4|2562U5|2739U6|3073U6|3652U8|4098U8)$", "Maxtor DiamondMax 40 ATA 66"
        .Add "^MAXTOR (54098[UH]8|53073[UH]6|52732[UH]6|52049[UH]4|51536[UH]3|51369[UH]3|51024[UH]2)$", "Maxtor DiamondMax Plus 40 (Ultra ATA 66 and Ultra ATA 100)"
        .Add "^MAXTOR 3(1024H1|1535H2|2049H2|3073H3|4098H4)( B)?$", "Maxtor DiamondMax 40 VL Ultra ATA 100"
        .Add "^MAXTOR 5(4610H6|4098H6|3073H4|2049H3|1536H2|1369H2|1023H2)$", "Maxtor DiamondMax Plus 45 Ulta ATA 100"
        .Add "^MAXTOR 9(1023U2|1536U2|2049U3|2305U3|3073U4|4610U6|6147U8)$", "Maxtor DiamondMax 60 ATA 66"
        .Add "^MAXTOR 9(1023H2|1536H2|2049H3|2305H3|3073H4|4610H6|6147H8)$", "Maxtor DiamondMax 60 ATA 100"
        .Add "^MAXTOR 5T0(60H6|40H4|30H3|20H2|10H1)$", "Maxtor DiamondMax Plus 60"
        .Add "^MAXTOR (98196H8|96147H6)$", "Maxtor DiamondMax 80"
        .Add "^MAXTOR 6[EN]040T0$", "Maxtor DiamondMax 8S"
        .Add "^MAXTOR 4W(100H6|080H6|060H4|040H3|030H2)$", "Maxtor DiamondMax 536DX"
        .Add "^MAXTOR 6(E0[234]|K04)0L0$", "Maxtor DiamondMax Plus 8"
        .Add "^MAXTOR 6(B(30|25|20|16|12|08)0[MPRS]|L(080[MLP]|(100|120)[MP]|160[MP]|200[MPRS]|250[RS]|300[RS]))0$", "Maxtor DiamondMax 10 (ATA/133 and SATA/150)"
        .Add "^MAXTOR 6V(080E|160E|200E|250F|300F|320F)0$", "Maxtor DiamondMax 10 (SATA/300)"
        .Add "^MAXTOR 6Y((060|080|120|160)L0|(080|120|160|200|250)P0|(060|080|120|160|200|250)M0)$", "Maxtor DiamondMax Plus 9"
        .Add "^MAXTOR 6H[45]00[FR]0$", "Maxtor DiamondMax 11"
        .Add "^MAXTOR 4(R0[468]0[JL]0|R1[26]0L0|A160J0|A250J0|A300J0|R120L4)$", "Maxtor DiamondMax 16"
        .Add "^MAXTOR 6G(080[LE]|160[PE]|250[PE]|320[PE])0$", "Maxtor DiamondMax 17"
        .Add "^MAXTOR STM3(402111|802110|160[28]12|200827|25062[34]|300622)A$", "Seagate Maxtor DiamondMax 20"
        .Add "^MAXTOR STM3(((402|80[28]|160[28])15)|250310|((2008|250[68]|3006|320[68])20)|250824|500630)AS?$", "Seagate Maxtor DiamondMax 21"
        .Add "^MAXTOR STM3(160813|320614|500[368]20|750[36]30|1000[36](34|40))AS$", "Seagate Maxtor DiamondMax 22"
        .Add "^MAXTOR STM3(16031|25031|32041|50041|75052|100052)8AS$", "Seagate Maxtor DiamondMax 23"
        .Add "^MAXTOR STM30(2503|5004|7504)OT[AB]3E1-RK$", "Seagate Maxtor OneTouch 4"
        .Add "^MAXTOR STM90(8|12|16)03OT[AB]3E1-RK$", "Seagate Maxtor OneTouch 4 Mini"
        .Add "^MAXTOR 7Y250[PM]0$", "Maxtor MaXLine Plus II"
        .Add "^MAXTOR [45]A(25|30|32)0[JN]0$", "Maxtor MaXLine II"
        .Add "^MAXTOR 7L(25|30)0[SR]0$", "Maxtor MaXLine III (ATA/133 and SATA 1)"
        .Add "^MAXTOR 7V(25|30)0F0$", "Maxtor MaXLine III (SATA 2)"
        .Add "^MAXTOR 7H500F0$", "Maxtor MaXLine Pro 500"
        .Add "^HITACHI_DK14FA-20B$", "Hitachi DK14FA-20B"
        .Add "^HITACHI_DK23..-..B?$", "Hitachi Travelstar DK23XX/DK23XXB"
        .Add "^(HITACHI_DK23FA-20J|HTA422020F9AT[JN]0)$", "Hitachi Endurastar J4K20/N4K20"
        .Add "^IBM-DTTA-3(7101|7129|7144|5032|5043|5064|5084|5101|5129|5168)0$", "IBM Deskstar 14GXP and 16GP"
        .Add "^IBM-DJNA-3(5(101|152|203|250)|7(091|135|180|220))0$", "IBM Deskstar 25GP and 22GXP"
        .Add "^IBM-DTCA-2(324|409)0$", "IBM Travelstar 4GT"
        .Add "^IBM-DARA-2(25|18|15|12|09|06)000$", "IBM Travelstar 25GS, 18GT, and 12GN"
        .Add "^(IBM-|Hitachi )?IC25(T048ATDA05|N0(30|20|15|12|10|07|06|05)ATDA04)-.$", "IBM Travelstar 48GH, 30GN, and 15GN"
        .Add "^IBM-DJSA-2(32|30|20|10|05)$", "IBM Travelstar 32GH, 30GT, and 20GN"
        .Add "^IBM-DKLA-2(216|324|432)0$", "IBM Travelstar 4GN"
        .Add "^IBM-DPTA-3(5(375|300|225|150)|7(342|273|205|136))0$", "IBM Deskstar 37GP and 34GXP"
        .Add "^(IBM-|Hitachi )?IC25(T060ATC[SX]05|N0[4321]0ATC[SX]04)-.$", "IBM/Hitachi Travelstar 60GH and 40GN"
        .Add "^(IBM-|Hitachi )?IC25N0[42]0ATC[SX]05-.$", "IBM/Hitachi Travelstar 40GNX"
        .Add "^(HITACHI )?IC25N0[23468]0ATMR04-.$", "Hitachi Travelstar 80GN"
        .Add "^(HITACHI )?HTS4240[234]0M9AT00$", "Hitachi Travelstar 4K40"
        .Add "^(HITACHI )?HTS5480[8642]0M9AT00$", "Hitachi Travelstar 5K80"
        .Add "^(HITACHI )?HTS5410[1864]0G9(AT|SA)00$", "Hitachi Travelstar 5K100"
        .Add "^(HITACHI )?HTE541040G9(AT|SA)00$", "Hitachi Travelstar E5K100"
        .Add "^(HITACHI )?HTS5412(60|80|10|12)H9(AT|SA)00$", "Hitachi Travelstar 5K120"
        .Add "^(HITACHI )?HTS5416([468]0|1[26])J9(AT|SA)00$", "Hitachi Travelstar 5K160"
        .Add "^(HITACHI )?HTS726060M9AT00$", "Hitachi Travelstar 7K60"
        .Add "^(HITACHI )?HTE7260[46]0M9AT00$", "Hitachi Travelstar E7K60"
        .Add "^(HITACHI )?HTS7210[168]0G9(AT|SA)00$", "Hitachi Travelstar 7K100"
        .Add "^(HITACHI )?HTE7210[168]0G9(AT|SA)00$", "Hitachi Travelstar E7K100"
        .Add "^(HITACHI )?HTS7220(80|10|12|16|20)K9(A3|SA)00$", "Hitachi Travelstar 7K200"
        .Add "^(IBM-)?IC35L((020|040|060|080|120)AVVA|0[24]0AVVN)07-[01]$", "IBM/Hitachi Deskstar 120GXP"
        .Add "^(IBM-)?IC35L(030|060|090|120|180)AVV207-[01]$", "IBM/Hitachi Deskstar GXP-180"
        .Add "^IBM-DCYA-214000$", "IBM Travelstar 14GS"
        .Add "^IBM-DTNA-2(180|216)0$", "IBM Travelstar 4LP"
        .Add "^(HITACHI )?HDS7280([48]0PLAT20|(40)?PLA320|80PLA380)$", "Hitachi Deskstar 7K80"
        .Add "^(HITACHI )?HDS7216(80|16)PLA[3T]80$", "Hitachi Deskstar 7K160"
        .Add "^(HITACHI )?HDS7225((40|80|12|16)VLAT20|(12|16|25)VLAT80|(80|12|16|25)VLSA80)$", "Hitachi Deskstar 7K250"
        .Add "^(HITACHI )?HDT7225((25|20|16)DLA(T80|380))$", "Hitachi Deskstar T7K250"
        .Add "^(HITACHI )?HDS724040KL(AT|SA)80$", "Hitachi Deskstar 7K400"
        .Add "^(HITACHI )?HDS725050KLA(360|T80)$", "Hitachi Deskstar 7K500"
        .Add "^(HITACHI )?HDT7250(25|32|40|50)VLA(360|380|T80)$", "Hitachi Deskstar T7K500"
        .Add "^(HITACHI )?HDS7210(75|10)KLA330$", "Hitachi Deskstar 7K1000"
        .Add "^(HITACHI )?HUA7220(20|10|50)CLA33(0|1)$", "Hitachi Ultrastar A7K2000"
        .Add "^(HITACHI )?HUS1560(30|45|60)VL(S|F)(4|6)0(0|1)$", "Hitachi Ultrastar 15K600"
        .Add "^(HITACHI )?HUS1545(30|45)VL(S|F)(3|4)00$", "Hitachi Ultrastar 15K450"
        .Add "^(HITACHI )?HUC1514(14|73)CSS60(0|1)$", "Hitachi Ultrastar C15K147"
        .Add "^(HITACHI )?HUC1030(14|30)CSS600$", "Hitachi Ultrastar C10K300"
        .Add "^(HITACHI )?HUA7210(50|75|10)KLA330$", "Hitachi Ultrastar 7K1000"
        .Add "^(HITACHI )?HDS7210(16|25|32|50|64|75|10|20)CLA3(3|6|8)(0|2)$", "Hitachi Ultrastar 7K1000.C"
        .Add "^(HITACHI )?HDS722020ALA330$", "Hitachi Ultrastar 7K2000"
        .Add "^(HITACHI )?HCS7210(16|25|32|50|75|10)CLA3(3|8)2$", "Hitachi CinemaStar 7K1000.C"
        .Add "^(HITACHI )?HCS5C(10(16|25|32|50|75|10)CLA382|3225SLA380)$", "Hitachi CinemaStar 5K1000"
        .Add "^(HITACHI )?HCC5450(16|25|32|50)B9A300$", "Hitachi CinemaStar C5K500"
        .Add "^(HITACHI )?HCC5432(16|25|32)A7A380$", "Hitachi CinemaStar Z5K320"
        .Add "^(HITACHI )?HTE7250(16|25|32|50)A9A364$", "Hitachi CinemaStar 7K500"
        .Add "^(HITACHI )?HTE5450(16|25|32|50)B9A300$", "Hitachi CinemaStar 5K500.B"
        .Add "^(HITACHI )?HTE7232(16|25|32)A7A364$", "Hitachi CinemaStar Z7K320"
        .Add "^(HITACHI )?HTE5432(16|25|32)A7A384$", "Hitachi CinemaStar Z5K320"
        .Add "^(HITACHI )?HEJ4210(40|80|10)G9(AT|SA)00$", "Hitachi EnduraStar J4K100"
        .Add "^(HITACHI )?HEN4210(40|80|10)G9(AT00|T00)$", "Hitachi EnduraStar N4K100"
        .Add "^TOSHIBA MK((6034|4032)GSX|(6034|4032)GAX|(6026|4026|4019|3019)GAXB?|(6025|6021|4025|4021|4018|3025|3021|3018)GAS|(4036|3029)GACE?|(4018|3017)GAP)$", "Toshiba 2.5'' HDD (30-60 GB)"
        .Add "^TOSHIBA MK(80(25GAS|26GAX|32GAX|32GSX)|10(31GAS|32GAX)|12(33GAS|34G[As]X)|2035GSS)$", "Toshiba 2.5'' HDD (80 GB And above)"
        .Add "^TOSHIBA MK[23468]00[4-9]GA[HL]$", "Toshiba 1.8'' HDD"
        .Add "^TOSHIBA MK6022GAX$", "Toshiba MK6022GAX"
        .Add "^TOSHIBA MK6409MAV$", "Toshiba MK6409MAV"
        .Add "^TOS MK3019GAXB SUN30G$", "Toshiba MK3019GAXB SUN30G"
        .Add "^TOSHIBA MK2016GAP$", "Toshiba MK2016GAP"
        .Add "^TOSHIBA MK2017GAP$", "Toshiba MK2017GAP"
        .Add "^TOSHIBA MK2018GAP$", "Toshiba MK2018GAP"
        .Add "^TOSHIBA MK2018GAS$", "Toshiba MK2018GAS"
        .Add "^TOSHIBA MK2023GAS$", "Toshiba MK2023GAS"
        .Add "^(MAXTOR )?8C(018|036|073)[JL]0$", "Seagate Atlas 15K"
        .Add "^(MAXTOR )?8[EK](036|073|147)[JL]0$", "Seagate Atlas 15K II"
        .Add "^(MAXTOR )?8K(036|073|147)S0$", "Seagate Atlas 15K II SAS"
        .Add "^(MAXTOR )?KU(018[JL]2|036[JL]4|073[JL]8)$", "Seagate Atlas 10K III"
        .Add "^(MAXTOR )?8B(036|073|146)[LJ]0$", "Seagate Atlas 10K IV"
        .Add "^(MAXTOR )?8[DJ](073|147|300)[LJ]0$", "Seagate Atlas 10K V"
        .Add "^ST9(20|28|40|48)11A$", "Seagate Momentus"
        .Add "^ST94811AB$", "Seagate Momentus Blade Server"
        .Add "^ST9(500562|320562|250561)0AS$", "Seagate Momentus XT"
        .Add "^ST9(16|25)03010AS$", "Seagate Momentus Thin"
        .Add "^ST9(2014|3015|4019)A$", "Seagate Momentus 42"
        .Add "^ST9(30219|402113)A$", "Seagate Momentus 42.2"
        .Add "^ST9(60821|808210|100822)A$", "Seagate Momentus 42.2"
        .Add "^ST9[24][08]11A$", "Seagate Momentus 54"
        .Add "^ST9(120824|10082[25]|8082(9|10)|608(12|21)|50214|402112|30218)A$", "Seagate Momentus 4200.2"
        .Add "^ST940110A$", "Seagate Momentus 4200 N-Lite"
        .Add "^ST9(50212|40211[23]|3021[89])A$", "Seagate Momentus 4200.2"
        .Add "^ST9(408116|960813|80811)A$", "Seagate Momentus 4200.3"
        .Add "^ST9(1[26]08220|808212)AS$", "Seagate Momentus 5400 PSD"
        .Add "^ST9((500327|250317)AS|(40811|60814|80821|100826|120826)A)$", "Seagate Momentus 5400 FDE"
        .Add "^ST9(80816|120827|160824)AS$", "Seagate Momentus 5400 FDE.2 SATA"
        .Add "^ST9(32032|1[26]031)[29]AS$", "Seagate Momentus 5400 FDE.3"
        .Add "^ST9(80314|(1[26]|25)0317|(32|50)0327)AS$", "Seagate Momentus 5400 FDE.4"
        .Add "^ST93012AM?$", "Seagate Momentus 5400.1"
        .Add "^ST9(808211|60822|408114|308110|120821|10082[34]|8823|6812|4813|3811)AS?$", "Seagate Momentus 5400.2"
        .Add "^ST9(4081[45]|6081[35]|8081[15]|100828|120822|160821)A$", "Seagate Momentus 5400.3"
        .Add "^ST9(4081[45]|6081[35]|8081[15]|100828|120822|160821)AB$", "Seagate Momentus 5400.3 ED"
        .Add "^ST9((25|20|16)0827|120817)AS$", "Seagate Momentus 5400.4"
        .Add "^ST9((8|16)0310|320320)ASG?$", "Seagate Momentus 5400.5"
        .Add "^ST9((50|32)0325|250315|160314|120315)ASG?$", "Seagate Momentus 5400.6"
        .Add "^ST9(160316|250310|32031[02]|50032[01]|64032[02])AS$", "Seagate Momentus 5400.7"
        .Add "^ST9(500421|250411)AS$", "Seagate Momentus 7200 FDE"
        .Add "^ST9(10021|80825|6023|4015)AS?$", "Seagate Momentus 7200.1"
        .Add "^ST9(80813|100821|120823|160823|200420)ASG?$", "Seagate Momentus 7200.2"
        .Add "^ST9(((8|12|16)0411)|((20|25|32)0421))ASG?$", "Seagate Momentus 7200.3"
        .Add "^ST9(120410|160412|250410|320423|500420)ASG?$", "Seagate Momentus 7200.4"
        .Add "^ST3(10014A(CE)?|20014A)$", "Seagate UX"
        .Add "^ST3(30012|40012|60012|80022|120020)A$", "Seagate U7"
        .Add "^ST3(8002|6002|4081|3061|2041)0A$", "Seagate U6"
        .Add "^ST3(40823|30621|20413|15311|10211)A$", "Seagate U5"
        .Add "^ST3(2112|4311|6421|8421)A$", "Seagate U4"
        .Add "^ST3(8410|4313|17221|13021)A$", "Seagate U8"
        .Add "^ST3(20423|15323|10212)A$", "Seagate U10"
        .Add "^ST950043[01]SS$", "Seagate Constellation SAS"
        .Add "^ST9(500530|160511)NS$", "Seagate Constellation SATA"
        .Add "^ST3(5|10|20)00(4[124][45]SS|5[124]4NS)$", "Seagate Constellation ES"
        .Add "^ST3(300[56]|450[78]|600[09])57(FC|SS)$", "Seagate Cheetah 15K.7"
        .Add "^ST3(4508|3006|1463)56(FC|SS)$", "Seagate Cheetah 15K.6"
        .Add "^ST3(3006|1468|734)55(LW|LC|FC|SS)$", "Seagate Cheetah 15K.5"
        .Add "^ST3(1468|734|367)54(LW|LC|FC|SS)$", "Seagate Cheetah 15K.4"
        .Add "^ST3(734|367|184)53(LW|LC|FC)$", "Seagate Cheetah 15K.3"
        .Add "^ST3(3000|1467|732|368)07(LW|LC|FC)$", "Seagate Cheetah 10K.7"
        .Add "^ST3(1468|733|366)07(LW|LC|FC)$", "Seagate Cheetah 10K.6"
        .Add "^ST3(184|92)51(LW|LC)$", "Seagate Cheetah X15"
        .Add "^ST3(367|184)[53]2(LW|LC|FC)$", "Seagate Cheetah X15-36LP"
        .Add "^ST3(733|1467|3005)55SS$", "Seagate Cheetah T10 SAS"
        .Add "^ST3(3009|4007)55(FC|SS)$", "Seagate Cheetah NS"
        .Add "^ST19101(N|W|WC|WD|DC)$", "Seagate Cheetah 9"
        .Add "^ST3(91|45)02L[WC]$", "Seagate Cheetah 9LP"
        .Add "^ST3(734|366)05(LC|LCV|LW|FC)$", "Seagate Cheetah 73LP"
        .Add "^ST173404(LC|LCV|LWV|LW|FC|FCV)$", "Seagate Cheetah 73"
        .Add "^ST136403(LC|LCV|LWV|LW|FC|FCV)$", "Seagate Cheetah 36"
        .Add "^ST336704(LC|LCV|LWV|LW|FC|FCV)$", "Seagate Cheetah 36LP"
        .Add "^ST3(184|92)05L[WC]$", "Seagate Cheetah 36XL"
        .Add "^ST3(367|184)[04]6L[WC]$", "Seagate Cheetah 36ES"
        .Add "^ST118202L[WC]$", "Seagate Cheetah 18"
        .Add "^ST3(92|184)04L[WC]$", "Seagate Cheetah 18XL"
        .Add "^ST3(91|182)[03]3(LW|LC|FC|LWV|LCV|FCV)$", "Seagate Cheetah 18LP"
        .Add "^ST34501(N|W|WC|WD|DC)$", "Seagate Cheetah 4LP"
        .Add "^ST3250310CS$", "Seagate DB35.4"
        .Add "^ST3(8|16|25|32|40|50|75)0(640|8[234]0|215)[SA]CE$", "Seagate DB35.3"
        .Add "^ST3(8|12|16|20|25|30|40|50)0(8(41|33|2[247])|2110|21[23])[SA]CE$", "Seagate DB32 7200.2"
        .Add "^ST3(20|25|30|40)08(2[36]|3[12])ACE$", "Seagate DB35 (IDE)"
        .Add "^ST3(20|25|30|40)08(2[36]|3[12])SCE$", "Seagate DB35 (SATA)"
        .Add "^ST9[48]081[78][SA]M$", "Seagate EE25.2"
        .Add "^ST9[234]081[34]AM$", "Seagate EE25.1"
        .Add "^ST31000533CS$", "Seagate Pipeline HD Pro"
        .Add "^ST9(16|25|32|50)03[12](1|3|8|10)CS$", "Seagate Pipeline HD Mini SATA"
        .Add "^ST3(16|25|32|50|100)03[12][0126]CS$", "Seagate Pipeline HD"
        .Add "^ST9((146[78]52)|73[34]52)SS$", "Seagate Savvio 15K.2"
        .Add "^ST9(367|734)51FC$", "Seagate Savvio 15K.1 FC"
        .Add "^ST9(367|734)51SS$", "Seagate Savvio 15K"
        .Add "^ST9(600[12]|450[34])04(SS|FC)$", "Seagate Savvio 10K.4"
        .Add "^ST9(300[65]|146[78])03SS$", "Seagate Savvio 10K.3"
        .Add "^ST9(1468|734)02SS$", "Seagate Savvio 10K.2"
        .Add "^ST9(734|367)01(LC|SS)$", "Seagate Savvio 10K"
        .Add "^ST3(1000525|500410|250311)SV$", "Seagate SV35.5"
        .Add "^ST3320410SV$", "Seagate SV35.4"
        .Add "^ST3(100034|75033|50032|25031)0SV$", "Seagate SV35.3"
        .Add "^ST3(750640|500630|320620|250820|160815)[AS]V$", "Seagate SV35.2"
        .Add "^ST3(500641|250824|160812)[AS]V$", "Seagate SV35"
        .Add "^ST3(2804|2043|1362|1022|6810)0A$", "Seagate Barracuda ATA"
        .Add "^ST3(3063|2042|1532|1021)0A$", "Seagate Barracuda ATA II"
        .Add "^ST3(30631|20424|15324|10216)A$", "Seagate Barracuda ATA II 100"
        .Add "^ST3(40824|30620|20414|15310|10215)A$", "Seagate Barracuda ATA III"
        .Add "^ST3(20011|30011|40016|60021|80021)A$", "Seagate Barracuda ATA IV"
        .Add "^ST3(12002(3A|4A|9A|3AS)|800(23A|15A|23AS)|60(015A|210A)|40017A)$", "Seagate Barracuda ATA V"
        .Add "^ST340015A$", "Seagate Barracuda 5400.1"
        .Add "^ST3(200021A|200822AS?|16002[13]AS?|12002[26]AS?|1[26]082[78]AS|8001[13]AS?|80817AS|60014A|40111AS|40014AS?)$", "Seagate Barracuda 7200.7 and 7200.7 Plus"
        .Add "^ST3(400[68]32|300[68]31|250[68]23|200826)AS?$", "Seagate Barracuda 7200.8"
        .Add "^ST3(402111?|80[28]110?|120[28]1[0134]|160[28]1[012]|200827|250[68]24|300[68]22|(320|400)[68]33|500[68](32|41))AS?$", "Seagate Barracuda 7200.9"
        .Add "^ST3((80|160)[28]15|200820|250[34]10|(250|300|320|400)[68]20|500[68]30|750[68]40)AS?$", "Seagate Barracuda 7200.10"
        .Add "^ST3(500[368]2|750[36]3|1000[36]4)0AS$", "Seagate Barracuda 7200.11"
        .Add "^ST3(1500341|1000333|640[36]23|320[68]13|160813)AS$", "Seagate Barracuda 7200.11"
        .Add "^ST3(160318|250318|320418|50041[08]|750528|1000528)AS$", "Seagate Barracuda 7200.12"
        .Add "^ST3(10|15|20)005(20|41|42)|(500412)AS?$", "Seagate Barracuda LP"
        .Add "^ST3(91|45)73(N|W|WD|LW|WC|LC)$", "Seagate Barracuda 9LP"
        .Add "^ST19171(N|W|WD|WC|DC)$", "Seagate Barracuda 9"
        .Add "^ST15150(N|W|ND|WD|WC|DC)$", "Seagate Barracuda 4"
        .Add "^ST3(22|45)72(N|W|WD|WC|DC)$", "Seagate Barracuda 4XL"
        .Add "^ST3(22|45|21|43)71(N|W|WD|WC|DC)$", "Seagate Barracuda 4LP"
        .Add "^ST3(12|25)50(N|ND|W|WD|WC|DC)$", "Seagate Barracuda 2LP"
        .Add "^ST12450(W|WD)$", "Seagate Barracuda 2 / 2HP"
        .Add "^ST118273(N|W|WD|LW|WC|LC)$", "Seagate Barracuda 18"
        .Add "^ST11950(N|ND|W|WD)$", "Seagate Barracuda 1"
        .Add "^ST12550(N|ND|W|WD)$", "Seagate Barracuda 2"
        .Add "^ST3(91|182)75L[WC]$", "Seagate Barracuda 18LP"
        .Add "^ST3(92|184)[123]6(N|W|LW|LC|LCV|LWV)$", "Seagate Barracuda 18XL"
        .Add "^ST150176L[WC]$", "Seagate Barracuda 50"
        .Add "^ST3(250[68]2|320[68]2|400[68]2|500[68]3|750[68]4)0NS$", "Seagate Barracuda ES"
        .Add "^ST3(25|32|40|50|75)0[68][234]1NS$", "Seagate Barracuda ES+"
        .Add "^ST3((500620|750630|1000640)S)|((250310|250610|500320|750330|1000340)N)S$", "Seagate Barracuda ES.2"
        .Add "^ST136475L[WC]$", "Seagate Barracuda 36"
        .Add "^ST3(3673|184[13])7(N|W|LW|LC)$", "Seagate Barracuda 36ES"
        .Add "^ST3(369|184)(18N|38LW)$", "Seagate Barracuda 36ES2"
        .Add "^ST1181677(LC|LW|FC)V$", "Seagate Barracuda 180"
        .Add "^ST32000641AS$", "Seagate Barracuda XT"
        .Add "^ST317240A$", "Seagate Medalist 17240"
        .Add "^ST313030A$", "Seagate Medalist 13030"
        .Add "^ST310231A$", "Seagate Medalist 10231"
        .Add "^ST38420A$", "Seagate Medalist 8420"
        .Add "^ST34310A$", "Seagate Medalist 4310"
        .Add "^ST317242A$", "Seagate Medalist 17242"
        .Add "^ST313032A$", "Seagate Medalist 13032"
        .Add "^ST310232A$", "Seagate Medalist 10232"
        .Add "^ST38422A$", "Seagate Medalist 8422"
        .Add "^ST34312A$", "Seagate Medalist 4312"
        .Add "^ST32110A$", "Seagate Medalist 2110"
        .Add "^ST33221A$", "Seagate Medalist 3221"
        .Add "^ST34321A$", "Seagate Medalist 4321"
        .Add "^ST36531A$", "Seagate Medalist 6531"
        .Add "^ST38641A$", "Seagate Medalist 8641"
        .Add "^ST3(250623|250823|400632|400832|250824|250624|400633|400833|500641|500841)NS$", "Seagate NL35"
        .Add "^ST3((8|16)0215|(25|32|40)0820|500830|750[68]40)[AS]CE$", "Seagate DB35.3"
        .Add "^WDC WD([2468]00E|1[26]00A)B-.*$", "Western Digital Protege"
        .Add "^WDC WD(2|3|4|6|8|10|12|16|18|20|25)00BB-.*$", "Western Digital Caviar"
        .Add "^WDC WD((((50|64|75)01A)|1001F)ALS)|2001FASS|1002FAEX$", "Western Digital Caviar Black"
        .Add "^WDC WD((50|64|75)00AAC|(64|80)00AAR|(10|15|20)EAR|(50|64|75)00AAD|(10|15|20)EAD|(50|64|75)00AAV|10EAV)S$", "Western Digital Caviar Green"
        .Add "^WDC WD10EALS$", "Western Digital Caviar Blue SATA with 32MB cache"
        .Add "^WDC WD(25|32|40|50|64|75)00AAKS$", "Western Digital Caviar Blue SATA with 16MB cache"
        .Add "^WDC WD(8|12|16|25|32|40|50)00AAJS$", "Western Digital Caviar Blue SATA with 8MB cache"
        .Add "^WDC WD(400JB|800JD)", "Western Digital Caviar Blue SATA with 8MB cache"
        .Add "^WDC WD(800BD|1600AABS)$", "Western Digital Caviar Blue SATA with 2MB cache"
        .Add "^WDC WD(32|40|50)00AAKB$", "Western Digital Caviar Blue EIDE with 16MB cache"
        .Add "^WDC WD(8|16|25|32|40|50)00AAJB$", "Western Digital Caviar Blue EIDE with 8MB cache"
        .Add "^WDC WD800JB$", "Western Digital Caviar Blue EIDE with 8MB cache"
        .Add "^WDC WD1600AABB$", "Western Digital Caviar Blue EIDE with 2MB cache"
        .Add "^WDC WD([48]00BB)$", "Western Digital Caviar Blue EIDE with 2MB cache"
        .Add "^WDC WD(16|25|32|50)00AVJ[BS]$", "Western Digital AV"
        .Add "^WDC WD1600AVBS$", "Western Digital AV"
        .Add "^WDC WD((16|25|32|50|64|75)00A|10E)VV|(50|75)00AVD|((10|15|20)EVD)|(10EUR)S$", "Western Digital AV-GP"
        .Add "^WDC WD(16|25|32|50)BUDT$", "Western Digital AV-25"
        .Add "^WDC AC12500.?", "Western Digital Caviar AC12500"
        .Add "^WDC AC14300.?", "Western Digital Caviar AC14300"
        .Add "^WDC AC23200.?", "Western Digital Caviar AC23200"
        .Add "^WDC AC24300.?", "Western Digital Caviar AC24300"
        .Add "^WDC AC25100.?", "Western Digital Caviar AC25100"
        .Add "^WDC AC36400.?", "Western Digital Caviar AC36400"
        .Add "^WDC AC38400.?", "Western Digital Caviar AC38400"
        .Add "^WDC WD(3|4|6)00AB-.*$", "Western Digital Caviar WDxxxAB"
        .Add "^WDC WD...?AA(-.*)?$", "Western Digital Caviar WDxxxAA"
        .Add "^WDC WD...BA$", "Western Digital Caviar WDxxxBA"
        .Add "^WDC WD(4|8|20|32)00BD-.*$", "Western Digital Caviar Serial ATA"
        .Add "^WDC WD((4|6|8|10|12|16|18|20|25|30|32|40|50)00(JB|PB|AAJB|AAKB))-.*$", "Western Digital Caviar SE"
        .Add "^WDC WD((4|8|12|16|20|25|32|40)00(JD|KD))-.*$", "Western Digital Caviar SE Serial ATA"
        .Add "^WDC WD((8|12|16|20|25|30|32|40|50|75)00(JS|KS|AABS|AAJS|AAKS))-.*$", "Western Digital Caviar Second Generation Serial ATA"
        .Add "^WDC WD((12|16|25|32|40|50|75)00(SD|YD|YR|YS|ABYS|AYYS))-.*$", "Western Digital Caviar RE Serial ATA"
        .Add "^WDC WD((12|16|25|32)00SB)-.*$", "Western Digital Caviar RE EIDE"
        .Add "^WDC WD(7500A|1000F)YPS$", "Western Digital RE2-GP"
        .Add "^WDC WD((25|32|50|75)02A|1002F)BYS$", "Western Digital RE3"
        .Add "^WDC WD(25|50|10|15|20)03([AF]BYX|FYYS)$", "Western Digital RE4"
        .Add "^WDC WD2002FYPS$", "Western Digital RE4-GP"
        .Add "^WDC WD((360|740|800)GD|(360|740|1500)ADFD)-.*$", "Western Digital Raptor"
        .Add "^WDC WD(15|30|45|60)00[BH]L(FS|HX)$", "Western Digital VelociRaptor"
        .Add "^WDC WD((4|6|8|10|12|16|20|25)00(UE|VE|BEAS|BEVE|BEVS))-.*$", "Western Digital Scorpio"
        .Add "^WDC WD(8|12|16|25|32|50)00B[EJ]KT$", "Western Digital Scorpio Black"
        .Add "^WDC WD(8|12|16|25|32|40|50|64|75)00B(EV[TSE]|PVT)$", "Western Digital Scorpio Blue"
        .Add "^WDC SSC-D0(064|128|256)SC-2100$", "Western Digital SiliconEdge"
        .Add "^QUANTUM BIGFOOT TS10.0A$", "Quantum BIGFOOT TS10.0A"
        .Add "^QUANTUM FIREBALLlct15 [123]0$", "Quantum FIREBALLlct15 20 And QUANTUM FIREBALLlct15 30"
        .Add "^QUANTUM FIREBALLlct20 [234]0$", "Quantum FIREBALLlct20"
        .Add "^QUANTUM FIREBALL CX10.2A$", "Quantum FIREBALL CX10.2A"
        .Add "^QUANTUM FIREBALLP LM(10.2|15|20.5|30)$", "Quantum Fireball Plus LM"
        .Add "^QUANTUM FIREBALL CR(4.3|8.4)A$", "Quantum Fireball CR"
        .Add "^QUANTUM FIREBALLP AS(10.2|20.5|30.0|40.0)$", "Quantum FIREBALLP AS10.2, AS20.5 And AS40.0"
        .Add "^QUANTUM FIREBALL EX6.4A$", "Quantum FIREBALL EX6.4A"
        .Add "^QUANTUM FIREBALL ST(3.2|4.3)A$", "Quantum FIREBALL ST3.2A"
        .Add "^QUANTUM FIREBALL EX3.2A$", "Quantum FIREBALL EX3.2A"
        .Add "^QUANTUM FIREBALLP KX27.3$", "Quantum FIREBALLP KX27.3"
        .Add "^QUANTUM FIREBALLP KA(9|10).1$", "Quantum Fireball Plus KA"
        .Add "^QUANTUM FIREBALL SE4.3A$", "Quantum Fireball SE"
    End With

    Set LoadHardDiskFamilies = d
End Function

Private Function GetWMIDriveInfo(drvNumber As IDE_DRIVE_NUMBER) As DRIVE_INFO
On Error Resume Next

    Dim objWMIService, colItems, objItem
    Dim di     As DRIVE_INFO

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive WHERE Index = '" & drvNumber & "'")  ' AND MediaType LIKE '%fixed%'", , 48)


    For Each objItem In colItems

        With di
            .bDriveType = 1
            .Index = drvNumber
            .Model = GetProp(objItem.Caption)
            .Family = GetFamily(.Model)
            .InterfaceType = GetProp(objItem.InterfaceType)
            If GetProp(objItem.Size) > 0 Then
                .RealSize = Round(GetProp(objItem.Size) / 1073741824, 2)
                .Size = Round(GetProp(objItem.Size) / 1000000000, 0)
            End If
            .DeviceID = GetProp(objItem.DeviceID)
            .PNPDeviceId = GetProp(objItem.PNPDeviceId)
            .Removable = InStr(LCase$(GetProp(objItem.MediaType)), "fixed") = 0
        End With

    Next

    GetWMIDriveInfo = di

    Set objWMIService = Nothing
    Set colItems = Nothing
End Function

Public Function GetHardDiskPartitions(drvNumber As Integer) As SoftwareHardDrivePartition()
On Error Resume Next

    Dim wmiService, wmiServices ', colItems, objItem
    Dim wmiDiskDrive, wmiDiskDrives
    Dim wmiDiskPartition, wmiDiskPartitions
    Dim wmiLogicalDisk, wmiLogicalDisks
    Dim tmpArr() As SoftwareHardDrivePartition
    Dim idx      As Integer
    Dim query    As String

    idx = 0
    ReDim tmpArr(0)
    Set wmiServices = GetObject("winmgmts:\\.\root\CIMV2")
    Set wmiDiskDrives = wmiServices.ExecQuery("SELECT * FROM Win32_DiskDrive WHERE Index='" & drvNumber & "'")

    For Each wmiDiskDrive In wmiDiskDrives
        query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" & wmiDiskDrive.DeviceID & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
        Set wmiDiskPartitions = wmiServices.ExecQuery(query)

        For Each wmiDiskPartition In wmiDiskPartitions
            Set wmiLogicalDisks = wmiServices.ExecQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" & wmiDiskPartition.DeviceID & "'} WHERE AssocClass = Win32_LogicalDiskToPartition")

            For Each wmiLogicalDisk In wmiLogicalDisks
                ReDim Preserve tmpArr(idx)

                With tmpArr(idx)
                    .Caption = wmiLogicalDisk.Caption
                    .FileSystem = GetProp(wmiLogicalDisk.FileSystem)

                    If Not IsNull(wmiLogicalDisk.FreeSpace) Then
                        .FreeSpace = wmiLogicalDisk.FreeSpace
                    End If

                    If Not IsNull(wmiLogicalDisk.Size) Then
                        .Size = wmiLogicalDisk.Size
                    End If

                    .VolumeName = GetProp(wmiLogicalDisk.VolumeName)
                End With

                idx = idx + 1
            Next
        Next
    Next

    Set wmiServices = Nothing
    Set wmiDiskDrives = Nothing
    Set wmiDiskPartitions = Nothing
    Set wmiLogicalDisks = Nothing
    
    GetHardDiskPartitions = tmpArr
End Function
