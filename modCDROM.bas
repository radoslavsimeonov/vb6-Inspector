Attribute VB_Name = "hwCDROM"
Option Explicit

Private Declare Function GetLogicalDriveStrings _
                Lib "KERNEL32" _
                Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                 ByVal lpBuffer As String) As Long

Private Declare Function GetDriveType _
                Lib "KERNEL32" _
                Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Type SCSI_PASS_THROUGH
    Length As Integer
    ScsiStatus As Byte
    PathId As Byte
    TargetId As Byte
    Lun As Byte
    CdbLength As Byte
    SenseInfoLength As Byte
    DataIn As Byte
    DataTransferLength As Long
    TimeOutValue As Long
    DataBufferOffset As Long
    SenseInfoOffset As Long
    Cdb(15) As Byte
End Type

Private Type DEVICE_MEDIA_INFO
    Cylinders As Double
    MediaType As Long
    TracksPerCylinder As Long
    SectorsPerTrack As Long
    BytesPerSector As Long
    NumberMediaSides As Long
    MediaCharacteristics As Long
End Type

Private Type GET_MEDIA_TYPES
    DeviceType As Long
    MediaInfoCount As Long
    MediaInfo(10) As DEVICE_MEDIA_INFO
End Type

Private Type SCSI_PASS_THROUGH_WITH_BUFFERS
    SPT As SCSI_PASS_THROUGH
    SenseBuf(31) As Byte
    DataBuf(511) As Byte
End Type

Private Const FILE_DEVICE_CD_ROM As Long = &H2
Private Const FILE_DEVICE_DVD    As Long = &H33
Private Const DRIVE_REMOVABLE    As Long = 2
Private Const DRIVE_FIXED        As Long = 3
Private Const DRIVE_REMOTE       As Long = 4
Private Const DRIVE_CDROM        As Long = 5
Private Const DRIVE_RAMDISK      As Long = 6

Private Enum SCSI_PASS_THROUGH_CDBMODE
    MODE_PAGE_ERROR_RECOVERY = &H1
    MODE_PAGE_DISCONNECT = &H2
    MODE_PAGE_FORMAT_DEVICE = &H3
    MODE_PAGE_RIGID_GEOMETRY = &H4
    MODE_PAGE_FLEXIBILE = &H5
    MODE_PAGE_VERIFY_ERROR = &H7
    MODE_PAGE_CACHING = &H8
    MODE_PAGE_PERIPHERAL = &H9
    MODE_PAGE_CONTROL = &HA
    MODE_PAGE_MEDIUM_TYPES = &HB
    MODE_PAGE_NOTCH_PARTITION = &HC
    MODE_SENSE_RETURN_ALL = &H3F
    MODE_SENSE_CURRENT_VALUES = &H0
    MODE_SENSE_CHANGEABLE_VALUES = &H40
    MODE_SENSE_DEFAULT_VAULES = &H80
    MODE_SENSE_SAVED_VALUES = &HC0
    MODE_PAGE_DEVICE_CONFIG = &H10
    MODE_PAGE_MEDIUM_PARTITION = &H11
    MODE_PAGE_DATA_COMPRESS = &HF
    MODE_PAGE_CAPABILITIES = &H2A
End Enum

Private Enum SCSI_PASS_THROUGH_CDBLENGTH
    CDB6GENERIC_LENGTH = 6
    CDB10GENERIC_LENGTH = 10
    CDB12GENERIC_LENGTH = 12
End Enum

Private Enum SCSI_PASS_THROUGH_DATAIN
    SCSI_IOCTL_DATA_OUT = 0
    SCSI_IOCTL_DATA_IN = 1
    SCSI_IOCTL_DATA_UNSPECIFIED = 2
End Enum

Private Enum SCSI_PASS_THROUGH_OPERATION
    SCSIOP_TEST_UNIT_READY = &H0
    SCSIOP_REZERO_UNIT = &H1
    SCSIOP_REWIND = &H1
    SCSIOP_REQUEST_BLOCK_ADDR = &H2
    SCSIOP_REQUEST_SENSE = &H3
    SCSIOP_FORMAT_UNIT = &H4
    SCSIOP_READ_BLOCK_LIMITS = &H5
    SCSIOP_REASSIGN_BLOCKS = &H7
    SCSIOP_READ6 = &H8
    SCSIOP_RECEIVE = &H8
    SCSIOP_WRITE6 = &HA
    SCSIOP_PRINT = &HA
    SCSIOP_SEND = &HA
    SCSIOP_SEEK6 = &HB
    SCSIOP_TRACK_SELECT = &HB
    SCSIOP_SLEW_PRINT = &HB
    SCSIOP_SEEK_BLOCK = &HC
    SCSIOP_PARTITION = &HD
    SCSIOP_READ_REVERSE = &HF
    SCSIOP_WRITE_FILEMARKS = &H10
    SCSIOP_FLUSH_BUFFER = &H10
    SCSIOP_SPACE = &H11
    SCSIOP_INQUIRY = &H12
    SCSIOP_VERIFY6 = &H13
    SCSIOP_RECOVER_BUF_DATA = &H14
    SCSIOP_MODE_SELECT = &H15
    SCSIOP_RESERVE_UNIT = &H16
    SCSIOP_RELEASE_UNIT = &H17
    SCSIOP_COPY = &H18
    SCSIOP_ERASE = &H19
    SCSIOP_MODE_SENSE = &H1A
    SCSIOP_START_STOP_UNIT = &H1B
    SCSIOP_STOP_PRINT = &H1B
    SCSIOP_LOAD_UNLOAD = &H1B
    SCSIOP_RECEIVE_DIAGNOSTIC = &H1C
    SCSIOP_SEND_DIAGNOSTIC = &H1D
    SCSIOP_MEDIUM_REMOVAL = &H1E
    SCSIOP_READ_CAPACITY = &H25
    SCSIOP_READ = &H28
    SCSIOP_WRITE = &H2A
    SCSIOP_SEEK = &H2B
    SCSIOP_LOCATE = &H2B
    SCSIOP_WRITE_VERIFY = &H2E
    SCSIOP_VERIFY = &H2F
    SCSIOP_SEARCH_DATA_HIGH = &H30
    SCSIOP_SEARCH_DATA_EQUAL = &H31
    SCSIOP_SEARCH_DATA_LOW = &H32
    SCSIOP_SET_LIMITS = &H33
    SCSIOP_READ_POSITION = &H34
    SCSIOP_SYNCHRONIZE_CACHE = &H35
    SCSIOP_COMPARE = &H39
    SCSIOP_COPY_COMPARE = &H3A
    SCSIOP_WRITE_DATA_BUFF = &H3B
    SCSIOP_READ_DATA_BUFF = &H3C
    SCSIOP_CHANGE_DEFINITION = &H40
    SCSIOP_READ_SUB_CHANNEL = &H42
    SCSIOP_READ_TOC = &H43
    SCSIOP_READ_HEADER = &H44
    SCSIOP_PLAY_AUDIO = &H45
    SCSIOP_PLAY_AUDIO_MSF = &H47
    SCSIOP_PLAY_TRACK_INDEX = &H48
    SCSIOP_PLAY_TRACK_RELATIVE = &H49
    SCSIOP_PAUSE_RESUME = &H4B
    SCSIOP_LOG_SELECT = &H4C
    SCSIOP_LOG_SENSE = &H4D
End Enum

Dim CDROMs() As HardwareCDRomDrive
Dim idx      As Integer

Public Function EnumCDROMs() As HardwareCDRomDrive()

    Dim DrvDesc    As String
    Dim sAllDrives As String
    Dim sDrives()  As String
    Dim cnt        As Long

    idx = 0
    ReDim CDROMs(idx)
    
    sAllDrives = GetDriveString()
    sAllDrives = Replace$(sAllDrives, Chr$(0), Chr$(32))
    sDrives() = Split(Trim$(sAllDrives), Chr$(32))

    For cnt = LBound(sDrives) To UBound(sDrives)

        If GetDriveType(sDrives(cnt)) = DRIVE_CDROM Then
            ReDim Preserve CDROMs(idx)
            
            DrvDesc = "Optical Drive (CD or DVD)"

            Select Case GetDriveTypeEx(sDrives(cnt))
                Case FILE_DEVICE_CD_ROM
                    DrvDesc = "CD-ROM drive"
                Case FILE_DEVICE_DVD
                    DrvDesc = "DVD drive"
            End Select

            CDROMs(idx).Description = DrvDesc
            CDROMs(idx).DriveLetter = sDrives(cnt)
            CDROMs(idx).Virtual = False
            DebugSCSIInfo (sDrives(cnt))
            CDROMs(idx).PNPDeviceId = GetPNPDeviceId(sDrives(cnt))
            
            idx = idx + 1
        End If

    Next

    EnumCDROMs = CDROMs
End Function

Private Function DebugSCSIInfo(DevPath As Variant) As Boolean

    On Error GoTo Err

    Dim hDevice  As Long
    Dim RetVal   As Long
    Dim tmplen   As Long
    Dim SCSIPass As SCSI_PASS_THROUGH_WITH_BUFFERS
    Dim tmpStr   As String
    Dim TmpInt   As Integer

    DebugSCSIInfo = True
    DevPath = UnQualifyPath(DevPath)

    If Len(DevPath) = 2 And Right$(DevPath, 1) = ":" Then
        DevPath = "\\.\" & DevPath
    End If

    hDevice = CreateFile(DevPath, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)

    If hDevice = -1 Then
        Exit Function
    End If

    With SCSIPass.SPT
        .Length = Len(SCSIPass.SPT) + 3
        .PathId = 0
        .TargetId = 1
        .Lun = 0
        .CdbLength = SCSI_PASS_THROUGH_CDBLENGTH.CDB6GENERIC_LENGTH
        .SenseInfoLength = 24
        .DataIn = SCSI_PASS_THROUGH_DATAIN.SCSI_IOCTL_DATA_IN
        .DataTransferLength = 192
        .TimeOutValue = 2
        .DataBufferOffset = Len(SCSIPass.SPT) + 35
        .SenseInfoOffset = Len(SCSIPass.SPT) + 3
        .Cdb(0) = SCSI_PASS_THROUGH_OPERATION.SCSIOP_INQUIRY
        .Cdb(4) = .DataTransferLength
    End With

    tmplen = 0
    
    RetVal = DeviceIoControl(hDevice, IOCOMMANDS.IOCTL_SCSI_PASS_THROUGH, SCSIPass, Len(SCSIPass), SCSIPass, Len(SCSIPass), tmplen, ByVal 0&)

    If SCSIPass.DataBuf(8) Then CDROMs(idx).Manufacturer = GetSTRbyBuff(SCSIPass.DataBuf, 8, 15, False)
    If SCSIPass.DataBuf(16) Then CDROMs(idx).Model = GetSTRbyBuff(SCSIPass.DataBuf, 16, 31, False)
    If SCSIPass.DataBuf(32) Then CDROMs(idx).FirmWare = GetSTRbyBuff(SCSIPass.DataBuf, 32, 35, False)
    If SCSIPass.DataBuf(36) Then CDROMs(idx).SerialNumber = GetSTRbyBuff(SCSIPass.DataBuf, 36, 55, False)
    
    With SCSIPass.SPT
        .Cdb(0) = SCSI_PASS_THROUGH_OPERATION.SCSIOP_MODE_SENSE
        .Cdb(1) = &H8
        .Cdb(2) = SCSI_PASS_THROUGH_CDBMODE.MODE_PAGE_CAPABILITIES
        .Cdb(4) = .DataTransferLength
    End With

    tmplen = 0
    RetVal = DeviceIoControl(hDevice, IOCTL_SCSI_PASS_THROUGH, SCSIPass, Len(SCSIPass), SCSIPass, Len(SCSIPass), tmplen, ByVal 0&)

    If tmplen > 0 Then

        With SCSIPass
            tmpStr = ""

            If (.DataBuf(6) And &H1) Or (.DataBuf(6) And &H2) Then

                If .DataBuf(6) And &H1 Then
                    tmpStr = tmpStr & "CD-R,"
                End If

                If .DataBuf(6) And &H2 Then
                    tmpStr = tmpStr & "CD-RW,"
                End If
            End If

            If (.DataBuf(6) And &H8) Or (.DataBuf(6) And &H10) Or (.DataBuf(6) And &H20) Then

                If .DataBuf(6) And &H10 Then
                    tmpStr = tmpStr & "DVD-R,DVD-RW,"
                End If

                If .DataBuf(6) And &H20 Then
                    tmpStr = tmpStr & "DVD-RAM,"
                End If
            End If
            
            If tmpStr <> vbNullString Then
                tmpStr = Left$(tmpStr, Len(tmpStr) - 1)
                CDROMs(idx).ReadMedia = tmpStr
            End If
            tmpStr = ""

            If (.DataBuf(7) And &H1) Or (.DataBuf(7) And &H2) Then
                If .DataBuf(7) And &H1 Then
                    tmpStr = "CD-R,"
                End If

                If .DataBuf(7) And &H2 Then
                    tmpStr = tmpStr & "CD-RW,"
                End If
            End If

            If (.DataBuf(7) And &H10) Or (.DataBuf(7) And &H20) Then
                If .DataBuf(7) And &H10 Then
                    tmpStr = tmpStr & "DVD-R,DVD-RW,"
                End If

                If .DataBuf(7) And &H20 Then
                    tmpStr = tmpStr & "DVD-RAM,"
                End If
            End If
            
            If tmpStr <> vbNullString Then
                tmpStr = Left$(tmpStr, Len(tmpStr) - 1)
                CDROMs(idx).WriteMedia = tmpStr
            End If
        End With

    End If

    'idx = idx + 1
    Call CloseHandle(hDevice)
    Exit Function
Err:
    DebugSCSIInfo = False
End Function

Private Function GetDriveString() As String

    Dim sBuffer As String

    sBuffer = Space$((26 * 4) + 1)

    If GetLogicalDriveStrings(Len(sBuffer), sBuffer) Then
        GetDriveString = Trim$(sBuffer)
    End If

End Function

Private Function GetDriveTypeEx(sDrive As String) As Long
    GetDriveTypeEx = GetDriveType(sDrive)

    If GetDriveTypeEx = DRIVE_CDROM Then
        GetDriveTypeEx = GetMediaType(sDrive)
    End If

End Function

Private Function GetMediaType(sDrive As String) As Long

    Dim hDrive   As Long
    Dim gmt      As GET_MEDIA_TYPES
    Dim status   As Long
    Dim returned As Long
    Dim mynull   As Long

    sDrive = UnQualifyPath(sDrive)
    hDrive = CreateFile("\\.\" & UCase$(sDrive), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, mynull, OPEN_EXISTING, 0, mynull)

    If hDrive <> INVALID_HANDLE_VALUE Then
        status = DeviceIoControl(hDrive, IOCTL_STORAGE_GET_MEDIA_TYPES_EX, mynull, 0, gmt, 2048, returned, ByVal 0)

        If status <> 0 Then
            GetMediaType = gmt.DeviceType
        End If  'status
    End If  'hDrive

    CloseHandle hDrive
End Function

Private Function UnQualifyPath(ByVal sPath As String) As String

    If Len(sPath) > 0 Then
        sPath = Trim$(sPath)

        If Right$(sPath, 1) = "\" Then
            UnQualifyPath = Left$(sPath, Len(sPath) - 1)
        Else
            UnQualifyPath = sPath
        End If

    Else
        UnQualifyPath = ""
    End If

End Function

Private Function GetSTRbyBuff(ByRef Buffer() As Byte, _
                              Optional ByVal StartIndex As Long, _
                              Optional EndIndex As Long = -1, _
                              Optional ByVal ReturnFor0 As Boolean = True)

    Dim i As Long, DataByte() As Byte

    i = UBound(Buffer)

    If EndIndex = -1 Then
        EndIndex = i
    ElseIf i < EndIndex Then
        EndIndex = i
    End If

    If StartIndex < LBound(Buffer) Then StartIndex = LBound(Buffer)
    i = EndIndex - StartIndex

    If i >= 0 Then
        ReDim DataByte(i)

        For i = 0 To UBound(DataByte)

            If ReturnFor0 And Buffer(i + StartIndex) = 0 Then
                ReDim Preserve DataByte(i - 1)
                Exit For
            End If

            DataByte(i) = Buffer(i + StartIndex)
        Next

        GetSTRbyBuff = Trim$(StrConv(DataByte, vbUnicode))
        GetSTRbyBuff = Replace$(GetSTRbyBuff, vbNullChar, "")
    End If

End Function

Private Function GetPNPDeviceId(sDrive As String) As String
On Error Resume Next
    
    Dim objWMIService, colItems, objItem

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    
    Set colItems = objWMIService.ExecQuery("SELECT PNPDeviceID,Caption FROM Win32_CDROMDrive WHERE Drive='" & sDrive & "'", , 48)
    For Each objItem In colItems
    
        If CDROMs(idx).Model = vbNullString Then
            CDROMs(idx).Model = GetProp(objItem.Caption)
        End If
        CDROMs(idx).Virtual = InStr(LCase$(GetProp(objItem.Caption)), "virtual") > 0
        GetPNPDeviceId = GetProp(objItem.PNPDeviceId)
    Next

    Set objWMIService = Nothing
    Set colItems = Nothing

End Function
