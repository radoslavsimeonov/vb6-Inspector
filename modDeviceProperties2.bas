Attribute VB_Name = "modDeviceProperties2"
Option Explicit

Public Declare Sub CopyMemory _
               Lib "KERNEL32" _
               Alias "RtlMoveMemory" (hpvDest As Any, _
                                      hpvSource As Any, _
                                      ByVal cbCopy As Long)

Public Declare Function CreateFile _
               Lib "KERNEL32" _
               Alias "CreateFileA" (ByVal lpFileName As String, _
                                    ByVal dwDesiredAccess As Long, _
                                    ByVal dwShareMode As Long, _
                                    lpSecurityAttributes As Any, _
                                    ByVal dwCreationDisposition As Long, _
                                    ByVal dwFlagsAndAttributes As Long, _
                                    ByVal hTemplateFile As Long) As Long

Public Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Public Declare Function IIDFromString _
               Lib "ole32" (ByVal lpsz As Long, _
                            ByRef lpiid As GUID) As Long

Private Declare Function SetupDiGetClassDescription _
                Lib "setupapi.dll" _
                Alias "SetupDiGetClassDescriptionA" (ByRef ClassGuid As GUID, _
                                                     ByVal ClassDescription As String, _
                                                     ByVal ClassDescriptionSize As Long, _
                                                     ByRef RequiredSize As Long) As Long

Private Declare Function SetupDiGetClassDevs _
                Lib "setupapi.dll" _
                Alias "SetupDiGetClassDevsA" (ByRef ClassGuid As GUID, _
                                              ByVal Enumerator As String, _
                                              ByVal hwndParent As Long, _
                                              ByVal Flags As DEVICEFLAGS) As Long

Private Declare Function SetupDiDestroyDeviceInfoList _
                Lib "setupapi.dll" (ByVal DeviceInfoSet As Long) As Long

Private Declare Function SetupDiGetDeviceRegistryProperty Lib "setupapi" Alias "SetupDiGetDeviceRegistryPropertyA" (ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO_DATA, ByVal Property As Long, ByRef PropertyRegDataType As Long, ByVal PropertyBuffer As Long, ByVal PropertyBufferSize As Long, RequiredSize As Long) As Long
Private Declare Function SetupDiEnumDeviceInfo _
                Lib "setupapi.dll" (ByVal DeviceInfoSet As Long, _
                                    ByVal MemberIndex As Long, _
                                    ByRef DeviceInfoData As SP_DEVINFO_DATA) As Long

Private Declare Function SetupDiGetDeviceInstanceId Lib "setupapi.dll" Alias "SetupDiGetDeviceInstanceIdA" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByVal DeviceInstanceId As String, ByVal DeviceInstanceIdSize As Long, ByRef RequiredSize As Long) As Long

Private Declare Function SetupDiEnumDeviceInterfaces _
                Lib "setupapi.dll" (ByVal DeviceInfoSet As Long, _
                                    ByRef DeviceInfoData As Any, _
                                    ByRef InterfaceClassGuid As GUID, _
                                    ByVal MemberIndex As Long, _
                                    ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Long

Private Declare Function SetupDiGetDeviceInterfaceDetail _
                Lib "setupapi.dll" _
                Alias "SetupDiGetDeviceInterfaceDetailA" (ByVal DeviceInfoSet As Long, _
                                                          ByRef DeviceInterfaceData As Any, _
                                                          ByRef DeviceInterfaceDetailData As Any, _
                                                          ByVal DeviceInterfaceDetailDataSize As Long, _
                                                          ByRef RequiredSize As Long, _
                                                          ByRef DeviceInfoData As Any) As Long

Private Declare Function SetupDiSetClassInstallParams _
                Lib "setupapi.dll" _
                Alias "SetupDiSetClassInstallParamsA" (ByVal DeviceInfoSet As Long, _
                                                       ByRef DeviceInfoData As SP_DEVINFO_DATA, _
                                                       ByRef ClassInstallParams As SP_CLASSINSTALL_HEADER, _
                                                       ByVal ClassInstallParamsSize As Long) As Long

Private Declare Function SetupDiChangeState _
                Lib "setupapi.dll" (ByVal DeviceInfoSet As Long, _
                                    ByRef DeviceInfoData As SP_DEVINFO_DATA) As Long

Private Declare Function CM_Get_DevNode_Status _
                Lib "setupapi.dll" (lStatus As Long, _
                                    lProblem As Long, _
                                    ByVal hDevice As Long, _
                                    ByVal dwFlags As Long) As Long

Private Declare Function CM_Get_Parent _
                Lib "setupapi.dll" (hParentDevice As Long, _
                                    ByVal hDevice As Long, _
                                    ByVal dwFlags As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const OPEN_EXISTING = 3
Public Const CREATE_NEW = 1
Public Const INVALID_HANDLE_VALUE = -1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const IDENTIFY_BUFFER_SIZE = 512
Public Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16

Private Const DICS_DISABLE As Long = &H2
Private Const DICS_ENABLE As Long = &H1
Private Const DICS_FLAG_GLOBAL As Long = &H1
Private Const DICS_FLAG_CONFIGSPECIFIC As Long = &H2
Private Const DIF_PROPERTYCHANGE As Long = &H12

Private Type SP_CLASSINSTALL_HEADER
    cbSize As Long
    InstallFunction As Long
End Type

Private Type SP_PROPCHANGE_PARAMS
    ClassInstallHeader As SP_CLASSINSTALL_HEADER
    StateChange As Long
    Scope As Long
    HwProfile As Long
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Enum DEVICEFLAGS
    DIGCF_ALLCLASSES = &H4&
    DIGCF_DEVICEINTERFACE = &H10
    DIGCF_PRESENT = &H2
    DIGCF_PROFILE = &H8
End Enum

Private Type SP_DEVINFO_DATA
    cbSize As Long
    ClassGuid As GUID
    devInst As Long
    reserved As Long
End Type

Private Enum DEVICEPROPERTYINDEX
    SPDRP_ADDRESS = (&H1C)
    SPDRP_BUSNUMBER = (&H15)
    SPDRP_BUSTYPEGUID = (&H13)
    SPDRP_CAPABILITIES = (&HF)
    SPDRP_CHARACTERISTICS = (&H1B)
    SPDRP_CLASS = (&H7)
    SPDRP_CLASSGUID = (&H8)
    SPDRP_COMPATIBLEIDS = (&H2)
    SPDRP_CONFIGFLAGS = (&HA)
    SPDRP_DEVICEDESC = &H0
    SPDRP_DEVTYPE = (&H19)
    SPDRP_DRIVER = (&H9)
    SPDRP_ENUMERATOR_NAME = (&H16)
    SPDRP_EXCLUSIVE = (&H1A)
    SPDRP_FRIENDLYNAME = (&HC)
    SPDRP_HARDWAREID = (&H1)
    SPDRP_LEGACYBUSTYPE = (&H14)
    SPDRP_LOCATION_INFORMATION = (&HD)
    SPDRP_LOWERFILTERS = (&H12)
    SPDRP_MFG = (&HB)
    SPDRP_PHYSICAL_DEVICE_OBJECT_NAME = (&HE)
    SPDRP_SECURITY = (&H17)
    SPDRP_SECURITY_SDS = (&H18)
    SPDRP_SERVICE = (&H4)
    SPDRP_UI_NUMBER = (&H10)
    SPDRP_UPPERFILTERS = (&H11)
    SPDRP_INSTALL_STATE = (&H22)
End Enum

Private Enum DN_Flags
    DN_ROOT_ENUMERATED = &H1    ' Was enumerated by ROOT
    DN_DRIVER_LOADED = &H2    ' Has Register_Device_Driver
    DN_ENUM_LOADED = &H4    ' Has Register_Enumerator
    DN_STARTED = &H8    ' Is currently configured
    DN_MANUAL = &H10    ' Manually installed
    DN_NEED_TO_ENUM = &H20    ' May need reenumeration
    DN_NOT_FIRST_TIME = &H40    ' Has received a config
    DN_HARDWARE_ENUM = &H80    ' Enum generates hardware ID
    DN_LIAR = &H100    ' Lied about can reconfig once
    DN_HAS_MARK = &H200    ' Not CM_Create_DevInst lately
    DN_HAS_PROBLEM = &H400    ' Need device installer
    DN_FILTERED = &H800    ' Is filtered
    DN_MOVED = &H1000    ' Has been moved
    DN_DISABLEABLE = &H2000    ' Can be disabled
    DN_REMOVABLE = &H4000    ' Can be removed
    DN_PRIVATE_PROBLEM = &H8000    ' Has a private problem
    DN_MF_PARENT = &H10000    ' Multi function parent
    DN_MF_CHILD = &H20000    ' Multi function child
    DN_WILL_BE_REMOVED = &H40000    ' DevInst is being removed
    DN_NOT_FIRST_TIMEE = &H80000    ' Has received a config enumerate
    DN_STOP_FREE_RES = &H100000    ' When child is stopped, free resources
    DN_REBAL_CANDIDATE = &H200000    ' Don't skip during rebalance
    DN_BAD_PARTIAL = &H400000    ' This devnode's log_confs do not have same resources
    DN_NT_ENUMERATOR = &H800000    ' This devnode's is an NT enumerator
    DN_NT_DRIVER = &H1000000    ' This devnode's is an NT driver
    DN_NEEDS_LOCKING = &H2000000    ' Devnode need lock resume processing
    DN_ARM_WAKEUP = &H4000000    ' Devnode can be the wakeup device
    DN_APM_ENUMERATOR = &H8000000    ' APM aware enumerator
    DN_APM_DRIVER = &H10000000    ' APM aware driver
    DN_SILENT_INSTALL = &H20000000    ' Silent install
    DN_NO_SHOW_IN_DM = &H40000000    ' No show in device manager
    DN_BOOT_LOG_PROB = &H80000000    ' Had a problem during preassignment of boot log conf
End Enum

Public Enum CM_PROB
    CM_PROB_NOT_CONFIGURED = &H1    ' no config for device
    CM_PROB_DEVLOADER_FAILED = &H2    ' service load failed
    CM_PROB_OUT_OF_MEMORY = &H3    ' out of memory
    CM_PROB_ENTRY_IS_WRONG_TYPE = &H4    '
    CM_PROB_LACKED_ARBITRATOR = &H5    '
    CM_PROB_BOOT_CONFIG_CONFLICT = &H6    ' boot config conflict
    CM_PROB_FAILED_FILTER = &H7    '
    CM_PROB_DEVLOADER_NOT_FOUND = &H8    ' Devloader not found
    CM_PROB_INVALID_DATA = &H9    ' Invalid ID
    CM_PROB_FAILED_START = &HA    '
    CM_PROB_LIAR = &HB    '
    CM_PROB_NORMAL_CONFLICT = &HC    ' config conflict
    CM_PROB_NOT_VERIFIED = &HD    '
    CM_PROB_NEED_RESTART = &HE    ' requires restart
    CM_PROB_REENUMERATION = &HF    '
    CM_PROB_PARTIAL_LOG_CONF = &H10    '
    CM_PROB_UNKNOWN_RESOURCE = &H11    ' unknown res type
    CM_PROB_REINSTALL = &H12    '
    CM_PROB_REGISTRY = &H13    '
    CM_PROB_VXDLDR = &H14    ' WINDOWS 95 ONLY
    CM_PROB_WILL_BE_REMOVED = &H15    ' devinst will remove
    CM_PROB_DISABLED = &H16    ' devinst is disabled
    CM_PROB_DEVLOADER_NOT_READY = &H17    ' Devloader not ready
    CM_PROB_DEVICE_NOT_THERE = &H18    ' device doesn't exist
    CM_PROB_MOVED = &H19    '
    CM_PROB_TOO_EARLY = &H1A    '
    CM_PROB_NO_VALID_LOG_CONF = &H1B    ' no valid log config
    CM_PROB_FAILED_INSTALL = &H1C    ' install failed
    CM_PROB_HARDWARE_DISABLED = &H1D    ' device disabled
    CM_PROB_CANT_SHARE_IRQ = &H1E    ' can't share IRQ
    CM_PROB_FAILED_ADD = &H1F    ' driver failed add
    CM_PROB_DISABLED_SERVICE = &H20    ' service's Start = 4
    CM_PROB_TRANSLATION_FAILED = &H21    ' resource translation failed
    CM_PROB_NO_SOFTCONFIG = &H22    ' no soft config
    CM_PROB_BIOS_TABLE = &H23    ' device missing in BIOS table
    CM_PROB_IRQ_TRANSLATION_FAILED = &H24    ' IRQ translator failed
    CM_PROB_FAILED_DRIVER_ENTRY = &H25    ' DriverEntry() failed.
    CM_PROB_DRIVER_FAILED_PRIOR_UNLOAD = &H26    ' Driver should have unloaded.
    CM_PROB_DRIVER_FAILED_LOAD = &H27    ' Driver load unsuccessful.
    CM_PROB_DRIVER_SERVICE_KEY_INVALID = &H28    ' Error accessing driver's service key
    CM_PROB_LEGACY_SERVICE_NO_DEVICES = &H29    ' Loaded legacy service created no devices
    CM_PROB_DUPLICATE_DEVICE = &H2A    ' Two devices were discovered with the same name
    CM_PROB_FAILED_POST_START = &H2B    ' The drivers set the device state to failed
    CM_PROB_HALTED = &H2C    ' This device was failed post start via usermode
    CM_PROB_PHANTOM = &H2D    ' The devinst currently exists only in the registry
    CM_PROB_SYSTEM_SHUTDOWN = &H2E    ' The system is shutting down
    CM_PROB_HELD_FOR_EJECT = &H2F    ' The device is offline awaiting removal
    CM_PROB_DRIVER_BLOCKED = &H30    ' One or more drivers is blocked from loading
    CM_PROB_REGISTRY_TOO_LARGE = &H31    ' System hive has grown too large
    CM_PROB_SETPROPERTIES_FAILED = &H32    ' Failed to apply one or more registry properties
    NUM_CM_PROB = &H33    '
End Enum

Public Enum IOCOMMANDS
    IOCTL_STORAGE_CHECK_VERIFY = &H2D4800
    IOCTL_STORAGE_CHECK_VERIFY2 = &H2D0800
    IOCTL_STORAGE_MEDIA_REMOVAL = &H2D4804
    IOCTL_STORAGE_EJECT_MEDIA = &H2D4808
    IOCTL_STORAGE_LOAD_MEDIA = &H2D480C
    IOCTL_STORAGE_LOAD_MEDIA2 = &H2D080C
    IOCTL_STORAGE_RESERVE = &H2D4810
    IOCTL_STORAGE_RELEASE = &H2D4814
    IOCTL_STORAGE_FIND_NEW_DEVICES = &H2D4818
    IOCTL_STORAGE_EJECTION_CONTROL = &H2D0940
    IOCTL_STORAGE_MCN_CONTROL = &H2D0944
    IOCTL_STORAGE_GET_MEDIA_TYPES = &H2D0C00
    IOCTL_STORAGE_GET_MEDIA_TYPES_EX = &H2D0C04
    IOCTL_STORAGE_GET_MEDIA_SERIAL_NUMBER = &H2D0C10
    IOCTL_STORAGE_GET_HOTPLUG_INFO = &H2D0C14
    IOCTL_STORAGE_SET_HOTPLUG_INFO = &H2DCC18
    IOCTL_STORAGE_RESET_BUS = &H2D5000
    IOCTL_STORAGE_RESET_DEVICE = &H2D5004
    IOCTL_STORAGE_BREAK_RESERVATION = &H2D5014
    IOCTL_STORAGE_GET_DEVICE_NUMBER = &H2D1080
    IOCTL_STORAGE_PREDICT_FAILURE = &H2D1100
    IOCTL_STORAGE_QUERY_PROPERTY = &H2D1400
    IOCTL_SCSI_PASS_THROUGH = &H4D004
    IOCTL_SCSI_MINIPORT = &H4D008
    IOCTL_SCSI_GET_INQUIRY_DATA = &H4100C
    IOCTL_SCSI_GET_CAPABILITIES = &H41010
    IOCTL_SCSI_PASS_THROUGH_DIRECT = &H4D014
    IOCTL_SCSI_GET_ADDRESS = &H41018
    IOCTL_SCSI_RESCAN_BUS = &H4101C
    IOCTL_SCSI_GET_DUMP_POINTERS = &H41020
    IOCTL_SCSI_FREE_DUMP_POINTERS = &H41024
    IOCTL_IDE_PASS_THROUGH = &H4D028
End Enum

Private Enum REGPROPERTYTYPES
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_EXPAND_SZ = 2
    REG_MULTI_SZ = 7
    REG_SZ = 1
End Enum

Private Type SP_DEVICE_INTERFACE_DATA
    cbSize As Long
    InterfaceClassGuid As GUID
    devInst As Long
    reserved As Long
End Type

Private Enum SetupErrors
    ERROR_INSUFFICIENT_BUFFER = 122
    ERROR_INVALID_DATA = 13&
    ERROR_NO_MORE_ITEMS = 259&
End Enum

Private Const GUID_ACPI_CMOS_INTERFACE_STANDARD = "{3a8d0384-6505-40ca-bc39-56c15f8c5fed}"
Private Const GUID_ACPI_INTERFACE_STANDARD = "{b091a08a-ba97-11d0-bd14-00aa00b7b32a}"
Private Const GUID_ACPI_PORT_RANGES_INTERFACE_STANDARD = "{f14f609b-cbbd-4957-a674-bc00213f1c97}"
Private Const GUID_ACPI_REGS_INTERFACE_STANDARD = "{06141966-7245-6369-462e-4e656c736f6e}"
Private Const GUID_AGP_TARGET_BUS_INTERFACE_STANDARD = "{b15cfce8-06d1-4d37-9d4c-bedde0c2a6ff}"
Private Const GUID_ARBITER_INTERFACE_STANDARD = "{e644f185-8c0e-11d0-becf-08002be2092f}"
Private Const GUID_BUS_INTERFACE_STANDARD = "{496b8280-6f25-11d0-beaf-08002be2092f}"
Private Const GUID_BUS_TYPE_1394 = "{f74e73eb-9ac5-45eb-be4d-772cc71ddfb3}"
Private Const GUID_BUS_TYPE_AVC = "{c06ff265-ae09-48f0-812c-16753d7cba83}"
Private Const GUID_BUS_TYPE_DOT4PRT = "{441ee001-4342-11d5-a184-00c04f60524d}"
Private Const GUID_BUS_TYPE_EISA = "{ddc35509-f3fc-11d0-a537-0000f8753ed1}"
Private Const GUID_BUS_TYPE_HID = "{eeaf37d0-1963-47c4-aa48-72476db7cf49}"
Private Const GUID_BUS_TYPE_INTERNAL = "{1530ea73-086b-11d1-a09f-00c04fc340b1}"
Private Const GUID_BUS_TYPE_IRDA = "{7ae17dc1-c944-44d6-881f-4c2e61053bc1}"
Private Const GUID_BUS_TYPE_ISAPNP = "{e676f854-d87d-11d0-92b2-00a0c9055fc5}"
Private Const GUID_BUS_TYPE_LPTENUM = "{c4ca1000-2ddc-11d5-a17a-00c04f60524d}"
Private Const GUID_BUS_TYPE_MCA = "{1c75997a-dc33-11d0-92b2-00a0c9055fc5}"
Private Const GUID_BUS_TYPE_PCI = "{c8ebdfb0-b510-11d0-80e5-00a0c92542e3}"
Private Const GUID_BUS_TYPE_PCMCIA = "{09343630-af9f-11d0-92e9-0000f81e1b30}"
Private Const GUID_BUS_TYPE_SD = "{e700cc04-4036-4e89-9579-89ebf45f00cd}"
Private Const GUID_BUS_TYPE_SERENUM = "{77114a87-8944-11d1-bd90-00a0c906be2d}"
Private Const GUID_BUS_TYPE_USB = "{9d7debbc-c85d-11d1-9eb4-006008c3a19a}"
Private Const GUID_BUS_TYPE_USBPRINT = "{441ee000-4342-11d5-a184-00c04f60524d}"

Private Const GUID_DEVCLASS_1394 = "{6bdd1fc1-810f-11d0-bec7-08002be2092f}"
Private Const GUID_DEVCLASS_1394DEBUG = "{66f250d6-7801-4a64-b139-eea80a450b24}"
Private Const GUID_DEVCLASS_61883 = "{7ebefbc0-3200-11d2-b4c2-00a0c9697d07}"
Private Const GUID_DEVCLASS_ADAPTER = "{4d36e964-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_APMSUPPORT = "{d45b1c18-c8fa-11d1-9f77-0000f805f530}"
Private Const GUID_DEVCLASS_AVC = "{c06ff265-ae09-48f0-812c-16753d7cba83}"
Private Const GUID_DEVCLASS_BATTERY = "{72631e54-78a4-11d0-bcf7-00aa00b7b32a}"
Private Const GUID_DEVCLASS_BIOMETRIC = "{53d29ef7-377c-4d14-864b-eb3a85769359}"
Private Const GUID_DEVCLASS_BLUETOOTH = "{e0cbf06c-cd8b-4647-bb8a-263b43f0f974}"
Private Const GUID_DEVCLASS_CDROM = "{4d36e965-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_COMPUTER = "{4d36e966-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_DECODER = "{6bdd1fc2-810f-11d0-bec7-08002be2092f}"
Private Const GUID_DEVCLASS_DISKDRIVE = "{4d36e967-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_DISPLAY = "{4d36e968-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_DOT4 = "{48721b56-6795-11d2-b1a8-0080c72e74a2}"
Private Const GUID_DEVCLASS_DOT4PRINT = "{49ce6ac8-6f86-11d2-b1e5-0080c72e74a2}"
Private Const GUID_DEVCLASS_ENUM1394 = "{c459df55-db08-11d1-b009-00a0c9081ff6}"
Private Const GUID_DEVCLASS_FDC = "{4d36e969-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_FLOPPYDISK = "{4d36e980-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_FSFILTER_ACTIVITYMONITOR = "{b86dff51-a31e-4bac-b3cf-e8cfe75c9fc2}"
Private Const GUID_DEVCLASS_FSFILTER_ANTIVIRUS = "{b1d1a169-c54f-4379-81db-bee7d88d7454}"
Private Const GUID_DEVCLASS_FSFILTER_CFSMETADATASERVER = "{cdcf0939-b75b-4630-bf76-80f7ba655884}"
Private Const GUID_DEVCLASS_FSFILTER_COMPRESSION = "{f3586baf-b5aa-49b5-8d6c-0569284c639f}"
Private Const GUID_DEVCLASS_FSFILTER_CONTENTSCREENER = "{3e3f0674-c83c-4558-bb26-9820e1eba5c5}"
Private Const GUID_DEVCLASS_FSFILTER_CONTINUOUSBACKUP = "{71aa14f8-6fad-4622-ad77-92bb9d7e6947}"
Private Const GUID_DEVCLASS_FSFILTER_COPYPROTECTION = "{89786ff1-9c12-402f-9c9e-17753c7f4375}"
Private Const GUID_DEVCLASS_FSFILTER_ENCRYPTION = "{a0a701c0-a511-42ff-aa6c-06dc0395576f}"
Private Const GUID_DEVCLASS_FSFILTER_HSM = "{d546500a-2aeb-45f6-9482-f4b1799c3177}"
Private Const GUID_DEVCLASS_FSFILTER_INFRASTRUCTURE = "{e55fa6f9-128c-4d04-abab-630c74b1453a}"
Private Const GUID_DEVCLASS_FSFILTER_OPENFILEBACKUP = "{f8ecafa6-66d1-41a5-899b-66585d7216b7}"
Private Const GUID_DEVCLASS_FSFILTER_PHYSICALQUOTAMANAGEMENT = "{6a0a8e78-bba6-4fc4-a709-1e33cd09d67e}"
Private Const GUID_DEVCLASS_FSFILTER_QUOTAMANAGEMENT = "{8503c911-a6c7-4919-8f79-5028f5866b0c}"
Private Const GUID_DEVCLASS_FSFILTER_REPLICATION = "{48d3ebc4-4cf8-48ff-b869-9c68ad42eb9f}"
Private Const GUID_DEVCLASS_FSFILTER_SECURITYENHANCER = "{d02bc3da-0c8e-4945-9bd5-f1883c226c8c}"
Private Const GUID_DEVCLASS_FSFILTER_SYSTEM = "{5d1b9aaa-01e2-46af-849f-272b3f324c46}"
Private Const GUID_DEVCLASS_FSFILTER_SYSTEMRECOVERY = "{2db15374-706e-4131-a0c7-d7c78eb0289a}"
Private Const GUID_DEVCLASS_FSFILTER_UNDELETE = "{fe8f1572-c67a-48c0-bbac-0b5c6d66cafb}"
Private Const GUID_DEVCLASS_GPS = "{6bdd1fc3-810f-11d0-bec7-08002be2092f}"
Private Const GUID_DEVCLASS_HDC = "{4d36e96a-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_HIDCLASS = "{745a17a0-74d3-11d0-b6fe-00a0c90f57da}"
Private Const GUID_DEVCLASS_IMAGE = "{6bdd1fc6-810f-11d0-bec7-08002be2092f}"
Private Const GUID_DEVCLASS_INFINIBAND = "{30ef7132-d858-4a0c-ac24-b9028a5cca3f}"
Private Const GUID_DEVCLASS_INFRARED = "{6bdd1fc5-810f-11d0-bec7-08002be2092f}"
Private Const GUID_DEVCLASS_KEYBOARD = "{4d36e96b-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_LEGACYDRIVER = "{8ecc055d-047f-11d1-a537-0000f8753ed1}"
Private Const GUID_DEVCLASS_MEDIA = "{4d36e96c-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_MEDIUM_CHANGER = "{ce5939ae-ebde-11d0-b181-0000f8753ec4}"
Private Const GUID_DEVCLASS_MODEM = "{4d36e96d-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_MONITOR = "{4d36e96e-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_MOUSE = "{4d36e96f-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_MTD = "{4d36e970-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_MULTIFUNCTION = "{4d36e971-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_MULTIPORTSERIAL = "{50906cb8-ba12-11d1-bf5d-0000f805f530}"
Private Const GUID_DEVCLASS_NET = "{4d36e972-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_NETCLIENT = "{4d36e973-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_NETSERVICE = "{4d36e974-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_NETTRANS = "{4d36e975-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_NODRIVER = "{4d36e976-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_PCMCIA = "{4d36e977-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_PNPPRINTERS = "{4658ee7e-f050-11d1-b6bd-00c04fa372a7}"
Private Const GUID_DEVCLASS_PORTS = "{4d36e978-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_PRINTER = "{4d36e979-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_PRINTERUPGRADE = "{4d36e97a-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_PROCESSOR = "{50127dc3-0f36-415e-a6cc-4cb3be910b65}"
Private Const GUID_DEVCLASS_SBP2 = "{d48179be-ec20-11d1-b6b8-00c04fa372a7}"
Private Const GUID_DEVCLASS_SCSIADAPTER = "{4d36e97b-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_SECURITYACCELERATOR = "{268c95a1-edfe-11d3-95c3-0010dc4050a5}"
Private Const GUID_DEVCLASS_SMARTCARDREADER = "{50dd5230-ba8a-11d1-bf5d-0000f805f530}"
Private Const GUID_DEVCLASS_SOUND = "{4d36e97c-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_SYSTEM = "{4d36e97d-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_TAPEDRIVE = "{6d807884-7d21-11cf-801c-08002be10318}"
Private Const GUID_DEVCLASS_UNKNOWN = "{4d36e97e-e325-11ce-bfc1-08002be10318}"
Private Const GUID_DEVCLASS_USB = "{36fc9e60-c465-11cf-8056-444553540000}"
Private Const GUID_DEVCLASS_VOLUME = "{71a27cdd-812a-11d0-bec7-08002be2092f}"
Private Const GUID_DEVCLASS_VOLUMESNAPSHOT = "{533c5b84-ec70-11d2-9505-00c04f79deaf}"
Private Const GUID_DEVCLASS_WCEUSBS = "{25dbce51-6c8f-4a72-8a6d-b54c2b4fc835}"

Private Const GUID_DEVICE_ARRIVAL = "{cb3a4009-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_EJECT = "{cb3a400f-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_EJECT_VETOED = "{cf7b71e8-d8fd-11d2-97b5-00a0c940522e}"
Private Const GUID_DEVICE_ENUMERATE_REQUEST = "{cb3a400b-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_ENUMERATED = "{cb3a400a-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_EVENT_RBC = "{d0744792-a98e-11d2-917a-00a0c9068ff3}"
Private Const GUID_DEVICE_HIBERNATE_VETOED = "{61173ad9-194f-11d3-97dc-00a0c940522e}"
Private Const GUID_DEVICE_INTERFACE_ARRIVAL = "{cb3a4004-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_INTERFACE_REMOVAL = "{cb3a4005-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_INVALID_ID = "{57a49b33-8b85-4e75-a081-166ce241f407}"
Private Const GUID_DEVICE_KERNEL_INITIATED_EJECT = "{14689b54-0703-11d3-97d2-00a0c940522e}"
Private Const GUID_DEVICE_NOOP = "{cb3a4010-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_QUERY_AND_REMOVE = "{cb3a400e-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_REMOVAL_VETOED = "{60dbd5fa-ddd2-11d2-97b8-00a0c940522e}"
Private Const GUID_DEVICE_REMOVE_PENDING = "{cb3a400d-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_SAFE_REMOVAL = "{8fbef967-d6c5-11d2-97b5-00a0c940522e}"
Private Const GUID_DEVICE_STANDBY_VETOED = "{03b21c13-18d6-11d3-97db-00a0c940522e}"
Private Const GUID_DEVICE_START_REQUEST = "{cb3a400c-46f0-11d0-b08f-00609713053f}"
Private Const GUID_DEVICE_SURPRISE_REMOVAL = "{ce5af000-80dd-11d2-a88d-00a0c9696b4b}"
Private Const GUID_DEVICE_WARM_EJECT_VETOED = "{cbf4c1f9-18d5-11d3-97db-00a0c940522e}"

Private Const GUID_DEVINTERFACE_CDCHANGER = "{53f56312-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_CDROM = "{53f56308-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_DISK = "{53f56307-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_FLOPPY = "{53f56311-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_MEDIUMCHANGER = "{53f56310-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_PARTITION = "{53f5630a-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_PROCESSOR = "{97fadb10-4e33-40ae-359c-8bef029dbdd0}"
Private Const GUID_DEVINTERFACE_STORAGEPORT = "{2accfe60-c130-11d2-b082-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_TAPE = "{53f5630b-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_VOLUME = "{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_DEVINTERFACE_WRITEONCEDISK = "{53f5630c-b6bf-11d0-94f2-00a0c91efb8b}"

Private Const GUID_DOCK_INTERFACE = "{a9956ff5-13da-11d3-97db-00a0c940522e}"
Private Const GUID_DRIVER_BLOCKED = "{1bc87a21-a3ff-47a6-96aa-6d010906805a}"
Private Const GUID_HWPROFILE_CHANGE_CANCELLED = "{cb3a4002-46f0-11d0-b08f-00609713053f}"
Private Const GUID_HWPROFILE_CHANGE_COMPLETE = "{cb3a4003-46f0-11d0-b08f-00609713053f}"
Private Const GUID_HWPROFILE_QUERY_CHANGE = "{cb3a4001-46f0-11d0-b08f-00609713053f}"
Private Const GUID_INT_ROUTE_INTERFACE_STANDARD = "{70941bf4-0073-11d1-a09e-00c04fc340b1}"
Private Const GUID_IO_DEVICE_BECOMING_READY = "{d07433f0-a98e-11d2-917a-00a0c9068ff3}"
Private Const GUID_IO_DEVICE_EXTERNAL_REQUEST = "{d07433d0-a98e-11d2-917a-00a0c9068ff3}"
Private Const GUID_IO_DISK_CLONE_ARRIVAL = "{6a61885b-7c39-43dd-9b56-b8ac22a549aa}"
Private Const GUID_IO_DISK_LAYOUT_CHANGE = "{11dff54c-8469-41f9-b3de-ef836487c54a}"
Private Const GUID_IO_DRIVE_REQUIRES_CLEANING = "{7207877c-90ed-44e5-a000-81428d4c79bb}"
Private Const GUID_IO_MEDIA_ARRIVAL = "{d07433c0-a98e-11d2-917a-00a0c9068ff3}"
Private Const GUID_IO_MEDIA_EJECT_REQUEST = "{d07433d1-a98e-11d2-917a-00a0c9068ff3}"
Private Const GUID_IO_MEDIA_REMOVAL = "{d07433c1-a98e-11d2-917a-00a0c9068ff3}"
Private Const GUID_IO_TAPE_ERASE = "{852d11eb-4bb8-4507-9d9b-417cc2b1b438}"
Private Const GUID_IO_VOLUME_CHANGE = "{7373654a-812a-11d0-bec7-08002be2092f}"
Private Const GUID_IO_VOLUME_DEVICE_INTERFACE = "{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}"
Private Const GUID_IO_VOLUME_DISMOUNT = "{d16a55e8-1059-11d2-8ffd-00a0c9a06d32}"
Private Const GUID_IO_VOLUME_DISMOUNT_FAILED = "{e3c5b178-105d-11d2-8ffd-00a0c9a06d32}"
Private Const GUID_IO_VOLUME_LOCK = "{50708874-c9af-11d1-8fef-00a0c9a06d32}"
Private Const GUID_IO_VOLUME_LOCK_FAILED = "{ae2eed10-0ba8-11d2-8ffb-00a0c9a06d32}"
Private Const GUID_IO_VOLUME_MOUNT = "{b5804878-1a96-11d2-8ffd-00a0c9a06d32}"
Private Const GUID_IO_VOLUME_NAME_CHANGE = "{2de97f83-4c06-11d2-a532-00609713055a}"
Private Const GUID_IO_VOLUME_PHYSICAL_CONFIGURATION_CHANGE = "{2de97f84-4c06-11d2-a532-00609713055a}"
Private Const GUID_IO_VOLUME_UNLOCK = "{9a8c3d68-d0cb-11d1-8fef-00a0c9a06d32}"
Private Const GUID_LEGACY_DEVICE_DETECTION_STANDARD = "{50feb0de-596a-11d2-a5b8-0000f81a4619}"
Private Const GUID_MF_ENUMERATION_INTERFACE = "{aeb895f0-5586-11d1-8d84-00a0c906b244}"
Private Const GUID_MOF_RESOURCE_ADDED_NOTIFICATION = "{b48d49a2-e777-11d0-a50c-00a0c9062910}"
Private Const GUID_MOF_RESOURCE_REMOVED_NOTIFICATION = "{b48d49a3-e777-11d0-a50c-00a0c9062910}"
Private Const GUID_PCI_BUS_INTERFACE_STANDARD = "{496b8281-6f25-11d0-beaf-08002be2092f}"
Private Const GUID_PCI_DEVICE_PRESENT_INTERFACE = "{d1b82c26-bf49-45ef-b216-71cbd7889b57}"
Private Const GUID_PCMCIA_BUS_INTERFACE_STANDARD = "{76173af0-c504-11d1-947f-00c04fb960ee}"
Private Const GUID_PNP_CUSTOM_NOTIFICATION = "{aca73f8e-8d23-11d1-ac7d-0000f87571d0}"
Private Const GUID_PNP_LOCATION_INTERFACE = "{70211b0e-0afb-47db-afc1-410bf842497a}"
Private Const GUID_PNP_POWER_NOTIFICATION = "{c2cf0660-eb7a-11d1-bd7f-0000f87571d0}"
Private Const GUID_POWER_DEVICE_ENABLE = "{827c0a6f-feb0-11d0-bd26-00aa00b7b32a}"
Private Const GUID_POWER_DEVICE_TIMEOUTS = "{a45da735-feb0-11d0-bd26-00aa00b7b32a}"
Private Const GUID_POWER_DEVICE_WAKE_ENABLE = "{a9546a82-feb0-11d0-bd26-00aa00b7b32a}"
Private Const GUID_REGISTRATION_CHANGE_NOTIFICATION = "{b48d49a1-e777-11d0-a50c-00a0c9062910}"
Private Const GUID_SETUP_DEVICE_ARRIVAL = "{cb3a4000-46f0-11d0-b08f-00609713053f}"
Private Const GUID_TARGET_DEVICE_QUERY_REMOVE = "{cb3a4006-46f0-11d0-b08f-00609713053f}"
Private Const GUID_TARGET_DEVICE_REMOVE_CANCELLED = "{cb3a4007-46f0-11d0-b08f-00609713053f}"
Private Const GUID_TARGET_DEVICE_REMOVE_COMPLETE = "{cb3a4008-46f0-11d0-b08f-00609713053f}"
Private Const GUID_TRANSLATOR_INTERFACE_STANDARD = "{6c154a92-aacf-11d0-8d2a-00a0c906b244}"
Private Const MAX_PATH As Long = 260

Public Function EnableDevice(lEnumerator As Long, ByVal bEnable As Boolean) As Boolean

    Dim changeParams As SP_PROPCHANGE_PARAMS
    Dim hDevInfo     As Long
    Dim DevInfo      As SP_DEVINFO_DATA
    Dim tGUID        As GUID

    hDevInfo = SetupDiGetClassDevs(tGUID, vbNullString, 0, DEVICEFLAGS.DIGCF_PRESENT Or DEVICEFLAGS.DIGCF_ALLCLASSES)
    DevInfo.cbSize = LenB(DevInfo)

    If SetupDiEnumDeviceInfo(hDevInfo, lEnumerator, DevInfo) = 0 Then
        SetupDiDestroyDeviceInfoList hDevInfo
        Exit Function
    End If

    With changeParams
        .ClassInstallHeader.cbSize = LenB(.ClassInstallHeader)
        .ClassInstallHeader.InstallFunction = DIF_PROPERTYCHANGE
        .Scope = DICS_FLAG_CONFIGSPECIFIC    'DICS_FLAG_GLOBAL Or
        .StateChange = IIf(bEnable, DICS_ENABLE, DICS_DISABLE)
        .HwProfile = 0
    End With

    If SetupDiSetClassInstallParams(hDevInfo, DevInfo, changeParams.ClassInstallHeader, LenB(changeParams)) = 1 Then
        EnableDevice = (SetupDiChangeState(hDevInfo, DevInfo) = 1)
    End If

End Function

Public Function GetDevProp(sGUID As String, ByRef tDev() As HardwareDevice) As Boolean

    Dim i             As Integer
    Dim RetVal        As Long
    Dim SE            As Long
    Dim status        As Long
    Dim Problem       As Long
    Dim DIDC          As SP_DEVINFO_DATA
    Dim DIDI          As SP_DEVICE_INTERFACE_DATA
    Dim tGUID         As GUID
    Dim hDevInfoC     As Long
    Dim hDevInfoI     As Long
    Dim szBuf         As String
    Dim dwRequireSize As Long

    If sGUID = "" Then
        hDevInfoC = SetupDiGetClassDevs(tGUID, vbNullString, 0, DEVICEFLAGS.DIGCF_PRESENT Or DEVICEFLAGS.DIGCF_ALLCLASSES)
    Else

        If CheckGUID(sGUID) Then
            Call IIDFromString(StrPtr(sGUID), tGUID)
            hDevInfoC = SetupDiGetClassDevs(tGUID, vbNullString, 0, DEVICEFLAGS.DIGCF_PRESENT)
        Else
            hDevInfoC = SetupDiGetClassDevs(tGUID, sGUID, 0, DEVICEFLAGS.DIGCF_DEVICEINTERFACE Or DEVICEFLAGS.DIGCF_ALLCLASSES)
        End If
    End If

    GetDevProp = False
    DIDC.cbSize = Len(DIDC)
    DIDI.cbSize = Len(DIDI)

    Do
        RetVal = SetupDiEnumDeviceInfo(hDevInfoC, i, DIDC)

        If RetVal = 0 Then
            SE = Err.LastDllError

            If SE = ERROR_NO_MORE_ITEMS Then
                Exit Do
            Else
                GetDevProp = False
                Exit Function
            End If
        End If

        ReDim Preserve tDev(i)

        With tDev(i)
            
            dwRequireSize = 128: szBuf = String$(dwRequireSize, vbNullChar)
            
            SetupDiGetDeviceInstanceId hDevInfoC, DIDC, vbNullString, 0&, dwRequireSize
            szBuf = Space$(dwRequireSize)
            If SetupDiGetDeviceInstanceId(hDevInfoC, DIDC, szBuf, Len(szBuf), dwRequireSize) Then
                .PNPDeviceId = Trim$(Replace$(szBuf, vbNullChar, ""))
            End If
            
            SetupDiGetClassDescription DIDC.ClassGuid, vbNullString, 0&, dwRequireSize
            szBuf = Space$(dwRequireSize)

            If (SetupDiGetClassDescription(DIDC.ClassGuid, szBuf, dwRequireSize, dwRequireSize)) Then
                .ClassDesc = Trim$(Replace$(szBuf, vbNullChar, ""))
            End If

            If .ClassDesc = "" Then .ClassDesc = "Other devices"
            CM_Get_DevNode_Status status, Problem, DIDC.devInst, 0
            .Enabled = (Problem And CM_PROB.CM_PROB_DISABLED) = 0
            .Hidden = Not ((status And DN_NO_SHOW_IN_DM) = 0)
            .CanDisable = Not (((status And DN_DISABLEABLE) = 0) And ((Problem And CM_PROB.CM_PROB_DISABLED) = 0))
            .Removable = Not ((status And DN_REMOVABLE) = 0)
            .DeviceStatus = DeviceStatusMessage(Problem)
            .Index = i
            .devInst = DIDC.devInst
            .Class = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_CLASS)
            .BusNumber = CLng(val(GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_BUSNUMBER, True)))
            .ClassGuid = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_CLASSGUID)
            .DeviceDesc = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_DEVICEDESC)
            .Driver = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_DRIVER)
            .Location = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_LOCATION_INFORMATION)
            .EnumeratorName = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_ENUMERATOR_NAME)
            .FriendlyName = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_FRIENDLYNAME)
            .HardwareIDs = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_HARDWAREID)
            .HardwareIDs = Replace(.HardwareIDs, vbNullChar, "")
            .Manufacturer = GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_MFG)
            CM_Get_Parent .DevParent, DIDC.devInst, 0&

            Select Case UCase$(.EnumeratorName)
                Case "HDAUDIO"
                    .VenDev = GetHWID(GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_COMPATIBLEIDS))

                    If .ClassDesc = "Other devices" Or .ClassDesc = "Monitors" Then .VenDevInfo = ReadTextFile(.VenDev)
                Case "ACPI", "USB", "HID", "PCI", "MONITOR", "DISPLAY"
                    .VenDev = GetHWID(.HardwareIDs)

                    If .ClassDesc = "Other devices" Or .ClassDesc = "Monitors" Then .VenDevInfo = ReadTextFile(.VenDev)
            End Select

        End With

        i = i + 1
        GetDevProp = True
    Loop

    Call SetupDiDestroyDeviceInfoList(hDevInfoC)
    Call SetupDiDestroyDeviceInfoList(hDevInfoI)
End Function

Public Function GetProcessorName() As String

    Dim i         As Integer
    Dim RetVal    As Long
    Dim SE        As Long
    Dim DIDC      As SP_DEVINFO_DATA
    Dim DIDI      As SP_DEVICE_INTERFACE_DATA
    Dim tGUID     As GUID
    Dim hDevInfoC As Long
    Dim hDevInfoI As Long

    Call IIDFromString(StrPtr(GUID_DEVCLASS_PROCESSOR), tGUID)
    hDevInfoC = SetupDiGetClassDevs(tGUID, vbNullString, 0, DEVICEFLAGS.DIGCF_PRESENT)
    DIDC.cbSize = Len(DIDC)
    DIDI.cbSize = Len(DIDI)

    Do
        RetVal = SetupDiEnumDeviceInfo(hDevInfoC, i, DIDC)

        If RetVal = 0 Then
            SE = Err.LastDllError

            If SE = ERROR_NO_MORE_ITEMS Then
                Exit Do
            Else
                MsgBox "Error, can't enumerate Items"
                Exit Function
            End If
        End If

        GetProcessorName = Trim$(GetSetupRegSetting(hDevInfoC, DIDC, SPDRP_FRIENDLYNAME))
        i = i + 1
    Loop

    Call SetupDiDestroyDeviceInfoList(hDevInfoC)
    Call SetupDiDestroyDeviceInfoList(hDevInfoI)
End Function

Private Function GetSetupRegSetting(ByVal hDevInfo As Long, _
                                    DeviceInfoData As SP_DEVINFO_DATA, _
                                    ByVal lPropertyName As Long, _
                                    Optional asLong As Boolean = False) As String

    Dim bDevInfo()   As Byte
    Dim lBufferSize  As Long
    Dim lRegDataType As Long

    Call SetupDiGetDeviceRegistryProperty(hDevInfo, DeviceInfoData, lPropertyName, lRegDataType, 0, 0, lBufferSize)

    If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
        ReDim bDevInfo(lBufferSize * 2 - 1)
        Call SetupDiGetDeviceRegistryProperty(hDevInfo, DeviceInfoData, lPropertyName, lRegDataType, VarPtr(bDevInfo(0)), lBufferSize, ByVal 0)

        If asLong Then
            GetSetupRegSetting = bDevInfo(0)
        Else
            GetSetupRegSetting = Trim$(Left$(StrConv(bDevInfo, vbUnicode), lBufferSize - 1))
        End If
    End If

End Function

Public Function GetDeviceByDevInst(devInst As Long, _
                                          ByRef parentDevice As HardwareDevice) As Boolean

    Dim i As Integer

    GetDeviceByDevInst = False

    For i = 0 To UBound(HW_DEVICES)

        With HW_DEVICES(i)

            If .devInst = devInst And .DevParent <> 0 Then
                parentDevice = HW_DEVICES(i)
                GetDeviceByDevInst = True
                Exit For
            End If

        End With

    Next i

End Function

Public Function GetDeviceByPNPDevice(pnpDevice As String, ByRef foundDevice As HardwareDevice) As Boolean

    Dim i As Integer

    GetDeviceByPNPDevice = False

    For i = 0 To UBound(HW_DEVICES)

        With HW_DEVICES(i)

            If .PNPDeviceId = pnpDevice Then
                foundDevice = HW_DEVICES(i)
                GetDeviceByPNPDevice = True
                Exit For
            End If

        End With

    Next i

End Function

