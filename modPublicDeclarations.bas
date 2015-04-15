Attribute VB_Name = "modPublicDeclarations"
Option Explicit

Public INIPATH                     As String
Public COMPUTER_NAME               As String
Public DOMAIN_NAME                 As String
Public USER_NAME                   As String
Public IS_ADMIN                    As Boolean
Public OS_VERSION                  As String

Public DoUnload                    As Boolean

Public REG_USERS()                 As RegisterWorkstationUser
Public REG_WORKSTATION             As RegisterWorkstation

Public SW_OPERATING_SYSTEM         As SoftwareOS
Public SW_SERVICES()               As SoftwareService
Public SW_APPLICATIONS()           As SoftwareApp
Public SW_LOCAL_USERS()            As SoftwareLocalUser
Public SW_SHARED_FOLDERS()         As SoftwareSharedFolder
Public SW_START_COMMANDS()         As SoftwareStartCommand
Public SW_SECURITY_PRODUCTS        As SoftwareSecurityProducts
Public SW_LICENSES()               As SoftwareLicense

Public HW_DEVICES()                As HardwareDevice
Public HW_CDROMS()                 As HardwareCDRomDrive
Public HW_CPU                      As HardwareCPU
Public HW_HARDDISKS()              As HardwareHardDrive
Public HW_HID                      As HardwareHID
Public HW_SOUND_DEVICES()          As HardwareSoundDevice
Public HW_LOGICAL_DRIVES()         As SoftwareHardDrivePartition
Public HW_RAM_MEMORY               As HardwareRAMMemory
Public HW_MONITORS()               As HardwareDesktopMonitor
Public HW_MOTHERBOARD              As HardwareMotherBoard
Public HW_NETWORK_ADAPTERS()       As HardwareNetworkAdapter
Public HW_PRINTERS()               As HardwarePrintDevice
Public HW_VIDEO_ADAPTERS()         As HardwareVideoAdapter

Public ClassImageList              As SP_CLASSIMAGELIST_DATA

Public Const L_TABLE_TOP           As Long = 1080
Public Const L_TABLE_LEFT          As Long = 160
