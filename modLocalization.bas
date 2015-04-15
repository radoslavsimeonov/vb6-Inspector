Attribute VB_Name = "modLocalization"
Option Explicit

Public Enum StringIds
    l_Manufacturer = 101
    l_Model
    l_SerialNumber
    l_Description
    l_FirmWare
    l_ReadMedia
    l_WriteMedia
    l_DriveLetter
    l_Architecture
    l_Speed
    l_Cores
    l_Processor
    l_Socket
    l_Family
    l_Interface
    l_Partition
    l_Partitions
    l_free
    l_Dimensions
    l_AspectRatio
    l_VideoInput
    l_ManufactureDate
    l_DeviceType
    l_ConnectionStatus
    l_ConnectionName
    l_Memory
    l_MotherBoard
    l_ChassisType
    l_ChassisSN
    l_MaxMemory
    l_MemorySlots
    l_InstalledMemory
    l_Mouse
    l_Keyboard
    l_Size
    l_Mode
    l_status
    l_Ports = 138
    l_ShowHiddenDevices
    l_Disable
    l_Enable
    l_Computer
    l_DisplayAdapters
    l_DVDDevices
    l_IDEDevices
    l_Keyboards
    l_Mouses
    l_Monitors
    l_NetworkDevices
    l_OtherDevices
    l_PortableDevices
    l_PortComLpt
    l_Processors
    l_SoundVideoControllers
    l_SystemDevices
    l_USBControllers
    l_AudioInOut
    l_FloppyControllers
    l_DiskDrives
    l_HIDDevices
    l_IEEE1394Controller
    l_FloppyDrive
    l_SCSIControllers
    l_StorageVolumes
    l_NonPNPDevices
    l_MultifunctionAdapters
    l_VolumeSnapshot
    l_ParentDevice
    l_PrintQueues
    l_SoftwareDevices
    l_ImagingDevices
    l_IEEE1284_4Printers
    l_Printers
    l_Default
    l_PortName
    l_Driver
    l_Resolution
    l_Local
    l_Network
    l_Shared
    l_SharedName
    l_IPAddress
    l_HostName
    l_Server
    l_Yes
    l_No
    l_Capabilities
    l_Active
    l_HostConnection
    l_Virtual
    l_SystemMfg
    l_SystemModel
End Enum

Public Function locstr(ByVal ID As StringIds)

   locstr = LoadResString(ID)
   
End Function
