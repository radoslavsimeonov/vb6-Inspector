Attribute VB_Name = "modHardware"
Option Explicit

Public Type HardwareHardDrive
    Index               As Integer
    SerialNumber        As String
    Model               As String
    InterfaceType       As String
    Family              As String
    RealSize            As Double
    Size                As Double
    DeviceID            As String
    Mode                As String
    PNPDeviceId         As String
    Removable           As Boolean
    Paritions()         As SoftwareHardDrivePartition
    Registry            As RegisterHDD
End Type

Public Type HardwareDesktopMonitor
    Manufacturer      As String
    Model             As String
    PNPDeviceId       As String
    Size              As String
    HSize             As String
    VSize             As String
    SerialNumber      As String
    ModelId           As String
    ManufacureDate    As String
    VideoInput        As String
    AspectRatio       As String
End Type

Public Type HardwareCDRomDrive
    Model           As String
    Interface       As String
    Manufacturer    As String
    ReadMedia       As String
    WriteMedia      As String
    SerialNumber    As String
    Description     As String
    FirmWare        As String
    PNPDeviceId     As String
    DriveLetter     As String
    Virtual         As Boolean
End Type

Public Type HardwareCPU
    Manufacturer    As String
    Model           As String
    Architecture    As String
    ClockSpeed      As Integer
    Cores           As Integer
    LogicalProcessors      As Integer
    Socket          As Integer
    PNPDeviceId     As String
End Type

Public Type HardwareRAMBank
    BankLabel       As String
    Capacity        As Double
    Location        As String
    Speed           As Double
    Type            As Integer
    FormFactor      As Integer
End Type

Public Type HardwareRAMMemory
    Banks()         As HardwareRAMBank
    TotalMemory     As Double
End Type

Public Type HardwareNetworkAdapter
    GUID            As String
    AdapterType     As Integer
    Model           As String
    Manufacturer    As String
    MACAddress      As String
    PNPDeviceId     As String
    Speed           As Double
    Configuration   As SoftwareNetworkAdapterConfig
End Type

Public Type HardwareSoundDevice
    Model           As String
    Manufacturer    As String
    PNPDeviceId     As String
End Type

Public Type HardwareVideoAdapter
    Manufacturer    As String
    Model           As String
    PNPDeviceId     As String
    VideoRAM        As Double
End Type

Public Type HardwareHID
    Mouse           As String
    Keyboard        As String
End Type

Public Type HWID
    Type            As String
    dev             As String
    VEN             As String
    SUBSYS          As String
    SubSys1         As String
    SubSys2         As String
    REV             As String
End Type

Public Type HWDetails
    Vendor          As String
    Chip            As String
End Type

Public Type HardwareDevice
    Index           As Long
    devInst         As Long
    ClassGuid       As String
    ClassDesc       As String
    BusNumber       As String
    BusName         As String
    Class           As String
    DeviceDesc      As String
    DeviceType      As Long
    Driver          As String
    EnumeratorName  As String
    FriendlyName    As String
    HardwareIDs     As String
    PNPDeviceId     As String
    Location        As String
    Manufacturer    As String
    Enabled         As Boolean
    DeviceStatus    As String
    ErrorCode       As Integer
    Hidden          As Boolean
    DevParent       As Long
    CanDisable      As Boolean
    CanUninstall    As Boolean
    Removable       As Boolean
    VenDev          As HWID
    VenDevInfo      As HWDetails
End Type

Public Type HardwareMotherBoard
    SystemModel     As String
    SystemMfg       As String
    Model           As String
    Manufacturer    As String
    SerialNumber    As String
    MemorySlots     As Integer
    MemoryMax       As Double
    CPUSocket       As String
    ChassisType     As Integer
    ChassisSN       As String
    BIOS            As String
    Floppy()        As String
    Ports()         As String
    'USB()           As String
    'PCIDevices()    As HardwareDevice
End Type

Public Type HardwarePrintDevice
    Model           As String
    Manufacturer    As String
    DeviceID        As String
    ErrorState      As String
    DriverName      As String
    PortName        As String
    IPAddress       As String
    hostname        As String
    ShareName       As String
    ConnectionStat  As String
    IsOnline        As Boolean
    IsDefault       As Boolean
    IsLocal         As Boolean
    IsNetwork       As Boolean
    IsShared        As Boolean
    InventaryNo     As String
End Type
