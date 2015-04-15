Attribute VB_Name = "modSoftware"
Option Explicit

Public Type SoftwareHardDrivePartition
    Caption            As String
    FileSystem         As String
    VolumeName         As String
    Size               As Double
    FreeSpace          As Double
End Type

Public Type SoftwareNetworkAdapterConfig
    IP()            As String
    mask()          As String
    GateWay()       As String
    DNS()           As String
    DHCPEnabled     As Boolean
    DHCPServer      As String
    ConnectionStatus As Integer
    NetConnectionID As String
End Type

Public Type SoftwareApp
    AppName             As String
    Publisher           As String
    InstallLocation     As String
    Version             As String
    UninstallString     As String
    ModifyString        As String
    IsUpdate            As Boolean
    InstalledOn         As String
    AppBits             As String
End Type

Public Type SoftwareService
    AcceptPause     As Boolean
    AcceptStop      As Boolean
    Caption         As String
    Description     As String
    Name            As String
    PathName        As String
    StartMode       As String
    state           As String
    Manufacturer    As String
End Type

Public Type SoftwareLocalUser
    Index                 As Long
    Name                  As String
    FullName              As String
    Description           As String
    UserComment           As String
    CannotChangePassword     As Integer
    PasswordNeverExpires  As Integer
    PasswordExpired       As Integer
    AccountDisabled       As Integer
    AccountLocked         As Integer
    LastLogin             As Date
    Group                 As String
    Groups                As String
    MaxPasswordLen        As Integer
    Password              As String
End Type

Public Type SoftwareLocalGroup
    GroupName             As String
    Description           As String
    Members()             As String
End Type

Public Type SoftwareSharedFolder
    ShareName             As String
    FolderPath            As String
    Description           As String
    MaximumAllowed        As Long
End Type

Public Type SoftwareStartCommand
    CommandName         As String
    CommandNameShort    As String
    Command             As String
    Location            As String
    Architecture        As RegReadWOW64Constants
    UserRange           As String
    Vendor              As String
End Type

Public Type SoftwareSecurityProduct
    ProductName         As String
    Publisher           As String
    ProductVersion      As String
    ProductGUID         As String
    RTPStatus           As Integer
    UpToDate            As Integer
    Enabled             As Integer
    PathToProduct       As String
    ProductState        As String
End Type

Public Type SoftwareSecurityProducts
    AntiVirus()         As SoftwareSecurityProduct
    Firewall()          As SoftwareSecurityProduct
    Spyware()           As SoftwareSecurityProduct
End Type

Public Type SoftwareLicense
    Product             As String
    CDKey               As String
End Type

Public Type SoftwareOS
    Caption             As String
    Architecture        As String
    BuildNumber         As String
    CodeSet             As String
    CSDVersion          As String
    CountryCode         As String
    CSName              As String
    Domain              As String
    CurrentTimeZone     As String
    InstallDate         As String
    Locale              As String
    Organization        As String
    OSLanugage          As Long
    Primary             As Boolean
    ProductType         As String
    RegisteredUser      As String
    SPMajorVersion      As Integer
    SPMinorVersion      As Integer
    SystemDrive         As String
    SystemDirectory     As String
    WindowsDirecorty    As String
    Version             As String
    ActivationStatus    As Integer
End Type
