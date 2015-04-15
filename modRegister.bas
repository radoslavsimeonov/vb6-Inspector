Attribute VB_Name = "modRegister"
Option Explicit

Public Type RegisterHDD
    RegistryNo          As String
    InventaryNo         As String
    InventaryDate       As String
    InventarySerialNum  As String
    AdminName           As String
    AdminDate           As String
    AdminSticker        As String
    RegistryName        As String
    RegistryDate        As String
    RegistrySticker     As String
End Type

Public Type RegisterCase
    CaseStickers()       As String
End Type

Public Type RegisterUserService
    ServiceName         As String
    ServiceAddress      As String
    ServiceValue        As String
    ServicePeriod       As String
End Type

Public Type RegisterWorkstationUser
    username            As String
    Rank                As String
    FirstName           As String
    MiddleName          As String
    SurName             As String
    IsOwner             As Boolean
    IsNavy              As Boolean
    Email               As String
    Function            As String
    Department          As String
    Building            As String
    Room                As String
    Phone               As String
    UserServices()      As RegisterUserService
End Type

Public Type RegisterWorkstation
    CaseStickers()      As String
    InventaryNo         As String
    InventaryDate       As String
    BookNo              As String
    BookDate            As String
    Classification      As String
    SocketSKS           As String
End Type
