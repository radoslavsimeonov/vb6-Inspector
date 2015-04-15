Attribute VB_Name = "modDeviceProperties"
'Option Explicit
'
'Declare Function CM_Get_DevNode_Status Lib "setupapi.dll" '                (lStatus As Long, '                lProblem As Long, '                ByVal hDevice As Long, '                ByVal dwFlags As Long) '                As Long
'
'Public Type GUID
'    Data1           As Long
'    Data2           As Integer
'    Data3           As Integer
'    Data4(0 To 7)   As Byte
'End Type
'
'Public Type SP_DEVINFO_DATA
'    cbSize          As Long
'    ClassGuid       As GUID
'    DevInst         As Long
'    Reserved        As Long
'End Type
'
'Private Const CM_PROB_DISABLED          As Long = &H16
'
'Private Const DICS_DISABLE              As Long = &H2
'Private Const DICS_ENABLE               As Long = &H1
'Private Const DICS_FLAG_GLOBAL          As Long = &H1
'Private Const DICS_FLAG_CONFIGSPECIFIC  As Long = &H2
'
'Private Const DIF_PROPERTYCHANGE        As Long = &H12
'
'Private Type SP_CLASSINSTALL_HEADER
'    cbSize              As Long
'    InstallFunction     As Long
'End Type
'
'Private Type SP_PROPCHANGE_PARAMS
'    ClassInstallHeader  As SP_CLASSINSTALL_HEADER
'    StateChange         As Long
'    Scope               As Long
'    HwProfile           As Long
'End Type
'
'Private Declare Function SetupDiSetClassInstallParams Lib "setupapi.dll" Alias "SetupDiSetClassInstallParamsA" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByRef ClassInstallParams As SP_CLASSINSTALL_HEADER, ByVal ClassInstallParamsSize As Long) As Long
'Private Declare Function SetupDiChangeState Lib "setupapi.dll" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA) As Long
'Private Declare Function SetupDiEnumDeviceInfo Lib "setupapi.dll" (ByVal DeviceInfoSet As Long, ByVal MemberIndex As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA) As Long
'Private Declare Function SetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" (ByVal ClassGuid As Long, ByVal Enumerator As Long, ByVal hwndParent As Long, ByVal flags As Long) As Long
'Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi.dll" (ByVal DeviceInfoSet As Long) As Long
'Private Declare Function SetupDiGetDeviceRegistryProperty Lib "setupapi.dll" Alias "SetupDiGetDeviceRegistryPropertyW" (ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO_DATA, ByVal Property As Long, PropertyRegDataType As Long, ByVal PropertyBuffer As Long, ByVal PropertyBufferSize As Long, RequiredSize As Long) As Long
'
'Private Const DIGCF_DEFAULT                     As Long = &H1
'Private Const DIGCF_PRESENT                     As Long = &H2
'Private Const DIGCF_ALLCLASSES                  As Long = &H4
'Private Const DIGCF_PROFILE                     As Long = &H8
'Private Const DIGCF_DEVICEINTERFACE             As Long = &H10
'
'Private Const SPDRP_ADDRESS                     As Long = (&H1C)
'Private Const SPDRP_BUSNUMBER                   As Long = (&H15)
'Private Const SPDRP_BUSTYPEGUID                 As Long = (&H13)
'Private Const SPDRP_CAPABILITIES                As Long = (&HF)
'Private Const SPDRP_CHARACTERISTICS             As Long = (&H1B)
'Private Const SPDRP_CLASS                       As Long = (&H7)
'Private Const SPDRP_CLASSGUID                   As Long = (&H8)
'Private Const SPDRP_COMPATIBLEIDS               As Long = (&H2)
'Private Const SPDRP_CONFIGFLAGS                 As Long = (&HA)
'Private Const SPDRP_DEVICEDESC                  As Long = &H0
'Private Const SPDRP_DEVTYPE                     As Long = (&H19)
'Private Const SPDRP_DRIVER                      As Long = (&H9)
'Private Const SPDRP_ENUMERATOR_NAME             As Long = (&H16)
'Private Const SPDRP_EXCLUSIVE                   As Long = (&H1A)
'Private Const SPDRP_FRIENDLYNAME                As Long = (&HC)
'Private Const SPDRP_HARDWAREID                  As Long = (&H1)
'Private Const SPDRP_LEGACYBUSTYPE               As Long = (&H14)
'Private Const SPDRP_LOCATION_INFORMATION        As Long = (&HD)
'Private Const SPDRP_LOWERFILTERS                As Long = (&H12)
'Private Const SPDRP_MAXIMUM_PROPERTY            As Long = (&H1C)
'Private Const SPDRP_MFG                         As Long = (&HB)
'Private Const SPDRP_PHYSICAL_DEVICE_OBJECT_NAME As Long = (&HE)
'Private Const SPDRP_SECURITY                    As Long = (&H17)
'Private Const SPDRP_SECURITY_SDS                As Long = (&H18)
'Private Const SPDRP_SERVICE                     As Long = (&H4)
'Private Const SPDRP_UI_NUMBER                   As Long = (&H10)
'Private Const SPDRP_UI_NUMBER_DESC_FORMAT       As Long = (&H1E)
'Private Const SPDRP_UNUSED0                     As Long = (&H3)
'Private Const SPDRP_UNUSED1                     As Long = (&H5)
'Private Const SPDRP_UNUSED2                     As Long = (&H6)
'Private Const SPDRP_UPPERFILTERS                As Long = (&H11)
'
'Public Type Device1
'    DevInst                         As Long
'    ClassGuid                       As GUID
'    ADDRESS                         As Long
'    BusNumber                       As Long
'    BUSTYPEGUID                     As GUID
'    CAPABILITIES                    As Long
'    CHARACTERISTICS                 As Long
'    Class                           As String
'    strClassGuid                    As String
'    COMPATIBLEIDS                   As String
'    CONFIGFLAGS                     As Long
'    'DEVICE_POWER_DATA As CM_POWER_DATA
'    '(Windows XP and later) The function retrieves a CM_POWER_DATA
'    'structure containing the device's power management information.
'    DeviceDesc                      As String
'    DevType                         As Long
'    Driver                          As String
'    ENUMERATOR_NAME                 As String
'    Exclusive                       As Long
'    FriendlyName                    As String
'    HardwareId                      As String
'    'LEGACYBUSTYPE As INTERFACE_TYPE
'    'The function retrieves the device's legacy bus type as an
'    'INTERFACE_TYPE value (defined in wdm.h and ntddk.h).
'    LOCATION_INFORMATION            As String
'    LOWERFILTERS                    As String
'    MFG                             As String
'    PHYSICAL_DEVICE_OBJECT_NAME     As String
'    REMOVAL_POLICY                  As Long
'    REMOVAL_POLICY_HW_DEFAULT       As Long
'    REMOVAL_POLICY_OVERRIDE         As Long
'    'SPDRP_SECURITY
'    'The function retrieves a SECURITY_DESCRIPTOR structure for a device.
'    SECURITY_SDS                    As String
'    SERVICE                         As String
'    UI_NUMBER                       As Long
'    UI_NUMBER_DESC_FORMAT           As String
'    UPPERFILTERS                    As String
'    ISENABLED                       As Boolean
'    InstanceN                       As Long
'End Type
'
'Private cpuguid As GUID
'
'Public Function EnumDevices() As Device1()
'
'    Dim dwInstance  As Long
'    Dim hDevInfo    As Long
'    Dim DevInfo     As SP_DEVINFO_DATA
'    Dim arrTemp()   As Device1
'
'
'    hDevInfo = SetupDiGetClassDevs(ByVal 0&, ByVal 0&, ByVal 0&, DIGCF_PRESENT Or DIGCF_ALLCLASSES)
'
'    DevInfo.cbSize = LenB(DevInfo)
'
'    If hDevInfo = -1 Then Exit Function
'
'    Do
'        If SetupDiEnumDeviceInfo(hDevInfo, dwInstance, DevInfo) = 0 Then
'            SetupDiDestroyDeviceInfoList hDevInfo
'            Exit Do
'        End If
'
'        ReDim Preserve arrTemp(dwInstance)
'
'        arrTemp(dwInstance) = GetDeviceProperties(hDevInfo, DevInfo)
'
'        arrTemp(dwInstance).ClassGuid = DevInfo.ClassGuid
'
'        arrTemp(dwInstance).InstanceN = dwInstance
'
'        dwInstance = dwInstance + 1
'
'    Loop
'
'    EnumDevices = arrTemp
'End Function
'
'Public Function GetDeviceProperties(hDevInfo As Long, DevInfoData As SP_DEVINFO_DATA) As Device1
'
'    Dim RegDataType As Long
'    Dim dwReqLen As Long
'    Dim strBuffer As String
'    Dim status, Problem As Long
'
'    GetDeviceProperties.DevInst = DevInfoData.DevInst
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_ADDRESS, RegDataType, ByVal VarPtr(GetDeviceProperties.ADDRESS), 4&, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_BUSNUMBER, RegDataType, ByVal VarPtr(GetDeviceProperties.BusNumber), 4&, dwReqLen
'
'    'SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_BUSTYPEGUID, regDataType, ByVal 0&, 0&, dwReqLen
'    'strBuffer = Space$(dwReqLen-1)
'    '
'    'SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_ADDRESS, regDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_CAPABILITIES, RegDataType, ByVal VarPtr(GetDeviceProperties.CAPABILITIES), 4&, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_CHARACTERISTICS, RegDataType, ByVal VarPtr(GetDeviceProperties.CHARACTERISTICS), 4&, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_CLASS, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_CLASS, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.Class = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_CLASSGUID, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_CLASSGUID, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.strClassGuid = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_COMPATIBLEIDS, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_COMPATIBLEIDS, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.COMPATIBLEIDS = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_CONFIGFLAGS, RegDataType, ByVal VarPtr(GetDeviceProperties.CONFIGFLAGS), 4&, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_DEVICEDESC, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_DEVICEDESC, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.DeviceDesc = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_DEVTYPE, RegDataType, ByVal VarPtr(GetDeviceProperties.DevType), 4&, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_DRIVER, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_DRIVER, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.Driver = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_ENUMERATOR_NAME, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_ENUMERATOR_NAME, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.ENUMERATOR_NAME = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_EXCLUSIVE, RegDataType, ByVal VarPtr(GetDeviceProperties.Exclusive), 4&, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_FRIENDLYNAME, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_FRIENDLYNAME, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.FriendlyName = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_HARDWAREID, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_HARDWAREID, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.HardwareId = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_LOCATION_INFORMATION, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_LOCATION_INFORMATION, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.LOCATION_INFORMATION = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_LOWERFILTERS, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_LOWERFILTERS, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.LOWERFILTERS = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_MFG, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_MFG, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.MFG = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_PHYSICAL_DEVICE_OBJECT_NAME, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_PHYSICAL_DEVICE_OBJECT_NAME, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.PHYSICAL_DEVICE_OBJECT_NAME = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    'SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP, regDataType, ByVal VarPtr(GetDeviceProperties), 4&, dwReqLen
'
'    'SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_, regDataType, ByVal VarPtr(GetDeviceProperties), 4&, dwReqLen
'
'    'SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_, regDataType, ByVal VarPtr(GetDeviceProperties), 4&, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_SECURITY_SDS, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_SECURITY_SDS, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.SECURITY_SDS = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_SERVICE, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_SERVICE, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.SERVICE = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_UI_NUMBER, RegDataType, ByVal VarPtr(GetDeviceProperties.UI_NUMBER), 4&, dwReqLen
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_UI_NUMBER_DESC_FORMAT, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_UI_NUMBER_DESC_FORMAT, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.UI_NUMBER_DESC_FORMAT = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_UPPERFILTERS, RegDataType, ByVal 0&, 0&, dwReqLen
'    strBuffer = Space$(dwReqLen)
'    SetupDiGetDeviceRegistryProperty hDevInfo, DevInfoData, SPDRP_UPPERFILTERS, RegDataType, ByVal StrPtr(strBuffer), dwReqLen, dwReqLen
'    GetDeviceProperties.UPPERFILTERS = Trim$(Replace(strBuffer, vbNullChar, ""))
'
'    CM_Get_DevNode_Status status, Problem, DevInfoData.DevInst, 0
'    GetDeviceProperties.ISENABLED = ((Problem And CM_PROB_DISABLED) = 0)
'End Function
'
'Public Function EnableDevice(lEnumerator As Long, '                              ByVal bEnable As Boolean) As Boolean
'
'    Dim changeParams    As SP_PROPCHANGE_PARAMS
'    Dim hDevInfo        As Long
'    Dim DevInfo As SP_DEVINFO_DATA
'
'    hDevInfo = SetupDiGetClassDevs(ByVal 0&, ByVal 0&, ByVal 0&, DIGCF_PRESENT Or DIGCF_ALLCLASSES)
'
'    DevInfo.cbSize = LenB(DevInfo)
'
'    If SetupDiEnumDeviceInfo(hDevInfo, lEnumerator, DevInfo) = 0 Then
'        SetupDiDestroyDeviceInfoList hDevInfo
'        Exit Function
'    End If
'
'    With changeParams
'        .ClassInstallHeader.cbSize = LenB(.ClassInstallHeader)
'        .ClassInstallHeader.InstallFunction = DIF_PROPERTYCHANGE
'
'        .Scope = DICS_FLAG_CONFIGSPECIFIC 'DICS_FLAG_GLOBAL Or
'        .StateChange = IIf(bEnable, DICS_ENABLE, DICS_DISABLE)
'        .HwProfile = 0
'    End With
'
'    If SetupDiSetClassInstallParams(hDevInfo, '                                    DevInfo, '                                    changeParams.ClassInstallHeader, '                                    LenB(changeParams)) = 1 Then
'        EnableDevice = (SetupDiChangeState(hDevInfo, DevInfo) = 1)
'    End If
'End Function
'
