Attribute VB_Name = "modRegionalSettings"
Option Explicit

Public Const LOCALE_ILANGUAGE              As Long = &H1    'language id
Public Const LOCALE_SLANGUAGE              As Long = &H2    'localized name of lang
Public Const LOCALE_SENGLANGUAGE           As Long = &H1001 'English name of lang
Public Const LOCALE_SABBREVLANGNAME        As Long = &H3    'abbreviated lang name
Public Const LOCALE_SNATIVELANGNAME        As Long = &H4    'native name of lang
Public Const LOCALE_ICOUNTRY               As Long = &H5    'country code
Public Const LOCALE_SCOUNTRY               As Long = &H6    'localized name of country
Public Const LOCALE_SENGCOUNTRY            As Long = &H1002 'English name of country
Public Const LOCALE_SABBREVCTRYNAME        As Long = &H7    'abbreviated country name
Public Const LOCALE_SNATIVECTRYNAME        As Long = &H8    'native name of country
Public Const LOCALE_SINTLSYMBOL            As Long = &H15   'intl monetary symbol
Public Const LOCALE_IDEFAULTLANGUAGE       As Long = &H9    'def language id
Public Const LOCALE_IDEFAULTCOUNTRY        As Long = &HA    'def country code
Public Const LOCALE_IDEFAULTCODEPAGE       As Long = &HB    'def oem code page
Public Const LOCALE_IDEFAULTANSICODEPAGE   As Long = &H1004 'def ansi code page
Public Const LOCALE_IDEFAULTMACCODEPAGE    As Long = &H1011 'def mac code page

'#if(WINVER >=  &H0400)
Public Const LOCALE_SISO639LANGNAME        As Long = &H59   'ISO abbreviated language name
Public Const LOCALE_SISO3166CTRYNAME       As Long = &H5A   'ISO abbreviated country name

'#endif /* WINVER >= as long = &H0400 */
'#if(WINVER >=  &H0500)
Public Const LOCALE_SNATIVECURRNAME        As Long = &H1008 'native name of currency
Public Const LOCALE_IDEFAULTEBCDICCODEPAGE As Long = &H1012 'default ebcdic code page
Public Const LOCALE_SSORTNAME              As Long = &H1013 'sort name

'#endif /* WINVER >=  &H0500 */
'working var
Private LCID                               As Long
Private geoclass                           As Long

'SYSGEOTYPE
Private Const GEO_FRIENDLYNAME             As Long = &H8
Private Const GEO_OFFICIALNAME             As Long = &H9
Private Const GEOID_NOT_AVAILABLE          As Long = -1

'SYSGEOCLASS
Public Const GEOCLASS_NATION               As Long = 16 'only valid GeoClass value

' TIMEZONES
Private Const TIME_ZONE_ID_UNKNOWN         As Long = 1
Private Const TIME_ZONE_ID_STANDARD        As Long = 1
Private Const TIME_ZONE_ID_DAYLIGHT        As Long = 2
Private Const TIME_ZONE_ID_INVALID         As Long = &HFFFFFFFF

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 63) As Byte  'Unicode
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 63) As Byte  'Unicode
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation _
                Lib "KERNEL32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Declare Function GetUserDefaultLCID Lib "KERNEL32" () As Long

Private Declare Function GetUserGeoID Lib "KERNEL32" (ByVal geoclass As Long) As Long

Private Declare Function GetGeoInfo _
                Lib "KERNEL32" _
                Alias "GetGeoInfoA" (ByVal geoid As Long, _
                                     ByVal GeoType As Long, _
                                     lpGeoData As Any, _
                                     ByVal cchData As Long, _
                                     ByVal langid As Long) As Long

Public Declare Function EnumSystemGeoID _
               Lib "KERNEL32" (ByVal geoclass As Long, _
                               ByVal ParentGeoId As Long, _
                               ByVal lpGeoEnumProc As Long) As Long

Private Declare Function lstrlenW Lib "KERNEL32" (ByVal lpString As Long) As Long

Public Declare Function GetSystemDefaultLCID Lib "KERNEL32" () As Long
Public Declare Function GetLocaleInfo _
               Lib "KERNEL32" _
               Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                       ByVal LCType As Long, _
                                       ByVal lpLCData As String, _
                                       ByVal cchData As Long) As Long

Private Declare Function SystemParametersInfo _
                Lib "user32" _
                Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                               ByVal uParam As Long, _
                                               ByRef lpvParam As Any, _
                                               ByVal fuWinIni As Long) As Long
Private Declare Function ActivateKeyboardLayout _
                Lib "user32" (ByVal HKL As Long, _
                              ByVal Flags As Long) As Long

Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long

Private Declare Function GetKeyboardLayoutName _
                Lib "user32" _
                Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long

Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPI_SETDEFAULTINPUTLANG = 90

Const HKL = "00000409"

Public Function GetGeoFriendlyName(Optional LCID As Long) As String

    On Error Resume Next

    Dim lpGeoData As String
    Dim cchData   As Long
    Dim nRequired As Long

    geoclass = GetUserGeoID(GEOCLASS_NATION)

    If geoclass <> GEOID_NOT_AVAILABLE Then
        LCID = GetUserDefaultLCID()
    End If

    lpGeoData = ""
    cchData = 0
    nRequired = GetGeoInfo(geoclass, GEO_FRIENDLYNAME, ByVal lpGeoData, cchData, LCID)

    If (nRequired > 0) Then
        lpGeoData = Space$(nRequired)
        cchData = nRequired
        Call GetGeoInfo(geoclass, GEO_FRIENDLYNAME, ByVal lpGeoData, cchData, LCID)
        GetGeoFriendlyName = TrimNull(lpGeoData)
    End If

End Function

Public Function GetCurrentTimeZone() As String

    On Error Resume Next

    Dim tzi    As TIME_ZONE_INFORMATION
    Dim tmp    As String
    Dim dwBias As String
    Dim dwTime As String
    Dim dwSign As String

    Select Case GetTimeZoneInformation(tzi)
        Case 0
            tmp = "Cannot determine current time zone"
        Case 1
            dwBias = tzi.Bias + tzi.StandardBias
            tmp = tzi.StandardName 'tmp = "Стандартно време"
        Case 2
            dwBias = tzi.Bias + tzi.DaylightBias
            tmp = tzi.DaylightName '"Лятно часово време"
    End Select

    dwTime = Format$(CStr((dwBias \ 60) * (-1)), String$(2, "0")) & ":" & Format$(CStr(dwBias Mod 60), String$(2, "0"))
    dwSign = IIf((dwBias \ 60) < 0, "+", "-")
    GetCurrentTimeZone = TrimNull(tmp) & " (UTC" & dwSign & dwTime & ")"
End Function

Public Function GetCurrentTimeZone2() As String

    On Error Resume Next

    Dim objWMIService, colItems, objItem

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_TimeZone", , 48)

    For Each objItem In colItems
        GetCurrentTimeZone2 = GetProp(objItem.Caption)
    Next

End Function

Public Function GetCurrentCountry() As String

    On Error Resume Next

    Dim LCID As Long

    LCID = GetSystemDefaultLCID()
    GetCurrentCountry = GetUserLocaleInfo(LCID, LOCALE_SNATIVECTRYNAME) & " (" & GetUserLocaleInfo(LCID, LOCALE_ICOUNTRY) & ")"
End Function

Public Function GetCurrentLanguage() As String

    On Error Resume Next

    Dim LCID As Long

    LCID = GetSystemDefaultLCID()
    GetCurrentLanguage = GetUserLocaleInfo(LCID, LOCALE_SNATIVELANGNAME) & " (" & GetUserLocaleInfo(LCID, LOCALE_ILANGUAGE) & ")"
End Function

Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

    On Error Resume Next

    Dim sReturn As String
    Dim r       As Long

    'call the function passing the Locale type
    'variable to retrieve the required size of
    'the string buffer needed
    r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))

    'if successful..
    If r Then
        'pad the buffer with spaces
        sReturn = Space$(r)
        'and call again passing the buffer
        r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))

        'if successful (r > 0)
        If r Then
            'r holds the size of the string
            'including the terminating null
            GetUserLocaleInfo = Left$(sReturn, r - 1)
        End If
    End If

End Function

Private Function TrimNull(startstr As String) As String
    TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
End Function

Public Sub SwitchInputLanguages()
    Dim retval As Long
    
    retval = SystemParametersInfo(SPI_SETDEFAULTINPUTLANG, 0, HKL, SPIF_SENDWININICHANGE)
    
    ActivateKeyboardLayout retval, 0
End Sub
