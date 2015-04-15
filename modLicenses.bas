Attribute VB_Name = "swLicenses"
Option Explicit

Private Type tProductLincese
    ProductName     As String
    ProductPath     As String
    Offset          As Long
    SearchKey       As String
End Type

Private arrLic()   As SoftwareLicense
Private idx        As Integer

Public Function GetSoftwareLicenses() As SoftwareLicense()
    EnuumSoftwareLicenses
    If Is64bit Then EnuumSoftwareLicenses (KEY_READ64)
    GetSoftwareLicenses = arrLic
End Function


Public Function GetSoftwareLicenses1() As SoftwareLicense()
    
    Dim SEARCH_KEY As String
    Dim arrDPIDBytes, arrGUIDKeys, GUIDKey
    Dim iValues, arrDPID
    Dim x As Integer
    Dim arrLic() As SoftwareLicense
    Dim idx As Integer
    Dim Products(6) As tProductLincese
    
    Dim oReg
    
    iValues = Array()
    
    Products(0).ProductName = "Microsoft Windows Product Key"
    Products(0).ProductPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    Products(1).ProductName = "Microsoft Office XP"
    Products(1).ProductPath = "SOFTWARE\Microsoft\Office\10.0\Registration"
    Products(2).ProductName = "Microsoft Office 2003"
    Products(2).ProductPath = "SOFTWARE\Microsoft\Office\11.0\Registration"
    Products(3).ProductName = "Microsoft Office 2007"
    Products(3).ProductPath = "SOFTWARE\Microsoft\Office\12.0\Registration"
    Products(4).ProductName = "Microsoft Office 2010"
    Products(4).ProductPath = "SOFTWARE\Microsoft\Office\14.0\Registration"
    Products(4).Offset = &H328
    Products(5).ProductName = "Microsoft Office 2013"
    Products(5).ProductPath = "SOFTWARE\Microsoft\Office\15.0\Registration"
    Products(5).Offset = &H328
    Products(6).ProductName = "Microsoft Exchange Product Key"
    Products(6).ProductPath = "SOFTWARE\Microsoft\Exchange\Setup"

    ' <--------------- Open Registry Key and populate binary data into an array -------------------------->
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
    For x = LBound(Products, 1) To UBound(Products, 1)
        SEARCH_KEY = IIf(Products(x).SearchKey = vbNullString, "DigitalProductID", Products(x).SearchKey)
    
        oReg.GetBinaryValue HKEY_LOCAL_MACHINE, Products(x).ProductPath, SEARCH_KEY, arrDPIDBytes
        
        ReDim Preserve arrLic(idx)
        
        If Not IsNull(arrDPIDBytes) Then
            arrLic(idx).Product = Products(x).ProductName
            arrLic(idx).CDKey = ConvertToKey(arrDPIDBytes, Products(x).Offset)
        Else
            oReg.EnumKey HKEY_LOCAL_MACHINE, Products(x).ProductPath, arrGUIDKeys

            If Not IsNull(arrGUIDKeys) Then

                For Each GUIDKey In arrGUIDKeys
                    oReg.GetBinaryValue HKEY_LOCAL_MACHINE, Products(x).ProductPath & "\" & GUIDKey, SEARCH_KEY, arrDPIDBytes

                    If Not IsNull(arrDPIDBytes) Then
                        arrLic(idx).Product = Products(x).ProductName
                        'arrLic(idx).Product = Get.reg "ConvertToEdition"
                        
                        arrLic(idx).CDKey = ConvertToKey(arrDPIDBytes, Products(x).Offset)
                    End If

                Next

            End If
        End If
        
        If NotEmpty(arrLic(idx).Product) And NotEmpty(arrLic(idx).Product) Then _
            idx = idx + 1
    Next
    
    If idx > 0 Then _
        ReDim Preserve arrLic(idx - 1)
    
    GetSoftwareLicenses1 = arrLic

End Function

Function ConvertToKey(regKey, Optional KeyOffset As Long = 52)
On Error Resume Next

    'Const KeyOffset = 52
    'Const KeyOffset = &H328
    
    Dim isWin8 As Long
    Dim j, y, Cur, Last As Long
    Dim Chars, Insert, keypart1 As String
    Dim winKeyOutput As String
    Dim a, b, c, d, e As String
    
    If KeyOffset = 0 Then KeyOffset = 52
    
    isWin8 = (regKey(KeyOffset + 14) \ 6) And 1
    regKey(KeyOffset + 14) = (regKey(KeyOffset + 14) And &HF7) Or ((isWin8 And 2) * 4)
    
    j = 24
    
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    
    Do
        Cur = 0
        y = 14
        
        Do
            Cur = Cur * 256
            Cur = regKey(y + KeyOffset) + Cur
            regKey(y + KeyOffset) = (Cur \ 24)
            Cur = Cur Mod 24
            y = y - 1
        Loop While y >= 0
    
        j = j - 1
        
        winKeyOutput = Mid(Chars, Cur + 1, 1) & winKeyOutput
        
        Last = Cur
    
    Loop While j >= 0
    
    If (isWin8 = 1) Then
        keypart1 = Mid(winKeyOutput, 2, Last)
        Insert = "N"
        winKeyOutput = Replace(winKeyOutput, keypart1, keypart1 & Insert, 2, 1, 0)
        If Last = 0 Then winKeyOutput = Insert & winKeyOutput
    End If
    
    a = Mid(winKeyOutput, 1, 5)
    b = Mid(winKeyOutput, 6, 5)
    c = Mid(winKeyOutput, 11, 5)
    d = Mid(winKeyOutput, 16, 5)
    e = Mid(winKeyOutput, 21, 5)
    
    ConvertToKey = a & "-" & b & "-" & c & "-" & d & "-" & e

End Function

Private Sub EnuumSoftwareLicenses(Optional ByVal OSBits As Long = KEY_READ32)
    
    Dim KeyName    As String   ' receives name of each subkey
    Dim keylen     As Long     ' length of keyname
    Dim classname  As String   ' receives class of each subkey
    Dim classlen   As Long     ' length of classname
    Dim lastwrite  As FILETIME ' receives last-write-to time, but we ignore it here
    Dim hKey       As Long     ' handle to the HKEY_LOCAL_MACHINE\Software key
    Dim RetVal     As Long     ' function's return value
    Dim Index      As Long     ' counter variable for index
    Dim x          As Integer
    
    Dim SEARCH_KEY As String
    Dim arrDPIDBytes

    Dim Products(6) As tProductLincese
    
    Products(0).ProductName = "Microsoft Windows Product Key"
    Products(0).ProductPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    Products(1).ProductName = "Microsoft Office XP"
    Products(1).ProductPath = "SOFTWARE\Microsoft\Office\10.0\Registration"
    Products(2).ProductName = "Microsoft Office 2003"
    Products(2).ProductPath = "SOFTWARE\Microsoft\Office\11.0\Registration"
    Products(3).ProductName = "Microsoft Office 2007"
    Products(3).ProductPath = "SOFTWARE\Microsoft\Office\12.0\Registration"
    Products(4).ProductName = "Microsoft Office 2010"
    Products(4).ProductPath = "SOFTWARE\Microsoft\Office\14.0\Registration"
    Products(4).Offset = &H328
    Products(5).ProductName = "Microsoft Office 2013"
    Products(5).ProductPath = "SOFTWARE\Microsoft\Office\15.0\Registration"
    Products(5).Offset = &H328
    Products(6).ProductName = "Microsoft Exchange Product Key"
    Products(6).ProductPath = "SOFTWARE\Microsoft\Exchange\Setup"

'    Products(7).ProductName = "Microsoft Windows Product Key DPid4"
'    Products(7).ProductPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
'    Products(7).SearchKey = "DigitalProductId4"
'    Products(7).Offset = &H328
'    Products(8).ProductName = "Microsoft Windows Product Key Def"
'    Products(8).ProductPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\DefaultProductKey"
'    Products(9).ProductName = "Microsoft Windows Product Key Def_DPid4"
'    Products(9).ProductPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\DefaultProductKey"
'    Products(9).SearchKey = "DigitalProductId4"
'    Products(9).Offset = &H328

    For x = LBound(Products) To UBound(Products)
        SEARCH_KEY = IIf(Products(x).SearchKey = vbNullString, "DigitalProductID", Products(x).SearchKey)
        
        arrDPIDBytes = QueryValue(HKEY_LOCAL_MACHINE, Products(x).ProductPath, SEARCH_KEY, OSBits)
        
        If Not IsEmpty(arrDPIDBytes) Then
            ReDim Preserve arrLic(idx)
            arrLic(idx).Product = Products(x).ProductName
            arrLic(idx).CDKey = ConvertToKey(arrDPIDBytes, Products(x).Offset)
            idx = idx + 1
        Else
            Index = 0  ' initial index value
            RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Products(x).ProductPath, 0, OSBits, hKey)

            If RetVal = ERROR_SUCCESS Then
                While RetVal = 0
                    KeyName = Space$(255)
                    keylen = 255
                    
                    RetVal = RegEnumKeyEx(hKey, Index, KeyName, keylen, ByVal 0, vbNullString, 0&, lastwrite)
        
                    If RetVal = ERROR_SUCCESS Then
                        KeyName = Left$(KeyName, keylen)
                        
                        arrDPIDBytes = QueryValue(HKEY_LOCAL_MACHINE, Products(x).ProductPath & "\" & KeyName, SEARCH_KEY, OSBits)
                        
                        If Not IsEmpty(arrDPIDBytes) Then
                            ReDim Preserve arrLic(idx)
                            arrLic(idx).Product = Products(x).ProductName
                            arrLic(idx).CDKey = ConvertToKey(arrDPIDBytes, Products(x).Offset)
                            idx = idx + 1
                        End If
        
                        
                    End If
                    Index = Index + 1
                Wend
            End If
        End If

    Next x

    RetVal = RegCloseKey(hKey)

    If idx > 0 Then ReDim Preserve arrLic(idx - 1)

    Exit Sub
Err:
        MsgBox "Error " & Err.Description
    End Sub

