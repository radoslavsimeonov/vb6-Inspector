Attribute VB_Name = "mNetShare"
Option Explicit

Private Const ERROR_MORE_DATA = 234
Private Const SHPWLEN = 8

Enum enmSTYPE
    STYPE_DISKTREE = 0
    STYPE_PRINTQ = 1
    STYPE_DEVICE = 2
    STYPE_IPC = 3
    STYPE_SPECIAL = &H80000000
End Enum

Enum enmACCESS
    ACCESS_NONE = 0
    ACCESS_READ = &H1
    ACCESS_WRITE = &H2
    ACCESS_CREATE = &H4
    ACCESS_EXEC = &H8
    ACCESS_DELETE = &H10
    ACCESS_ATRIB = &H20
    ACCESS_PERM = &H40
    ACCESS_GROUP = &H8000
    ACCESS_ALL = (ACCESS_READ Or ACCESS_WRITE Or ACCESS_CREATE Or ACCESS_EXEC Or ACCESS_DELETE Or ACCESS_PERM)
End Enum

Private Type SHARE_INFO_0
    shi0_netname As Long
End Type

Private Type SHARE_INFO_1
    shi1_netname As Long
    shi1_type As enmSTYPE
    shi1_remark As Long
End Type

Private Type SHARE_INFO_2
    shi2_netname As Long
    shi2_type As enmSTYPE
    shi2_remark As Long
    shi2_permissions As enmACCESS
    shi2_max_uses As Long
    shi2_current_uses As Long
    shi2_path As Long
    shi2_passwd As Long
End Type

Private Type SHARE_INFO_502
    shi502_netname As Long
    shi502_type As enmSTYPE
    shi502_remark As Long
    shi502_permissions As enmACCESS
    shi502_max_uses As Long
    shi502_current_uses As Long
    shi502_path As Long
    shi502_passwd As Long
    shi502_reserved As Long
    shi502_security_descriptor As Long
End Type

Private Type SHARE_INFO_1004
    shi1004_remark As Long
End Type

Private Type SHARE_INFO_1006
    shi1006_max_uses As Long
End Type

Private Type SHARE_INFO_1501
    shi1501_reserved As Long
    shi1501_security_descriptor As Long
End Type


Type SHARE_INFO_0_VB
    vb_shi0_netname As String
End Type

Type SHARE_INFO_1_VB
    vb_shi1_netname As String
    vb_shi1_type As enmSTYPE
    vb_shi1_remark As String
End Type

Type SHARE_INFO_2_VB
    vb_shi2_netname As String
    vb_shi2_type As enmSTYPE
    vb_shi2_remark As String
    vb_shi2_permissions As enmACCESS
    vb_shi2_max_uses As Long
    vb_shi2_current_uses As Long
    vb_shi2_path As String
    vb_shi2_passwd As String
End Type

Type SHARE_INFO_502_VB
    vb_shi502_netname As String
    vb_shi502_type As enmSTYPE
    vb_shi502_remark As String
    vb_shi502_permissions As enmACCESS
    vb_shi502_max_uses As Long
    vb_shi502_current_uses As Long
    vb_shi502_path As String
    vb_shi502_passwd As String
    vb_shi502_reserved As Long
    vb_shi502_security_descriptor As Long
End Type


Private Declare Function NetShareGetInfo Lib "netapi32.dll" (ByVal servername As String, _
                                                             ByVal netname As String, _
                                                             ByVal level As Long, _
                                                             bufptr As Any) As Long

Private Declare Function NetShareEnum Lib "netapi32.dll" (ByVal servername As String, _
                                                          ByVal level As Long, _
                                                          bufptr As Any, _
                                                          prefmaxlen As Long, _
                                                          entriesread As Long, _
                                                          totalentries As Long, _
                                                          resume_handle As Any) As Long

Private Declare Function NetShareAdd Lib "netapi32.dll" (ByVal servername As String, _
                                                         ByVal level As Long, _
                                                         bufptr As Any, _
                                                         parm_err As Long) As Long

Private Declare Function NetShareDel Lib "netapi32.dll" (ByVal servername As String, _
                                                         ByVal netname As String, _
                                                         ByVal reserved As Long) As Long

Private Declare Function NetShareCheck Lib "netapi32.dll" (ByVal servername As String, _
                                                           ByVal device As String, _
                                                           lngType As Long) As Long

Private Declare Function NetShareSetInfo Lib "netapi32.dll" (ByVal servername As String, _
                                                             ByVal netname As String, _
                                                             ByVal level As Long, _
                                                             bufptr As Any, _
                                                             parm_err As Long) As Long

Private Declare Function NetAPIBufferAllocate Lib "netapi32.dll" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long
Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private strNetShareServer As String


Public Property Get NetShareServer() As String
    NetShareServer = strNetShareServer
End Property

Public Property Let NetShareServer(ByVal vNewValue As String)
    Dim strLocal As String

    If (Len(vNewValue) < 2 And vNewValue <> vbNullString) Then Exit Property
    If vNewValue = vbNullString Then
        vNewValue = Space$(255)
        If GetComputerName(vNewValue, Len(vNewValue)) <> 0 Then
            vNewValue = Left$(vNewValue, InStr(vNewValue, Chr$(0)) - 1)
        Else
            vNewValue = ""
        End If
    End If
    If Mid$(vNewValue, 1, 2) <> "\\" Then vNewValue = "\\" & vNewValue
    strNetShareServer = UCase$(vNewValue)
End Property

Public Property Get NetShareLocalCheck() As Boolean
    Dim str As String
    NetShareLocalCheck = True
    If strNetShareServer = "" Then Exit Property
    str = Space$(255)
    If GetComputerName(str, Len(str)) <> 0 Then
        str = "\\" & Left$(str, InStr(str, Chr$(0)) - 1)
        If UCase$(strNetShareServer) <> UCase$(str) Then
            NetShareLocalCheck = False
        End If
    End If
End Property

Public Function ShareGetInfo2(ByVal strShareName As String, tSI2_VB As SHARE_INFO_2_VB) As Long
    Dim lngBuffer As Long
    Dim tSI2 As SHARE_INFO_2
    
    ShareGetInfo2 = NetShareGetInfo(StrConv(strNetShareServer, vbUnicode), _
                                    StrConv(strShareName, vbUnicode), _
                                    2, _
                                    lngBuffer)
    If ShareGetInfo2 <> 0 Then Exit Function
    MoveMemory tSI2, ByVal lngBuffer, Len(tSI2)
    With tSI2_VB
        .vb_shi2_netname = GetPtrToStrA(tSI2.shi2_netname)
        .vb_shi2_type = tSI2.shi2_type
        .vb_shi2_remark = GetPtrToStrA(tSI2.shi2_remark)
        .vb_shi2_permissions = tSI2.shi2_permissions
        .vb_shi2_max_uses = tSI2.shi2_max_uses
        .vb_shi2_current_uses = tSI2.shi2_current_uses
        .vb_shi2_path = GetPtrToStrA(tSI2.shi2_path)
        .vb_shi2_passwd = GetPtrToStrA(tSI2.shi2_passwd)
    End With
    Call NetAPIBufferFree(lngBuffer)
End Function

Public Function ShareGetInfo502(ByVal strShareName As String, _
                                tSI502_VB As SHARE_INFO_502_VB) As Long
    Dim lngBuffer As Long
    Dim tSI502 As SHARE_INFO_502
    
    ShareGetInfo502 = NetShareGetInfo(StrConv(strNetShareServer, vbUnicode), _
                                      StrConv(strShareName, vbUnicode), _
                                      502, _
                                      lngBuffer)
    If ShareGetInfo502 <> 0 Then Exit Function
    MoveMemory tSI502.shi502_netname, ByVal lngBuffer, Len(tSI502)
    With tSI502_VB
        .vb_shi502_netname = GetPtrToStrA(tSI502.shi502_netname)
        .vb_shi502_type = tSI502.shi502_type
        .vb_shi502_remark = GetPtrToStrA(tSI502.shi502_remark)
        .vb_shi502_permissions = tSI502.shi502_permissions
        .vb_shi502_max_uses = tSI502.shi502_max_uses
        .vb_shi502_current_uses = tSI502.shi502_current_uses
        .vb_shi502_path = GetPtrToStrA(tSI502.shi502_path)
        .vb_shi502_passwd = GetPtrToStrA(tSI502.shi502_passwd)
        .vb_shi502_reserved = tSI502.shi502_reserved
        .vb_shi502_security_descriptor = tSI502.shi502_security_descriptor
    End With
    Call NetAPIBufferFree(lngBuffer)
End Function

Public Function EnumSharedFolders2(ByRef tSF() As SoftwareSharedFolder) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim tSI502() As SHARE_INFO_502
    Dim lngSF As Long
    Dim i As Long
    
    lngSF = 0
    
    Do
        EnumSharedFolders2 = NetShareEnum(StrConv(strNetShareServer, vbUnicode), _
                                    ByVal 502, _
                                    lngBuffer, _
                                    lngMaxLen, _
                                    lngEntries, _
                                    lngTotal, _
                                    lngResume)
                                    
        If (EnumSharedFolders2 <> ERROR_MORE_DATA And EnumSharedFolders2 <> 0) Then Exit Function
        
        If lngEntries > 0 Then
            
            ReDim tSI502(lngEntries - 1)
            
            MoveMemory tSI502(0).shi502_netname, ByVal lngBuffer, Len(tSI502(0)) * lngEntries
            
            For i = 0 To lngEntries - 1
                If (tSI502(i).shi502_type = STYPE_DISKTREE) Then
                    ReDim Preserve tSF(lngSF)
                    With tSF(lngSF)
                        .ShareName = GetPtrToStrA(tSI502(i).shi502_netname)
                        .FolderPath = GetPtrToStrA(tSI502(i).shi502_path)
                        .Description = GetPtrToStrA(tSI502(i).shi502_remark)
                        .MaximumAllowed = tSI502(i).shi502_max_uses
                        '.ShareType = STYPE_DISKTREE
                    End With
                    lngSF = lngSF + 1
                End If
            Next i
        End If
        If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
    Loop While EnumSharedFolders2 = ERROR_MORE_DATA
End Function

Public Function ShareEnum502(ByRef lngCount As Long, _
                             tSI502_VB() As SHARE_INFO_502_VB) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim tSI502() As SHARE_INFO_502
    Dim tSF() As SoftwareSharedFolder
    Dim lngSF As Long
    Dim i As Long
    
    lngCount = 0
    ReDim tSF(lngCount)
    lngSF = 0
    
    Do
        ShareEnum502 = NetShareEnum(StrConv(strNetShareServer, vbUnicode), _
                                    ByVal 502, _
                                    lngBuffer, _
                                    lngMaxLen, _
                                    lngEntries, _
                                    lngTotal, _
                                    lngResume)
                                    
        If (ShareEnum502 <> ERROR_MORE_DATA And ShareEnum502 <> 0) Then Exit Function
        
        If lngEntries > 0 Then
            
            ReDim tSI502(lngEntries - 1)
            
            MoveMemory tSI502(0).shi502_netname, ByVal lngBuffer, Len(tSI502(0)) * lngEntries
            
            For i = 0 To lngEntries - 1
                ReDim Preserve tSI502_VB(lngCount)
                With tSI502_VB(lngCount)
                    .vb_shi502_netname = GetPtrToStrA(tSI502(i).shi502_netname)
                    .vb_shi502_type = tSI502(i).shi502_type
                    .vb_shi502_remark = GetPtrToStrA(tSI502(i).shi502_remark)
                    .vb_shi502_permissions = tSI502(i).shi502_permissions
                    .vb_shi502_max_uses = tSI502(i).shi502_max_uses
                    .vb_shi502_current_uses = tSI502(i).shi502_current_uses
                    .vb_shi502_path = GetPtrToStrA(tSI502(i).shi502_path)
                    .vb_shi502_passwd = GetPtrToStrA(tSI502(i).shi502_passwd)
                    .vb_shi502_reserved = tSI502(i).shi502_reserved
                    .vb_shi502_security_descriptor = tSI502(i).shi502_security_descriptor
                End With
                
                If (tSI502_VB(lngCount).vb_shi502_type = STYPE_DISKTREE) Then
                    ReDim Preserve tSF(lngSF)
                    With tSF(lngSF)
                        .ShareName = tSI502_VB(lngCount).vb_shi502_netname
                        .FolderPath = tSI502_VB(lngCount).vb_shi502_path
                        .Description = tSI502_VB(lngCount).vb_shi502_remark
                        .MaximumAllowed = tSI502_VB(lngCount).vb_shi502_max_uses
                        '.ShareType = "Disk Share"
                    End With
                    lngSF = lngSF + 1
                End If
                lngCount = lngCount + 1
            Next i
        End If
        If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
    Loop While ShareEnum502 = ERROR_MORE_DATA
    
    SW_SHARED_FOLDERS = tSF
End Function

Public Function ShareAdd2(tSI2_VB As SHARE_INFO_2_VB) As Long
    Dim lngBuffer As Long
    Dim lngError As Long
    Dim tSI2 As SHARE_INFO_2
    Dim lngNetName As Long
    Dim lngRemark As Long
    Dim lngPath As Long
    Dim lngPasswd As Long

    With tSI2_VB
        lngNetName = NetGetStrToPtr(.vb_shi2_netname)
        lngRemark = NetGetStrToPtr(.vb_shi2_remark)
        lngPath = NetGetStrToPtr(.vb_shi2_path)
        'lngPasswd = NetGetStrToPtr(vbNullString)
    End With
    With tSI2
         .shi2_netname = lngNetName
         .shi2_type = tSI2_VB.vb_shi2_type
         .shi2_remark = lngRemark
         .shi2_permissions = tSI2_VB.vb_shi2_permissions
         .shi2_max_uses = tSI2_VB.vb_shi2_max_uses
         .shi2_current_uses = tSI2_VB.vb_shi2_current_uses
         .shi2_path = lngPath
         .shi2_passwd = 0
    End With
    ShareAdd2 = NetShareAdd(StrConv(strNetShareServer, vbUnicode), _
                            ByVal 2, _
                            tSI2.shi2_netname, _
                            lngError)
    Call NetAPIBufferFree(lngNetName)
    Call NetAPIBufferFree(lngRemark)
    Call NetAPIBufferFree(lngPath)
End Function

Public Function ShareDel(ByVal strShareName As String) As Long
    ShareDel = NetShareDel(StrConv(strNetShareServer, vbUnicode), _
                           StrConv(strShareName, vbUnicode), _
                           0)
End Function

Public Function ShareCheck(ByVal strDevice As String, _
                           ByRef lngType As enmSTYPE) As Long
    ShareCheck = NetShareCheck(StrConv(strNetShareServer, vbUnicode), _
                               StrConv(strDevice, vbUnicode), _
                               lngType)
End Function

Public Function ShareSetInfo1(ByVal strShareName As String, _
                              tSI1_VB As SHARE_INFO_1_VB) As Long
    Dim lngNetName As Long
    Dim lngRemark As Long
    Dim tSI1 As SHARE_INFO_1
    Dim lngError As Long

    With tSI1_VB
        lngNetName = NetGetStrToPtr(.vb_shi1_netname)
        lngRemark = NetGetStrToPtr(.vb_shi1_remark)
    End With
    With tSI1
         .shi1_netname = lngNetName
         .shi1_type = tSI1_VB.vb_shi1_type
         .shi1_remark = lngRemark
    End With
    ShareSetInfo1 = NetShareSetInfo(StrConv(strNetShareServer, vbUnicode), _
                                    StrConv(strShareName, vbUnicode), _
                                    1, _
                                    tSI1.shi1_netname, _
                                    lngError)
    Call NetAPIBufferFree(lngNetName)
    Call NetAPIBufferFree(lngRemark)
End Function

Public Function ShareSetInfo2(ByVal strShareName As String, _
                              tSI2_VB As SHARE_INFO_2_VB) As Long
    Dim lngBuffer As Long
    Dim lngError As Long
    Dim tSI2 As SHARE_INFO_2
    Dim lngNetName As Long
    Dim lngRemark As Long
    Dim lngPath As Long

    With tSI2_VB
        lngNetName = NetGetStrToPtr(.vb_shi2_netname)
        lngRemark = NetGetStrToPtr(.vb_shi2_remark)
        lngPath = NetGetStrToPtr(.vb_shi2_path)
    End With
    With tSI2
         .shi2_netname = lngNetName
         .shi2_type = tSI2_VB.vb_shi2_type
         .shi2_remark = lngRemark
         .shi2_permissions = tSI2_VB.vb_shi2_permissions
         .shi2_max_uses = tSI2_VB.vb_shi2_max_uses
         .shi2_current_uses = tSI2_VB.vb_shi2_current_uses
         .shi2_path = lngPath
         .shi2_passwd = 0
    End With
    ShareSetInfo2 = NetShareSetInfo(StrConv(strNetShareServer, vbUnicode), _
                                    StrConv(strShareName, vbUnicode), _
                                    2, _
                                    tSI2.shi2_netname, _
                                    lngError)
    Call NetAPIBufferFree(lngNetName)
    Call NetAPIBufferFree(lngRemark)
    Call NetAPIBufferFree(lngPath)
End Function

Public Function ShareSetInfo1004(ByVal strShareName As String, _
                                 ByVal strComment As String) As Long
    Dim lngError As Long
    Dim lngRemark As Long

    lngRemark = NetGetStrToPtr(strComment)
    ShareSetInfo1004 = NetShareSetInfo(StrConv(strNetShareServer, vbUnicode), _
                                       StrConv(strShareName, vbUnicode), _
                                       1004, _
                                       lngRemark, _
                                       lngError)
    Call NetAPIBufferFree(lngRemark)
End Function

Public Function ShareSetInfo1006(ByVal strShareName As String, _
                                 ByVal lngMaxUses As Long) As Long
    Dim lngError As Long

    ShareSetInfo1006 = NetShareSetInfo(StrConv(strNetShareServer, vbUnicode), _
                                       StrConv(strShareName, vbUnicode), _
                                       1006, _
                                       lngMaxUses, _
                                       lngError)
End Function


Private Function NetGetStrToPtr(ByVal str As String) As Long
    Dim byteBuffer() As Byte
    byteBuffer = str & vbNullChar
    Call NetAPIBufferAllocate(UBound(byteBuffer) + 1, NetGetStrToPtr)
    Call StrToPtr(NetGetStrToPtr, byteBuffer(0))
End Function

