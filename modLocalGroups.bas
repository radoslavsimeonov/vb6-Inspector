Attribute VB_Name = "swLocalGroups"
Option Explicit

Private Const ERROR_MORE_DATA = 234

Enum enmSE_GROUP
    SE_GROUP_MANDATORY = &H1&
    SE_GROUP_ENABLED_BY_DEFAULT = &H2&
    SE_GROUP_ENABLED = &H4&
    SE_GROUP_OWNER = &H8&
    SE_GROUP_LOGON_ID = &HC0000000
End Enum

Enum enmSID_NAME_USE
    SidTypeUser
    SidTypeGroup
    SidTypeDomain
    SidTypeAlias
    SidTypeWellKnownGroup
    SidTypeDeletedAccount
    SidTypeInvalid
    SidTypeUnknown
End Enum

Private Type LOCALGROUP_INFO_0
    lgrpi0_name As Long
End Type

Private Type LOCALGROUP_INFO_1
    lgrpi1_name As Long
    lgrpi1_comment As Long
End Type

Private Type LOCALGROUP_INFO_1002
    lgrpi1002_comment As Long
End Type

Private Type LOCALGROUP_MEMBERS_INFO_0
    lgrmi0_sid As Long
End Type

Private Type LOCALGROUP_MEMBERS_INFO_1
    lgrmi1_sid As Long
    lgrmi1_sidusage As enmSID_NAME_USE
    lgrmi1_name As Long
End Type

Private Type LOCALGROUP_MEMBERS_INFO_2
    lgrmi2_sid As Long
    lgrmi2_sidusage As enmSID_NAME_USE
    lgrmi2_domainandname As Long
End Type

Private Type LOCALGROUP_MEMBERS_INFO_3
    lgrmi3_domainandname As Long
End Type

Type LOCALGROUP_INFO_0_VB
    vb_lgrpi0_name As String
End Type

Type LOCALGROUP_INFO_1_VB
    vb_lgrpi1_name As String
    vb_lgrpi1_comment As String
End Type

Type LOCALGROUP_INFO_1002_VB
    vb_lgrpi1002_comment As Long
End Type

Type LOCALGROUP_MEMBERS_INFO_0_VB
    vb_lgrmi0_sid As Long
End Type

Type LOCALGROUP_MEMBERS_INFO_1_VB
    vb_lgrmi1_sid As Long
    vb_lgrmi1_sidusage As enmSID_NAME_USE
    vb_lgrmi1_name As String
End Type

Type LOCALGROUP_MEMBERS_INFO_2_VB
    vb_lgrmi2_sid As Long
    vb_lgrmi2_sidusage As enmSID_NAME_USE
    vb_lgrmi2_domainandname As String
End Type

Type LOCALGROUP_MEMBERS_INFO_3_VB
    vb_lgrmi3_domainandname As String
End Type

Private Type LOCALGROUP_USERS_INFO_0
    lgrui0_name As Long
End Type

Type GROUP_USERS_INFO_1_VB
    vb_grui1_name As String
    vb_grui1_attributes As enmSE_GROUP
End Type

Private Type GROUP_USERS_INFO_0
    grui0_name As Long
End Type

Private Type GROUP_USERS_INFO_1
    grui1_name As Long
    grui1_attributes As Long
End Type

Private Declare Function NetLocalGroupEnum Lib "netapi32.dll" (ByVal servername As String, _
                                                               ByVal level As Long, _
                                                               bufptr As Any, _
                                                               prefmaxlen As Long, _
                                                               entriesread As Long, _
                                                               totalentries As Long, _
                                                               resumehandle As Any) As Long

Private Declare Function NetLocalGroupGetInfo Lib "netapi32.dll" (ByVal servername As String, _
                                                                  ByVal LocalGroupName As String, _
                                                                  ByVal level As Long, _
                                                                  bufptr As Any) As Long

Private Declare Function NetLocalGroupSetInfo Lib "netapi32.dll" (ByVal servername As String, _
                                                                  ByVal LocalGroupName As String, _
                                                                  ByVal level As Long, _
                                                                  buf As Any, _
                                                                  parm_err As Long) As Long

Private Declare Function NetLocalGroupAdd Lib "netapi32.dll" (ByVal servername As String, _
                                                              ByVal level As Long, _
                                                              buf As Any, _
                                                              parm_err As Long) As Long

Private Declare Function NetLocalGroupDel Lib "netapi32.dll" (ByVal servername As String, _
                                                              ByVal LocalGroupName As String) As Long

Private Declare Function NetLocalGroupGetMembers Lib "netapi32.dll" (ByVal servername As String, _
                                                                     ByVal LocalGroupName As String, _
                                                                     ByVal level As Long, _
                                                                     bufptr As Any, _
                                                                     prefmaxlen As Long, _
                                                                     entriesread As Long, _
                                                                     totalentries As Long, _
                                                                     resumehandle As Any) As Long

Private Declare Function NetLocalGroupAddMembers Lib "netapi32.dll" (ByVal servername As String, _
                                                                     ByVal LocalGroupName As String, _
                                                                     ByVal level As Long, _
                                                                     buf As Any, _
                                                                     ByVal membercount As Long) As Long

Private Declare Function NetLocalGroupDelMembers Lib "netapi32.dll" (ByVal servername As String, _
                                                                     ByVal LocalGroupName As String, _
                                                                     ByVal level As Long, _
                                                                     buf As Any, _
                                                                     ByVal membercount As Long) As Long


Private Declare Function NetLocalGroupSetMembers Lib "netapi32.dll" (ByVal servername As String, _
                                                                     ByVal LocalGroupName As String, _
                                                                     ByVal level As Long, _
                                                                     buf As Any, _
                                                                     ByVal totalentries As Long) As Long

Private Declare Function NetUserGetLocalGroups Lib "netapi32.dll" (ByVal servername As String, _
                                                                   ByVal username As String, _
                                                                   ByVal level As Long, _
                                                                   ByVal Flags As Long, _
                                                                   bufptr As Any, _
                                                                   prefmaxlen As Long, _
                                                                   entriesread As Long, _
                                                                   totalentries As Long) As Long
                                                                   
Private Declare Function NetUserSetGroups Lib "netapi32.dll" (ByVal servername As String, _
                                                              ByVal username As String, _
                                                              ByVal level As Long, _
                                                              bufptr As Any, _
                                                              ByVal num_entries As Long) As Long
                                                              
Private Declare Function NetUserGetGroups Lib "netapi32.dll" (ByVal servername As String, _
                                                              ByVal username As String, _
                                                              ByVal level As Long, _
                                                              bufptr As Any, _
                                                              prefmaxlen As Long, _
                                                              entriesread As Long, _
                                                              totalentries As Long) As Long

Private Declare Function NetAPIBufferAllocate Lib "netapi32.dll" _
                                              Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, _
                                                                            Ptr As Long) As Long
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" _
                                          Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long

Private Declare Function GetComputerName Lib "KERNEL32" _
                                         Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                                                   nSize As Long) As Long

Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, _
                                                            pSource As Any, _
                                                            ByVal dwLength As Long)
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, _
                                                            pSrc As Any, _
                                                            ByVal ByteLen As Long)
Declare Function StrLen Lib "KERNEL32" Alias "lstrlenW" (ByVal Ptr As Long) As Long

Declare Function PtrToStr Lib "KERNEL32" Alias "lstrcpyW" (RetVal As Byte, _
                                                           ByVal Ptr As Long) As Long
Declare Function StrToPtr Lib "KERNEL32" Alias "lstrcpyW" (ByVal Ptr As Long, _
                                                           Source As Byte) As Long

Private strNetLocalGroupServer As String


Public Property Get NetLocalGroupServer() As String
    NetLocalGroupServer = strNetLocalGroupServer
End Property

Public Property Let NetLocalGroupServer(ByVal vNewValue As String)
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
    strNetLocalGroupServer = UCase$(vNewValue)
End Property


Public Function LocalGroupEnum0(ByRef lngCount As Long, _
                                tLocalGroupInfo0VB() As LOCALGROUP_INFO_0_VB) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim strDomain As String
    Dim tLocalGroupInfo0() As LOCALGROUP_INFO_0
    Dim i As Long
    
    lngCount = 0
    Do
        LocalGroupEnum0 = NetLocalGroupEnum(StrConv(strNetLocalGroupServer, vbUnicode), _
                                            ByVal 0, _
                                            lngBuffer, _
                                            lngMaxLen, _
                                            lngEntries, _
                                            lngTotal, _
                                            lngResume)
        If (LocalGroupEnum0 <> ERROR_MORE_DATA And LocalGroupEnum0 <> 0) Then Exit Function
        If lngEntries > 0 Then
            ReDim tLocalGroupInfo0(lngEntries - 1)
            MoveMemory tLocalGroupInfo0(0).lgrpi0_name, ByVal lngBuffer, Len(tLocalGroupInfo0(0)) * lngEntries
            For i = 0 To lngEntries - 1
                ReDim Preserve tLocalGroupInfo0VB(lngCount)
                With tLocalGroupInfo0VB(lngCount)
                    .vb_lgrpi0_name = GetPtrToStrA(tLocalGroupInfo0(i).lgrpi0_name)
                End With
                lngCount = lngCount + 1
            Next i
        End If
        If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
    Loop While LocalGroupEnum0 = ERROR_MORE_DATA
End Function

Public Function LocalGroupEnum1(ByRef lngCount As Long, _
                                tLocalGroupInfo1VB() As LOCALGROUP_INFO_1_VB) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim strDomain As String
    Dim tLocalGroupInfo1() As LOCALGROUP_INFO_1
    Dim i As Long
    
    lngCount = 0
    Do
        LocalGroupEnum1 = NetLocalGroupEnum(StrConv(strNetLocalGroupServer, vbUnicode), _
                                            ByVal 1, _
                                            lngBuffer, _
                                            lngMaxLen, _
                                            lngEntries, _
                                            lngTotal, _
                                            lngResume)
        If (LocalGroupEnum1 <> ERROR_MORE_DATA And LocalGroupEnum1 <> 0) Then Exit Function
        If lngEntries > 0 Then
            ReDim tLocalGroupInfo1(lngEntries - 1)
            MoveMemory tLocalGroupInfo1(0).lgrpi1_name, ByVal lngBuffer, Len(tLocalGroupInfo1(0)) * lngEntries
            For i = 0 To lngEntries - 1
                ReDim Preserve tLocalGroupInfo1VB(lngCount)
                With tLocalGroupInfo1VB(lngCount)
                    .vb_lgrpi1_name = GetPtrToStrA(tLocalGroupInfo1(i).lgrpi1_name)
                    .vb_lgrpi1_comment = GetPtrToStrA(tLocalGroupInfo1(i).lgrpi1_comment)
                End With
                lngCount = lngCount + 1
            Next i
        End If
        If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
    Loop While LocalGroupEnum1 = ERROR_MORE_DATA
End Function

Public Function LocalGroupGetMembers0(ByVal strLocalGroupName As String, _
                                      ByRef lngCount As Long, _
                                      tLocalGroupMInfo0VB() As LOCALGROUP_MEMBERS_INFO_0_VB) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim strDomain As String
    Dim tLocalGroupMInfo0() As LOCALGROUP_MEMBERS_INFO_0
    Dim i As Long
    
    lngCount = 0
    Do
        LocalGroupGetMembers0 = NetLocalGroupGetMembers(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                        StrConv(strLocalGroupName, vbUnicode), _
                                                        ByVal 0, _
                                                        lngBuffer, _
                                                        lngMaxLen, _
                                                        lngEntries, _
                                                        lngTotal, _
                                                        lngResume)
        If (LocalGroupGetMembers0 <> ERROR_MORE_DATA And LocalGroupGetMembers0 <> 0) Then Exit Function
        If lngEntries > 0 Then
            ReDim tLocalGroupMInfo0(lngEntries - 1)
            MoveMemory tLocalGroupMInfo0(0).lgrmi0_sid, ByVal lngBuffer, Len(tLocalGroupMInfo0(0)) * lngEntries
            For i = 0 To lngEntries - 1
                ReDim Preserve tLocalGroupMInfo0VB(lngCount)
                With tLocalGroupMInfo0VB(lngCount)
                    .vb_lgrmi0_sid = tLocalGroupMInfo0(i).lgrmi0_sid
                End With
                lngCount = lngCount + 1
            Next i
        End If
        If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
    Loop While LocalGroupGetMembers0 = ERROR_MORE_DATA
End Function

Public Function LocalGroupGetMembers1(ByVal strLocalGroupName As String, _
                                      ByRef lngCount As Long, _
                                      tLocalGroupMInfo1VB() As LOCALGROUP_MEMBERS_INFO_1_VB) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim strDomain As String
    Dim tLocalGroupMInfo1() As LOCALGROUP_MEMBERS_INFO_1
    Dim i As Long
    
    lngCount = 0
    Do
        LocalGroupGetMembers1 = NetLocalGroupGetMembers(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                        StrConv(strLocalGroupName, vbUnicode), _
                                                        ByVal 1, _
                                                        lngBuffer, _
                                                        lngMaxLen, _
                                                        lngEntries, _
                                                        lngTotal, _
                                                        lngResume)
        If (LocalGroupGetMembers1 <> ERROR_MORE_DATA And LocalGroupGetMembers1 <> 0) Then Exit Function
        If lngEntries > 0 Then
            ReDim tLocalGroupMInfo1(lngEntries - 1)
            MoveMemory tLocalGroupMInfo1(0).lgrmi1_sid, ByVal lngBuffer, Len(tLocalGroupMInfo1(0)) * lngEntries
            For i = 0 To lngEntries - 1
                ReDim Preserve tLocalGroupMInfo1VB(lngCount)
                With tLocalGroupMInfo1VB(lngCount)
                    .vb_lgrmi1_sid = tLocalGroupMInfo1(i).lgrmi1_sid
                    .vb_lgrmi1_sidusage = tLocalGroupMInfo1(i).lgrmi1_sidusage
                    .vb_lgrmi1_name = GetPtrToStrA(tLocalGroupMInfo1(i).lgrmi1_name)
                End With
                lngCount = lngCount + 1
            Next i
        End If
        If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
    Loop While LocalGroupGetMembers1 = ERROR_MORE_DATA
End Function

Public Function LocalGroupGetMembers2(ByVal strLocalGroupName As String, _
                                      ByRef lngCount As Long, _
                                      tLocalGroupMInfo2VB() As LOCALGROUP_MEMBERS_INFO_2_VB) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim strDomain As String
    Dim tLocalGroupMInfo2() As LOCALGROUP_MEMBERS_INFO_2
    Dim i As Long
    
    lngCount = 0
    Do
        LocalGroupGetMembers2 = NetLocalGroupGetMembers(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                        StrConv(strLocalGroupName, vbUnicode), _
                                                        ByVal 2, _
                                                        lngBuffer, _
                                                        lngMaxLen, _
                                                        lngEntries, _
                                                        lngTotal, _
                                                        lngResume)
        If (LocalGroupGetMembers2 <> ERROR_MORE_DATA And LocalGroupGetMembers2 <> 0) Then Exit Function
        If lngEntries > 0 Then
            ReDim tLocalGroupMInfo2(lngEntries - 1)
            MoveMemory tLocalGroupMInfo2(0).lgrmi2_sid, ByVal lngBuffer, Len(tLocalGroupMInfo2(0)) * lngEntries
            For i = 0 To lngEntries - 1
                ReDim Preserve tLocalGroupMInfo2VB(lngCount)
                With tLocalGroupMInfo2VB(lngCount)
                    .vb_lgrmi2_sid = tLocalGroupMInfo2(i).lgrmi2_sid
                    .vb_lgrmi2_sidusage = tLocalGroupMInfo2(i).lgrmi2_sidusage
                    .vb_lgrmi2_domainandname = GetPtrToStrA(tLocalGroupMInfo2(i).lgrmi2_domainandname)
                End With
                lngCount = lngCount + 1
            Next i
        End If
        If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
    Loop While LocalGroupGetMembers2 = ERROR_MORE_DATA
End Function

Public Function LocalGroupGetInfo1(ByVal strLocalGroupName As String, _
                                   tLocalGroupInfo1VB As LOCALGROUP_INFO_1_VB) As Long
    Dim lngBuffer As Long
    Dim tLocalGroupInfo1 As LOCALGROUP_INFO_1
    
    LocalGroupGetInfo1 = NetLocalGroupGetInfo(StrConv(strNetLocalGroupServer, vbUnicode), _
                                              StrConv(strLocalGroupName, vbUnicode), _
                                              ByVal 1, _
                                              lngBuffer)
    If LocalGroupGetInfo1 <> 0 Then Exit Function
    MoveMemory tLocalGroupInfo1.lgrpi1_name, ByVal lngBuffer, Len(tLocalGroupInfo1)
    With tLocalGroupInfo1VB
        .vb_lgrpi1_name = GetPtrToStrA(tLocalGroupInfo1.lgrpi1_name)
        .vb_lgrpi1_comment = GetPtrToStrA(tLocalGroupInfo1.lgrpi1_comment)
    End With
    If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
End Function

Public Function LocalGroupAdd0(ByVal strLocalGroupName As String) As Long
    Dim lngErr As Long
    Dim lngBuffer As Long

    lngBuffer = NetGetStrToPtr(strLocalGroupName)

    LocalGroupAdd0 = NetLocalGroupAdd(StrConv(strNetLocalGroupServer, vbUnicode), _
                                      0, _
                                      lngBuffer, _
                                      lngErr)

    Call NetAPIBufferFree(lngBuffer)
End Function

Public Function LocalGroupAdd1(ByVal strLocalGroupName As String, _
                               ByVal strComment As String) As Long
                               
    Dim lngErr As Long
    Dim lngLocalGroupName As Long
    Dim lngComment As Long
    Dim tLocalGroupInfo1 As LOCALGROUP_INFO_1
    

    lngLocalGroupName = NetGetStrToPtr(strLocalGroupName)
    lngComment = NetGetStrToPtr(strComment)

    With tLocalGroupInfo1
        .lgrpi1_name = lngLocalGroupName
        .lgrpi1_comment = lngComment
    End With

    LocalGroupAdd1 = NetLocalGroupAdd(StrConv(strNetLocalGroupServer, vbUnicode), _
                                      1, _
                                      tLocalGroupInfo1.lgrpi1_name, _
                                      lngErr)

    Call NetAPIBufferFree(lngLocalGroupName)
    Call NetAPIBufferFree(lngComment)
End Function

Public Function LocalGroupDel(ByVal strLocalGroupName As String) As Long
                               

    LocalGroupDel = NetLocalGroupDel(StrConv(strNetLocalGroupServer, vbUnicode), _
                                     StrConv(strLocalGroupName, vbUnicode))
End Function

Public Function LocalGroupAddMembers0(ByVal strLocalGroupName As String, _
                                      ByVal lngSID As Long) As Long

    LocalGroupAddMembers0 = NetLocalGroupAddMembers(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                    StrConv(strLocalGroupName, vbUnicode), _
                                                    0, _
                                                    lngSID, _
                                                    1)
End Function

Public Function LocalGroupAddMembers3(ByVal strLocalGroupName As String, _
                                      ByVal strDomainAndName As String) As Long
    Dim lngBuffer As Long

    lngBuffer = NetGetStrToPtr(strDomainAndName)

    LocalGroupAddMembers3 = NetLocalGroupAddMembers(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                    StrConv(strLocalGroupName, vbUnicode), _
                                                    3, _
                                                    lngBuffer, _
                                                    1)

    Call NetAPIBufferFree(lngBuffer)
End Function

Public Function LocalGroupDelMembers0(ByVal strLocalGroupName As String, _
                                      ByVal lngSID As Long) As Long

    LocalGroupDelMembers0 = NetLocalGroupDelMembers(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                    StrConv(strLocalGroupName, vbUnicode), _
                                                    0, _
                                                    lngSID, _
                                                    1)
End Function

Public Function LocalGroupDelMembers3(ByVal strLocalGroupName As String, _
                                      ByVal strDomainAndName As String) As Long
    Dim lngBuffer As Long

    lngBuffer = NetGetStrToPtr(strDomainAndName)

    LocalGroupDelMembers3 = NetLocalGroupDelMembers(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                    StrConv(strLocalGroupName, vbUnicode), _
                                                    3, _
                                                    lngBuffer, _
                                                    1)

    Call NetAPIBufferFree(lngBuffer)
End Function

Public Function LocalGroupSetInfo1002(ByVal strLocalGroupName As String, _
                                      ByVal strComment As String) As Long
    Dim lngBuffer As Long
    Dim lngError As Long


    lngBuffer = NetGetStrToPtr(strComment)

    LocalGroupSetInfo1002 = NetLocalGroupSetInfo(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                 StrConv(strLocalGroupName, vbUnicode), _
                                                 ByVal 1002, _
                                                 lngBuffer, _
                                                 lngError)

    If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
End Function

Public Function LocalGroupSetMembers3(ByVal strLocalGroupName As String, _
                                      tDomainAndName() As String) As Long
    Dim lngTotalEntries As Long
    Dim lngBuffer() As Long
    Dim i As Long
    

    lngTotalEntries = UBound(tDomainAndName)
    ReDim lngBuffer(lngTotalEntries)
    For i = 0 To lngTotalEntries
        lngBuffer(i) = NetGetStrToPtr(tDomainAndName(i))
    Next i

    LocalGroupSetMembers3 = NetLocalGroupSetMembers(StrConv(strNetLocalGroupServer, vbUnicode), _
                                                    StrConv(strLocalGroupName, vbUnicode), _
                                                    ByVal 3, _
                                                    lngBuffer(0), _
                                                    lngTotalEntries + 1)
    For i = 0 To lngTotalEntries
        NetAPIBufferFree (lngBuffer(i))
    Next i
End Function

Public Function UserGetLocalGroups(ByVal strUserName As String, _
                                   ByVal lngFlags As Long, _
                                   ByRef lngCount As Long, _
                                   tLocalGroupUsers() As String) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim tLGUsersInfo0() As LOCALGROUP_USERS_INFO_0
    Dim i As Long
    
    lngCount = 0
    
    UserGetLocalGroups = NetUserGetLocalGroups(StrConv(strNetLocalGroupServer, vbUnicode), _
                                               StrConv(strUserName, vbUnicode), _
                                               0, _
                                               lngFlags, _
                                               lngBuffer, _
                                               lngMaxLen, _
                                               lngEntries, _
                                               lngTotal)
    If UserGetLocalGroups <> 0 Then Exit Function
    If lngEntries > 0 Then
        ReDim tLGUsersInfo0(lngEntries - 1)

        MoveMemory tLGUsersInfo0(0).lgrui0_name, ByVal lngBuffer, Len(tLGUsersInfo0(0)) * lngEntries

        For i = 0 To lngEntries - 1
            ReDim Preserve tLocalGroupUsers(lngCount)
            tLocalGroupUsers(lngCount) = GetPtrToStrA(tLGUsersInfo0(i).lgrui0_name)
            lngCount = lngCount + 1
        Next i
    End If

    If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
End Function

Public Function UserGetGroups1(ByVal strUserName As String, _
                               ByRef lngCount As Long, _
                               tGUI1VB() As GROUP_USERS_INFO_1_VB) As Long
    Dim lngBuffer As Long
    Dim lngMaxLen As Long
    Dim lngEntries As Long
    Dim lngTotal As Long
    Dim lngResume As Long
    Dim tGUI1() As GROUP_USERS_INFO_1
    Dim i As Long
    
    lngCount = 0

    UserGetGroups1 = NetUserGetGroups(StrConv(strNetLocalGroupServer, vbUnicode), _
                                      StrConv(strUserName, vbUnicode), _
                                      1, _
                                      lngBuffer, _
                                      lngMaxLen, _
                                      lngEntries, _
                                      lngTotal)
    If UserGetGroups1 <> 0 Then Exit Function
    If lngEntries > 0 Then

        ReDim tGUI1(lngEntries - 1)

        MoveMemory tGUI1(0).grui1_name, ByVal lngBuffer, Len(tGUI1(0)) * lngEntries

        For i = 0 To lngEntries - 1
            ReDim Preserve tGUI1VB(lngCount)
            With tGUI1VB(lngCount)
                .vb_grui1_name = GetPtrToStrA(tGUI1(i).grui1_name)
                .vb_grui1_attributes = tGUI1(i).grui1_attributes
            End With
            lngCount = lngCount + 1
        Next i
    End If

    If lngBuffer <> 0 Then NetAPIBufferFree (lngBuffer)
End Function

Public Function UserSetGroups0(ByVal strUserName As String, _
                               tGroups() As String) As Long
    Dim lngTotalEntries As Long
    Dim lngBuffer() As Long
    Dim i As Long

    lngTotalEntries = UBound(tGroups)
    ReDim lngBuffer(lngTotalEntries)
    For i = 0 To lngTotalEntries
        lngBuffer(i) = NetGetStrToPtr(tGroups(i))
    Next i

    UserSetGroups0 = NetUserSetGroups(StrConv("WIN2K12R2", vbUnicode), _
                                      StrConv(strUserName, vbUnicode), _
                                      0, _
                                      lngBuffer(0), _
                                      ByVal lngTotalEntries + 1)

    For i = 0 To lngTotalEntries
        NetAPIBufferFree (lngBuffer(i))
    Next i
End Function

Private Function NetGetStrToPtr(ByVal str As String) As Long
    Dim byteBuffer() As Byte
    byteBuffer = str & vbNullChar
    Call NetAPIBufferAllocate(UBound(byteBuffer) + 1, NetGetStrToPtr)
    Call StrToPtr(NetGetStrToPtr, byteBuffer(0))
End Function

Private Function GetPtrToStrA(lngPoint As Long) As String
    Dim byteBuffer(1024) As Byte
    Call PtrToStr(byteBuffer(0), lngPoint)
    GetPtrToStrA = Left$(byteBuffer, StrLen(lngPoint))
End Function
