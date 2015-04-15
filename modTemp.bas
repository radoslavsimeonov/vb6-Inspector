Attribute VB_Name = "modTemp"
Option Explicit

Private Declare Function DnsQuery _
                Lib "dnsapi" _
                Alias "DnsQuery_A" (ByVal strName As String, _
                                    ByVal wType As Integer, _
                                    ByVal fOptions As Long, _
                                    ByVal pServers As Long, _
                                    ppQueryResultsSet As Long, _
                                    ByVal pReserved As Long) As Long

Private Declare Function DnsRecordListFree _
                Lib "dnsapi" (ByVal pDnsRecord As Long, _
                              ByVal FreeType As Long) As Long

Private Declare Function lstrlen Lib "KERNEL32" (ByVal straddress As Long) As Long
Private Declare Sub CopyMemory _
                Lib "KERNEL32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       ByVal Source As Long, _
                                       ByVal Length As Long)

Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal pIP As Long) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal sAddr As String) As Long
Private Declare Function StrCopyA _
                Lib "KERNEL32" _
                Alias "lstrcpyA" (ByVal RetVal As String, _
                                  ByVal Ptr As Long) As Long

Private Declare Function StrLenA _
                Lib "KERNEL32" _
                Alias "lstrlenA" (ByVal Ptr As Long) As Long

Private Const DNS_TYPE_PTR           As Long = &HC
Private Const DNS_QUERY_STANDARD     As Long = &H0
Private Const DnsFreeRecordListDeep  As Long = 1&
Private Const DnsFreeRecordList      As Long = 1
Private Const DNS_TYPE_A             As Long = &H1
Private Const DNS_QUERY_BYPASS_CACHE As Long = &H8

Private Enum DNS_STATUS
    ERROR_BAD_IP_FORMAT = -3&
    ERROR_NO_PTR_RETURNED = -2&
    ERROR_NO_RR_RETURNED = -1&
    DNS_STATUS_SUCCESS = 0&
End Enum

Private Type VBDnsRecord
    pNext           As Long
    pName           As Long
    wType           As Integer
    wDataLength     As Integer
    Flags           As Long
    dwTel           As Long
    dwReserved      As Long
    prt             As Long
    others(35)      As Byte
End Type

Public Function IP2HostName(ByVal IP As String, ByRef hostname As String) As Long

    Dim Octets()  As String
    Dim OctX      As Long
    Dim NumPart   As Long
    Dim BadIP     As Boolean
    Dim lngDNSRec As Long
    Dim Record    As VBDnsRecord
    Dim Length    As Long

    IP = Trim$(IP)

    If Len(IP) = 0 Then
        IP2HostName = ERROR_BAD_IP_FORMAT
        Exit Function
    End If

    Octets = Split(IP, ".")

    If UBound(Octets) <> 3 Then
        IP2HostName = ERROR_BAD_IP_FORMAT
        Exit Function
    End If

    For OctX = 0 To 3

        If IsNumeric(Octets(OctX)) Then
            NumPart = CInt(Octets(OctX))

            If 0 <= NumPart And NumPart <= 255 Then
                Octets(OctX) = CStr(NumPart)
            Else
                BadIP = True
                Exit For
            End If

        Else
            BadIP = True
            Exit For
        End If

    Next

    If BadIP Then
        IP2HostName = ERROR_BAD_IP_FORMAT
        Exit Function
    End If

    IP = Octets(3) & "." & Octets(2) & "." & Octets(1) & "." & Octets(0) & ".IN-ADDR.ARPA"
    IP2HostName = DnsQuery(IP, DNS_TYPE_PTR, DNS_QUERY_STANDARD, ByVal 0, lngDNSRec, 0)

    If IP2HostName = DNS_STATUS_SUCCESS Then
        If lngDNSRec <> 0 Then
            CopyMemory Record, ByVal lngDNSRec, LenB(Record)

            With Record

                If .wType = DNS_TYPE_PTR Then
                    Length = StrLenA(.prt)
                    hostname = String$(Length, 0)
                    StrCopyA hostname, .prt
                Else
                    IP2HostName = ERROR_NO_PTR_RETURNED
                End If

            End With

            DnsRecordListFree lngDNSRec, DnsFreeRecordListDeep
        Else
            IP2HostName = ERROR_NO_RR_RETURNED
        End If

    End If

End Function

Public Function HostNameToIP(sAddr As String, Optional sDnsServers As String = vbNullString) As String

    Dim pRecord     As Long
    Dim pNext       As Long
    Dim uRecord     As VBDnsRecord
    Dim lPtr        As Long
    Dim vSplit      As Variant
    Dim laServers() As Long
    Dim pServers    As Long
    Dim sName       As String

    If LenB(sDnsServers) <> 0 Then
        vSplit = Split(sDnsServers)
        ReDim laServers(0 To UBound(vSplit) + 1)
        laServers(0) = UBound(laServers)

        For lPtr = 0 To UBound(vSplit)
            laServers(lPtr + 1) = inet_addr(vSplit(lPtr))
        Next

        pServers = VarPtr(laServers(0))
    End If

    If DnsQuery(sAddr, DNS_TYPE_A, DNS_QUERY_BYPASS_CACHE, pServers, pRecord, 0) = 0 Then
        pNext = pRecord

        Do While pNext <> 0
            Call CopyMemory(uRecord, pNext, Len(uRecord))

            If uRecord.wType = DNS_TYPE_A Then
                lPtr = inet_ntoa(uRecord.prt)
                sName = String$(lstrlen(lPtr), 0)
                Call CopyMemory(ByVal sName, lPtr, Len(sName))

                If LenB(HostNameToIP) <> 0 Then
                    HostNameToIP = HostNameToIP & " "
                End If

                HostNameToIP = HostNameToIP & sName
            End If

            pNext = uRecord.pNext
        Loop

        Call DnsRecordListFree(pRecord, DnsFreeRecordList)
    End If

End Function
