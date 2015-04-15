Attribute VB_Name = "swSecurityProducts"
Option Explicit

Private Const SECURITY_ANTIVIRUS As String = "AntivirusProduct"
Private Const SECURITY_FIREWALL As String = "FirewallProduct"
Private Const SECURITY_SPYWARE As String = "AntiSpywareProduct"

Public Function GetSecuityProducts() As SoftwareSecurityProducts
    
    Dim secP As SoftwareSecurityProducts
    
    With secP
        .AntiVirus() = GetSecurityProduct(SECURITY_ANTIVIRUS)
        .Firewall() = GetSecurityProduct(SECURITY_FIREWALL)
        .Spyware() = GetSecurityProduct(SECURITY_SPYWARE)
    End With
    
    GetSecuityProducts = secP
    
End Function

Private Function GetSecurityProduct(sProduct As String) As SoftwareSecurityProduct()
    On Error Resume Next

    Dim oWMI, colItems, objItem
    Dim tAV() As SoftwareSecurityProduct
    Dim cnt   As Integer

    cnt = 0
    ReDim tAV(cnt)

    Select Case OS_VERSION
    
        Case Is < "5.2"
            
            Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\SecurityCenter")
              
            Set colItems = oWMI.ExecQuery("Select * from " & sProduct)
            
            If colItems.count > 0 Then
            
                For Each objItem In colItems
                    ReDim Preserve tAV(cnt)
                    
                    If GetProp(objItem.DisplayName) <> vbNullString Then

                        With tAV(cnt)
                            .Publisher = GetProp(objItem.CompanyName)
                            .ProductName = GetProp(objItem.DisplayName)
                            .ProductVersion = GetProp(objItem.versionNumber)
                            .ProductGUID = GetProp(objItem.instanceGuid)

                            If sProduct = SECURITY_ANTIVIRUS Then
                                .RTPStatus = GetProp(objItem.onAccessScanningEnabled)
                                .UpToDate = GetProp(objItem.productUptoDate)
                            End If

                            If sProduct = SECURITY_FIREWALL Then .Enabled = GetProp(objItem.Enabled)

                        End With
                        
                    cnt = cnt + 1
                End If

            Next

        End If
        
    Case Is > "5.1"
            
        Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\SecurityCenter2")
              
        Set colItems = oWMI.ExecQuery("Select * from " & sProduct)
            
        If colItems.count > 0 Then

            For Each objItem In colItems
                ReDim Preserve tAV(cnt)
                
                If GetProp(objItem.DisplayName) <> vbNullString Then
                    With tAV(cnt)
                        .ProductName = GetProp(objItem.DisplayName)
                        .PathToProduct = GetProp(objItem.pathToSignedProductExe)
                        .ProductState = GetProp(objItem.ProductState)
                        GetAVState .ProductState, .UpToDate, .RTPStatus
                        .Enabled = -2
                    End With
                                  
                    cnt = cnt + 1
                End If
            Next

        End If

End Select

GetSecurityProduct = tAV
    
End Function

Private Sub GetAVState(sProductState As String, ByRef bUPD As Integer, ByRef bRTP As Integer)
    
    Const bTrue As Integer = -1
    Const bFalse As Integer = 0

    Select Case sProductState

        Case "262144": bUPD = bTrue: bRTP = bFalse
        Case "262160": bUPD = bFalse: bRTP = bFalse
        Case "266240": bUPD = bTrue: bRTP = bTrue
        Case "266256": bUPD = bFalse: bRTP = bTrue
        Case "393216": bUPD = bTrue: bRTP = bFalse
        Case "393232": bUPD = bFalse: bRTP = bFalse
        Case "393488": bUPD = bFalse: bRTP = bFalse
        Case "397312": bUPD = bTrue: bRTP = bTrue
        Case "397328": bUPD = bFalse: bRTP = bTrue
        Case "397584": bUPD = bFalse: bRTP = bTrue
        Case Else
            bUPD = -2: bRTP = -2
            Dim avStatus As String
            
            avStatus = Trim$(str$(Hex(sProductState)))
            
            If Mid$(avStatus, 2, 2) = "10" Or Mid$(avStatus, 2, 2) = "11" Then
                bRTP = bTrue
            ElseIf Mid$(avStatus, 2, 2) = "00" Or Mid$(avStatus, 2, 2) = "01" Then
                bRTP = bFalse
            End If
            
            If Mid(avStatus, 4, 2) = "00" Then
                bUPD = bTrue
            ElseIf Mid(avStatus, 4, 2) = "10" Then
                bUPD = bFalse
            End If

     End Select
End Sub

