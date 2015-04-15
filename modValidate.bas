Attribute VB_Name = "modValidate"
Option Explicit

Public Function ValidateWorkstationDetails() As Boolean
    
    ValidateWorkstationDetails = False
    
    With REG_WORKSTATION
        If .Classification = vbNullString Then
            MsgBox "Please select security clearance", vbExclamation
            frmUserDetails.cmbClassification.SetFocus
        ElseIf .BookNo = vbNullString Then
            MsgBox "Please enter book number", vbExclamation
            frmUserDetails.txtBookNo.SetFocus
        ElseIf .BookDate = vbNullString Then
            MsgBox "Please enter book date", vbExclamation
            frmUserDetails.txtBookDate.SetFocus
        ElseIf Join(.CaseStickers) = vbNullString And .Classification <> "unclassified" Then
            MsgBox "Please enter security sticker number", vbExclamation
            frmUserDetails.lvSecurityStickers.SetFocus
        Else
            ValidateWorkstationDetails = True
        End If
    End With
End Function

Public Function ValidateHDDDetails() As Boolean
    
    Dim i As Integer
    
    For i = 0 To UBound(HW_HARDDISKS)
        With HW_HARDDISKS(i)
            If .Removable = False Then
            
                If .SerialNumber = vbNullString And .Registry.InventarySerialNum = vbNullString Then
                        MsgBox "Please enter hard disk drive serial number.", _
                            vbExclamation
                            frmUserDetails.ZOrder 0
                            frmUserDetails.Show
                            frmUserDetails.lvHardDisks.ListItems("HDD" & .Index).Selected = True
                            frmUserDetails.lvHardDisks.ListItems("HDD" & .Index).EnsureVisible
                            frmUserDetails.lvHardDisks_DblClick
                End If
                
                Select Case REG_WORKSTATION.Classification
                                    
                    Case "unclassified"
                        If .Registry.InventaryNo = vbNullString Or _
                           .Registry.InventaryDate = vbNullString Then
                                MsgBox "Please enter hard disk drive inventary number and inventrary date.", vbExclamation
                                frmUserDetails.ZOrder 0
                                frmUserDetails.Show
                                frmUserDetails.lvHardDisks.ListItems("HDD" & .Index).Selected = True
                                frmUserDetails.lvHardDisks.ListItems("HDD" & .Index).EnsureVisible
                                frmUserDetails.lvHardDisks_DblClick
                                Exit Function
                        End If
                    Case Else
                        If .Registry.RegistryNo = vbNullString Or _
                           .Registry.InventaryNo = vbNullString Or _
                           .Registry.InventaryDate = vbNullString Or _
                           .Registry.AdminSticker = vbNullString Or _
                           .Registry.RegistrySticker = vbNullString Then
                                MsgBox "Please enter hard disk drive data correctly.", vbExclamation
                                frmUserDetails.ZOrder 0
                                frmUserDetails.Show
                                frmUserDetails.lvHardDisks.ListItems("HDD" & .Index).Selected = True
                                frmUserDetails.lvHardDisks.ListItems("HDD" & .Index).EnsureVisible
                                frmUserDetails.lvHardDisks_DblClick
                                Exit Function
                        End If
                End Select
            End If
        End With
    Next
    
    ValidateHDDDetails = True
    
End Function

Public Function ValidateUserDetails() As Boolean

        ValidateUserDetails = False
        
        If UBound(REG_USERS) = 0 And REG_USERS(0).Rank = vbNullString Then
            MsgBox "Add minimum one workstation user (owner)", vbCritical
            frmUserDetails.cmbRanks.SetFocus
        ElseIf modUserDetails.bWorkstationHasOwner = False Then
            MsgBox "Coose workstation's owner", vbExclamation
            frmUserDetails.lvUsers.SetFocus
        Else
            ValidateUserDetails = True
        End If

End Function
