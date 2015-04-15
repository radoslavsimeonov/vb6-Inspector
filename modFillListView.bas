Attribute VB_Name = "modFillListView"
Option Explicit

Private P_ListView As ListView

Public Property Let MY_ListView(ByVal vNewValue As ListView)
    Set P_ListView = vNewValue
End Property

Public Sub AddListItem(sValue As Variant, sField As String, Optional bBold As Boolean = False, Optional bIndent As Boolean = False, Optional subColor As OLE_COLOR = vbWhite, Optional sImg As String = vbNullString)

    Const sIndent As String = "    "
    Static IsPrevEmpty As Boolean
    Dim itmX As ListItem
        
    If Trim$(sValue) = vbNullString And Trim$(sField) <> vbNullString Then Exit Sub
    
    If sValue = vbNullString And sField = vbNullString Then
        If IsPrevEmpty Then Exit Sub
        IsPrevEmpty = True
    Else
        IsPrevEmpty = False
    End If
    
    If sImg <> vbNullString Then
        Set itmX = P_ListView.ListItems.Add(, , IIf(bIndent, "", sIndent) & sField, , sImg)
    Else
        Set itmX = P_ListView.ListItems.Add(, , IIf(bIndent, "", sIndent) & sField)
    End If
    
    itmX.ForeColor = &H80FF&        '&HC0C0C0
    itmX.Bold = bBold
    itmX.SubItems(1) = Trim$(sValue)
    itmX.ListSubItems(1).Bold = bBold
    itmX.ListSubItems(1).ForeColor = subColor
End Sub
