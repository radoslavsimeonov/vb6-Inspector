Attribute VB_Name = "modVisualTweaksLV2TV"
Option Explicit

Public Type POINTAPI   ' pt
    x As Long
    Y As Long
End Type

Public Type RECT   ' rct
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function ScreenToClient _
        Lib "user32" (ByVal hWnd As Long, _
                      lpPoint As POINTAPI) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As KeyCodeConstants) As Integer
Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" (ByVal hWnd As Long, _
                              ByVal wMsg As Long, _
                              ByVal wParam As Long, _
                              lParam As Any) As Long    ' <---
Declare Function PostMessage _
        Lib "user32" _
        Alias "PostMessageA" (ByVal hWnd As Long, _
                              ByVal wMsg As Long, _
                              ByVal wParam As Long, _
                              lParam As Any) As Long    ' <---

Public Const LVI_NOITEM = -1

Public Enum LVItemStates
    lvisNoButton = 0
    lvisCollapsed = 1
    lvisExpanded = 2
End Enum

Public Enum LVRelativeItemFlags
    lvriParent = 0
    lvriChild = 1
    lvriFirstSibling = 2
    lvriLastSibling = 3
    lvriPrevSibling = 4
    lvriNextSibling = 5
End Enum

Public Enum LVItemCountFlags
    lvicParents = 0
    lvrcChildren = 1
    lvicSiblings = 2
End Enum

#Const WIN32_IE = &H300

Public Const LVM_FIRST = &H1000
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Public Const LVM_GETITEM = (LVM_FIRST + 5)
Public Const LVM_SETITEM = (LVM_FIRST + 6)
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Public Const LVM_HITTEST = (LVM_FIRST + 18)
Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
#If (WIN32_IE >= &H300) Then

    Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
#End If

Public Const LVSIL_STATE = 2

Public Type LVITEM   ' was LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long  ' if String, must be pre-allocated before before filled
    cchTextMax As Long
    iImage As Long
    lParam As Long
    #If (WIN32_IE >= &H300) Then
        iIndent As Long
    #End If
End Type

Public Const LVIF_STATE = &H8
#If (WIN32_IE >= &H300) Then

    Public Const LVIF_INDENT = &H10
#End If

Public Const LVIS_FOCUSED = &H1
Public Const LVIS_SELECTED = &H2
Public Const LVIS_STATEIMAGEMASK = &HF000

Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2

Public Type LVHITTESTINFO   ' was LV_HITTESTINFO
    pt As POINTAPI
    flags As Long
    iItem As Long
    #If (WIN32_IE >= &H300) Then
        iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
    #End If
End Type

Public Const LVHT_ONITEMICON = &H2
Public Const LVHT_ONITEMLABEL = &H4
Public Const LVHT_ONITEMSTATEICON = &H8
#If (WIN32_IE >= &H300) Then

    Public Const LVS_EX_FULLROWSELECT = &H20   ' // applies to report mode only
#End If

Public Function Listview_GetItemStateEx(hwndLV As Long, _
                                        iItem As Long, _
                                        iIndent As Long) As LVItemStates

    Dim lvi As LVITEM

    lvi.mask = LVIF_STATE Or LVIF_INDENT
    lvi.iItem = iItem
    lvi.stateMask = LVIS_STATEIMAGEMASK

    If ListView_GetItem(hwndLV, lvi) Then
        iIndent = lvi.iIndent
        Listview_GetItemStateEx = STATEIMAGEMASKTOINDEX(lvi.state And LVIS_STATEIMAGEMASK)
    End If

End Function

Public Function Listview_SetItemStateEx(hwndLV As Long, _
                                        iItem As Long, _
                                        iIndent As Long, _
                                        dwState As LVItemStates) As Boolean

    Dim lvi As LVITEM

    lvi.mask = LVIF_STATE Or LVIF_INDENT
    lvi.iItem = iItem
    lvi.state = INDEXTOSTATEIMAGEMASK(dwState)
    lvi.stateMask = LVIS_STATEIMAGEMASK
    lvi.iIndent = iIndent
    Listview_SetItemStateEx = ListView_SetItem(hwndLV, lvi)
End Function

Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
    ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
End Function

Public Function ListView_SetFocusedItem(hwndLV As Long, i As Long) As Boolean
    ListView_SetFocusedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, LVIS_FOCUSED Or LVIS_SELECTED)
End Function

Public Function ListView_GetRelativeItem(hwndLV As Long, _
                                         iItem As Long, _
                                         dwRelative As LVRelativeItemFlags) As Long

    Dim iIndentSrc    As Long
    Dim iIndentTarget As Long
    Dim i             As Long
    Dim nItems        As Long
    Dim iSave         As Long

    iIndentSrc = -1
    Call Listview_GetItemStateEx(hwndLV, iItem, iIndentSrc)

    If (iIndentSrc = -1) Then
        ListView_GetRelativeItem = LVI_NOITEM
        Exit Function
    End If

    i = iItem
    nItems = ListView_GetItemCount(hwndLV)

    Select Case dwRelative
        Case lvriParent

            Do
                i = i - 1

                If (i = LVI_NOITEM) Then Exit Do
                Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)
            Loop Until (iIndentTarget < iIndentSrc)

        Case lvriChild

            If (i = (nItems - 1)) Then
                i = LVI_NOITEM
            Else
                i = i + 1
                Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)

                If (iIndentTarget <= iIndentSrc) Then i = LVI_NOITEM
            End If

        Case lvriFirstSibling
            iSave = i

            Do
                i = i - 1

                If (i = LVI_NOITEM) Then Exit Do
                Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)

                If (iIndentTarget = iIndentSrc) Then
                    iSave = i
                ElseIf (iIndentTarget < iIndentSrc) Then
                    Exit Do
                End If

            Loop

            i = iSave
        Case lvriLastSibling
            iSave = i

            Do
                i = i + 1

                If (i = nItems) Then Exit Do
                Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)

                If (iIndentTarget = iIndentSrc) Then
                    iSave = i
                ElseIf (iIndentTarget < iIndentSrc) Then
                    Exit Do
                End If

            Loop

            i = iSave
        Case lvriPrevSibling

            Do
                i = i - 1

                If (i = LVI_NOITEM) Then Exit Do
                Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)

                If (iIndentTarget = iIndentSrc) Then
                    Exit Do
                ElseIf (iIndentTarget < iIndentSrc) Then
                    i = LVI_NOITEM
                    Exit Do
                End If

            Loop

        Case lvriNextSibling

            Do
                i = i + 1

                If (i = nItems) Then
                    i = LVI_NOITEM
                    Exit Do
                End If

                Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)

                If (iIndentTarget = iIndentSrc) Then
                    Exit Do
                ElseIf (iIndentTarget < iIndentSrc) Then
                    i = LVI_NOITEM
                    Exit Do
                End If

            Loop

        Case Else
            i = LVI_NOITEM
    End Select

    ListView_GetRelativeItem = i
End Function

Public Function ListView_GetItemCountEx(hwndLV As Long, _
                                        iItem As Long, _
                                        dwRelative As LVItemCountFlags) As Long

    Dim i      As Long
    Dim nItems As Long

    Select Case dwRelative
        Case lvicParents
            nItems = -1

            Do
                nItems = nItems + 1
                i = ListView_GetRelativeItem(hwndLV, i, lvriParent)
            Loop Until (i = LVI_NOITEM)

        Case lvrcChildren
            i = ListView_GetRelativeItem(hwndLV, i, lvriChild)

            Do Until (i = LVI_NOITEM)
                nItems = nItems + 1
                i = ListView_GetRelativeItem(hwndLV, i, lvriNextSibling)
            Loop

        Case lvicSiblings
            i = ListView_GetRelativeItem(hwndLV, i, lvriFirstSibling)

            Do Until (i = LVI_NOITEM)
                nItems = nItems + 1
                i = ListView_GetRelativeItem(hwndLV, i, lvriNextSibling)
            Loop

    End Select

    ListView_GetItemCountEx = nItems
End Function

Public Function ListView_SetImageList(hWnd As Long, _
                                      himl As Long, _
                                      iImageList As Long) As Long
    ListView_SetImageList = SendMessage(hWnd, LVM_SETIMAGELIST, ByVal iImageList, ByVal himl)
End Function

Public Function ListView_GetItemCount(hWnd As Long) As Long
    ListView_GetItemCount = SendMessage(hWnd, LVM_GETITEMCOUNT, 0, 0)
End Function

Public Function ListView_GetItem(hWnd As Long, pitem As LVITEM) As Boolean
    ListView_GetItem = SendMessage(hWnd, LVM_GETITEM, 0, pitem)
End Function

Public Function ListView_SetItem(hWnd As Long, pitem As LVITEM) As Boolean
    ListView_SetItem = SendMessage(hWnd, LVM_SETITEM, 0, pitem)
End Function

Public Function ListView_GetNextItem(hWnd As Long, i As Long, flags As Long) As Long
    ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal flags)    ' ByVal MAKELPARAM(flags, 0))
End Function

Public Function ListView_HitTest(hwndLV As Long, pinfo As LVHITTESTINFO) As Long
    ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pinfo)
End Function

Public Function ListView_EnsureVisible(hwndLV As Long, _
                                       i As Long, _
                                       fPartialOK As Boolean) As Boolean
    ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal Abs(fPartialOK))    ' ByVal MAKELPARAM(Abs(fPartialOK), 0))
End Function

Public Function ListView_SetColumnWidth(hWnd As Long, iCol As Long, cx As Long) As Boolean
    ListView_SetColumnWidth = SendMessage(hWnd, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal cx)    ' ByVal MAKELPARAM(cx, 0))
End Function

Public Function ListView_SetItemState(hwndLV As Long, _
                                      i As Long, _
                                      state As Long, _
                                      mask As Long) As Boolean

    Dim lvi As LVITEM

    lvi.state = state
    lvi.stateMask = mask
    ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
End Function

#If (WIN32_IE >= &H300) Then
    Public Function ListView_SetExtendedListViewStyleEx(hwndLV As Long, _
                                                        dwMask As Long, _
                                                        dw As Long) As Long
        ListView_SetExtendedListViewStyleEx = SendMessage(hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, ByVal dwMask, ByVal dw)
    End Function

#End If   ' (WIN32_IE >= &H300)
Public Function INDEXTOSTATEIMAGEMASK(iIndex As Long) As Long
    INDEXTOSTATEIMAGEMASK = iIndex * (2 ^ 12)
End Function

Public Function STATEIMAGEMASKTOINDEX(iState As Long) As Long
    STATEIMAGEMASKTOINDEX = iState / (2 ^ 12)
End Function
