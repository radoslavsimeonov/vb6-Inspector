Attribute VB_Name = "modIcons"
Option Explicit

Public Type SP_CLASSIMAGELIST_DATA
    cbSize    As Long
    ImageList As Long
    reserved  As Long
End Type

Private Declare Function SetupDiGetClassImageIndex _
                Lib "setupapi.dll" (ByRef ClassImageListData As SP_CLASSIMAGELIST_DATA, _
                                    ByRef ClassGuid As GUID, _
                                    ByRef ImageIndex As Long) As Long

Private Declare Function SetupDiGetClassImageList _
                Lib "setupapi.dll" (ByRef ClassImageListData As SP_CLASSIMAGELIST_DATA) As Long

Private Declare Function SetupDiDestroyClassImageList _
                Lib "setupapi.dll" (ByRef ClassImageListData As SP_CLASSIMAGELIST_DATA) As Long

Private Declare Function ImageList_GetImageCount _
                Lib "comctl32.dll" (ByVal himl As Long) As Long

Private Declare Function ImageList_GetIcon _
                Lib "comctl32.dll" (ByVal himl As Long, _
                                    ByVal i As Long, _
                                    ByVal Flags As Long) As Long

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Private Const ILD_NORMAL As Long = &H0
Private Const PICTYPE_ICON = 3

Private Type PICTDESC
    cbSizeofStruct  As Long
    picType         As Long
    hImage      As Long
    xExt        As Long
    yExt        As Long
End Type

Private Const S_OK                    As Long = 0
Private Const MAX_PATH                As Long = 260
Private Const SHGFI_ICON              As Long = &H100&
Private Const SHGFI_LARGEICON         As Long = &H0&  '32x32 pixels.
Private Const SHGFI_SMALLICON         As Long = &H1&  '16x16 pixels.
Private Const SHGFI_USEFILEATTRIBUTES As Long = &H10&

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Type PictDesc_Icon
    cbSizeofStruct As Long
    picType As Long
    hIcon As Long
End Type

Private Declare Function SHGetFileInfo _
                Lib "shell32" _
                Alias "SHGetFileInfoW" (ByVal pszPath As Long, _
                                        ByVal dwFileAttributes As Long, _
                                        ByVal psfi As Long, _
                                        ByVal cbSizeFileInfo As Long, _
                                        ByVal uFlags As Long) As Long

Private Declare Function GetFullPathName _
                Lib "KERNEL32" _
                Alias "GetFullPathNameW" (ByVal lpFileName As Long, _
                                          ByVal nBufferLength As Long, _
                                          ByVal lpBuffer As Long, _
                                          ByVal lpFilePart As Long) As Long

Private Declare Function OleCreatePictureIndirect _
                Lib "olepro32" (ByRef lpPictDesc As PICTDESC, _
                                ByRef riid As GUID, _
                                byvalfOwn As Long, _
                                ByRef lplpvObj As IPicture) As Long

Private IPictureIID As GUID

Public Function HICON2StdPicture(hIcon As Long) As IPictureDisp

    Dim IID_IPictureDisp As GUID
    Dim lpIcon           As PICTDESC

    IID_IPictureDisp.Data1 = &H7BF80981
    IID_IPictureDisp.Data2 = &HBF32
    IID_IPictureDisp.Data3 = &H101A
    IID_IPictureDisp.Data4(0) = &H8B
    IID_IPictureDisp.Data4(1) = &HBB
    IID_IPictureDisp.Data4(2) = &H0
    IID_IPictureDisp.Data4(3) = &HAA
    IID_IPictureDisp.Data4(4) = &H0
    IID_IPictureDisp.Data4(5) = &H30
    IID_IPictureDisp.Data4(6) = &HC
    IID_IPictureDisp.Data4(7) = &HAB
    lpIcon.cbSizeofStruct = Len(lpIcon)
    lpIcon.hImage = hIcon
    lpIcon.picType = PICTYPE_ICON
    OleCreatePictureIndirect lpIcon, IID_IPictureDisp, 0, HICON2StdPicture
End Function

Public Function FillImageListWithClassImageList(ClassImage As SP_CLASSIMAGELIST_DATA, _
                                                ImList As ImageList) As Long

    Dim numImage As Long
    Dim x        As Long
    Dim hIcon    As Long

    numImage = ImageList_GetImageCount(ClassImage.ImageList)

    For x = 0 To numImage - 1
        hIcon = ImageList_GetIcon(ClassImage.ImageList, x, ILD_NORMAL)
        ImList.ListImages.Add x + 1, , HICON2StdPicture(hIcon)
        DestroyIcon hIcon
    Next

End Function

Public Function GetClassImageList() As SP_CLASSIMAGELIST_DATA
    GetClassImageList.cbSize = Len(GetClassImageList)
    SetupDiGetClassImageList GetClassImageList
End Function

Public Sub DestroyClassImageList(ClassImage As SP_CLASSIMAGELIST_DATA)
    SetupDiDestroyClassImageList ClassImage
End Sub

Public Function GetClassImageListIndex(ClassImage As SP_CLASSIMAGELIST_DATA, _
                                       ClassGuid As String) As Long

    Dim tGUID As GUID

    Call IIDFromString(StrPtr(ClassGuid), tGUID)
    SetupDiGetClassImageIndex ClassImage, tGUID, GetClassImageListIndex
    GetClassImageListIndex = GetClassImageListIndex + 1
End Function

Public Function GetAssocIcon(ByVal PathToFile As String, _
                             Optional ByVal LargeIcon As Boolean = False, _
                             Optional ByVal Extension As Boolean = False) As StdPicture

    Dim SFI  As SHFILEINFO
    Dim Desc As PICTDESC

    If Len(PathToFile) = 0 And Extension Then PathToFile = "x" 'Win7 "generic icon" request fix.
    If Not FExists(PathToFile) Then PathToFile = "x"
    
    If SHGetFileInfo(StrPtr(PathToFile), 0, VarPtr(SFI), LenB(SFI), SHGFI_ICON Or IIf(LargeIcon, SHGFI_LARGEICON, SHGFI_SMALLICON) Or IIf(Extension, SHGFI_USEFILEATTRIBUTES, 0)) = 0 Then
        Exit Function
    End If

    If IPictureIID.Data1 = 0 Then

        With IPictureIID
            .Data1 = &H7BF80980
            .Data2 = &HBF32
            .Data3 = &H101A
            .Data4(0) = &H8B
            .Data4(1) = &HBB
            .Data4(2) = &H0
            .Data4(3) = &HAA
            .Data4(4) = &H0
            .Data4(5) = &H30
            .Data4(6) = &HC
            .Data4(7) = &HAB
        End With

    End If

    With Desc
        .cbSizeofStruct = Len(Desc)
        .picType = vbPicTypeIcon
        .hImage = SFI.hIcon
    End With

    If OleCreatePictureIndirect(Desc, IPictureIID, True, GetAssocIcon) <> S_OK Then
        Set GetAssocIcon = Nothing
    End If

End Function
