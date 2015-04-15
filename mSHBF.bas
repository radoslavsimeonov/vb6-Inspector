Attribute VB_Name = "mSHBF"
Option Explicit

Private Type BROWSEINFO
        hWndOwner As Long           '親ウィンドウのハンドル
        pidlRoot As Long            'ディレクトリルート（アイテムID）（[定数１]参照）
        pszDisplayName As String    '表示名
        lpszTitle As String         'ダイアログのタイトル
        ulFlags As Long             'フラグ（[定数２]参照）
        lpfn As Long                'コールバック関数のポインタ
        lParam As String            'コールバック関数へ渡す任意のデータ
        iImage As Long              'システムイメージリスト
End Type

Private Const MAX_PATH As Long = 260 'パス最大サイズ
Public Enum SHBFClassID
    CSIDL_DESKTOP = &H0             'デスクトップ
    CSIDL_PROGRAMS = &H2            'プログラム
    CSIDL_CONTROLS = &H3            'コントロールパネル
    CSIDL_PRINTERS = &H4            'プリンター
    CSIDL_PERSONAL = &H5            'パーソナル（C:\My Documents...）
    CSIDL_FAVORITES = &H6           'お気に入り
    CSIDL_STARTUP = &H7             'スタートアップ
    CSIDL_RECENT = &H8              '最近使ったファイル
    CSIDL_SENDTO = &H9              '送る（C:\Windows\SendTo...）
    CSIDL_BITBUCKET = &HA           'ごみ箱
    CSIDL_STARTMENU = &HB           'スタートメニュー
    CSIDL_DESKTOPDIRECTORY = &H10   'デスクトップフォルダ
    CSIDL_DRIVES = &H11             'ドライブ
    CSIDL_NETWORK = &H12            'ネットワーク
    CSIDL_NETHOOD = &H13            '（C:\Windows\NetHood）
    CSIDL_FONTS = &H14              'フォント（C:\Windows\Fonts）
    CSIDL_TEMPLATES = &H15          'テンプレート（C:\Windows\ShellNew）
    CSIDL_COMMON_STARTUP = &H18
End Enum

Public Enum SHBFFlags
    BIF_RETURNONLYFSDIRS = &H1          'フォルダのみ選択可能
    BIF_DONTGOBELOWDOMAIN = &H2         'コンピューター非表示
    BIF_STATUSTEXT = &H4                'ステータス表示
    BIF_RETURNFSANCESTORS = &H8         'ファイルシステムのみ選択可能
    BIF_BROWSEFORCOMPUTER = &H1000      'コンピューターのみ選択可能
    BIF_BROWSEFORPRINTER = &H2000       'プリンターのみ選択可能
    BIF_BROWSEINCLUDEFILES = &H4000     '全て選択可能
End Enum

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32.dll" Alias "#162" (ByVal strPath As String) As Long

Private m_Title As String
Private m_hWnd As Long
Private m_RootFolder As Long
Private m_UserRootFolder As String
Private m_Flags As Long

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal vNewValue As String)
    m_Title = vNewValue
End Property

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Let hWnd(ByVal vNewValue As Long)
    m_hWnd = vNewValue
End Property

Public Property Get RootFolder() As SHBFClassID
    RootFolder = m_RootFolder
End Property

Public Property Let RootFolder(ByVal vNewValue As SHBFClassID)
    m_RootFolder = vNewValue
End Property

Public Property Get UserRootFolder() As String
    UserRootFolder = m_UserRootFolder
End Property

Public Property Let UserRootFolder(ByVal vNewValue As String)
    m_UserRootFolder = vNewValue
End Property

Public Property Get Flags() As SHBFFlags
    Flags = m_Flags
End Property

Public Property Let Flags(ByVal vNewValue As SHBFFlags)
    m_Flags = vNewValue
End Property

Public Function Show() As String
    Dim typeBRINF As BROWSEINFO
    Dim strPath As String
    Dim lngpidl As Long
    
    strPath = String$(MAX_PATH, vbNullChar)
    If m_Flags = 0 Then m_Flags = BIF_RETURNONLYFSDIRS
    
    With typeBRINF
        .hWndOwner = m_hWnd
        If m_RootFolder = -1 Then
            If UserRootFolder <> "" Then
                .pidlRoot = SHSimpleIDListFromPath(UserRootFolder)
            End If
        Else
            .pidlRoot = m_RootFolder
        End If
        .pszDisplayName = strPath
        .lpszTitle = m_Title & vbNullChar
        .ulFlags = m_Flags
    End With
    
    lngpidl = SHBrowseForFolder(typeBRINF)
        
    If typeBRINF.ulFlags And BIF_BROWSEFORCOMPUTER Then
        strPath = Left$(typeBRINF.pszDisplayName, InStr(typeBRINF.pszDisplayName, vbNullChar) - 1)
    Else
        If lngpidl = 0 Then
            strPath = vbNullString
        Else
            If SHGetPathFromIDList(lngpidl, strPath) = 0 Then
                strPath = vbNullString
            Else
                strPath = Left$(strPath, InStr(strPath, vbNullChar) - 1&)
            End If
        End If
    End If
    Show = strPath
End Function
