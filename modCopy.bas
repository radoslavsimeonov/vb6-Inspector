Attribute VB_Name = "modCopy"
Option Explicit

Private Const ERROR_UNKNOWN = 99999

Private Const ERROR_SUCCESS = 0
Private Const ERROR_FILE_NOT_FOUND = 2
Private Const ERROR_PATH_NOT_FOUND = 3
Private Const ERROR_TOO_MANY_OPEN_FILES = 4
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_HANDLE = 6
Private Const ERROR_BAD_FORMAT = 11
Private Const ERROR_OUTOFMEMORY = 14
Private Const ERROR_NO_MORE_FILES = 18
Private Const ERROR_SHARING_VIOLATION = 32
Private Const ERROR_DUP_NAME = 52
Private Const ERROR_FILE_EXISTS = 80
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_BROKEN_PIPE = 109
Private Const ERROR_DISK_FULL = 112
Private Const ERROR_CALL_NOT_IMPLEMENTED = 120
Private Const ERROR_SEEK_ON_DEVICE = 132
Private Const ERROR_DIR_NOT_EMPTY = 145
Private Const ERROR_BUSY = 170
Private Const ERROR_ALREADY_EXISTS = 183
Private Const ERROR_FILENAME_EXCED_RANGE = 206
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const ERROR_IO_DEVICE = 1117
Private Const ERROR_POSSIBLE_DEADLOCK = 1131
Private Const ERROR_BAD_DEVICE = 1200

Private Const FO_COPY = &H2&   'Copies the files specified
                                        'in the pFrom member to the
                                        'location specified in the
                                        'pTo member.

Private Const FO_DELETE = &H3& 'Deletes the files specified
                                        'in pFrom (pTo is ignored.)
         
Private Const FO_MOVE = &H1&   'Moves the files specified
                                        'in pFrom to the location
                                        'specified in pTo.
         
Private Const FO_RENAME = &H4& 'Renames the files
                                        'specified in pFrom.
         
Private Const FOF_ALLOWUNDO = &H40&   'Preserve Undo information.
Private Const FOF_CONFIRMMOUSE = &H2& 'Not currently implemented.
Private Const FOF_CREATEPROGRESSDLG = &H0& 'handle to the parent
                                                    'window for the
                                                    'progress dialog box.
         
Private Const FOF_FILESONLY = &H80&        'Perform the operation
                                                    'on files only if a
                                                    'wildcard file name
                                                    '(*.*) is specified.
         
Private Const FOF_MULTIDESTFILES = &H1&    'The pTo member
                                                    'specifies multiple
                                                    'destination files (one
                                                    'for each source file)
                                                    'rather than one
                                                    'directory where all
                                                    'source files are
                                                    'to be deposited.
         
Private Const FOF_NOCONFIRMATION = &H10&   'Respond with Yes to
                                                    'All for any dialog box
                                                    'that is displayed.
         
Private Const FOF_NOCONFIRMMKDIR = &H200&  'Does not confirm the
                                                    'creation of a new
                                                    'directory if the
                                                    'operation requires one
                                                    'to be created.
         
Private Const FOF_RENAMEONCOLLISION = &H8& 'Give the file being
                                                    'operated on a new name
                                                    'in a move, copy, or
                                                    'rename operation if a
                                                    'file with the target
                                                    'name already exists.
         
Private Const FOF_SILENT = &H4&            'Does not display a
                                                    'progress dialog box.
         
Private Const FOF_SIMPLEPROGRESS = &H100&  'Displays a progress
                                                    'dialog box but does
                                                    'not show the
                                                    'file names.
         
Private Const FOF_WANTMAPPINGHANDLE = &H20&
                                   'If FOF_RENAMEONCOLLISION is specified,
                                   'the hNameMappings member will be filled
                                   'in if any files were renamed.
         ' The SHFILOPSTRUCT is not double-word aligned. If no steps are
         ' taken, the last 3 variables will not be passed correctly. This
         ' has no impact unless the progress title needs to be changed.
         
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type
         
Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" (hpvDest As Any, _
hpvSource As Any, ByVal cbCopy As Long)
         
Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As Any) As Long

Public Function CopySource(strSource As String, _
                           strDest As String, _
                           bDlg As Boolean) As Boolean

    Dim strIPRetr As String
    Dim Result    As Long
    Dim lenFileop As Long
    Dim foBuf()   As Byte
    Dim fileop    As SHFILEOPSTRUCT

    lenFileop = LenB(fileop)    ' double word alignment increase
    ReDim foBuf(1 To lenFileop) ' the size of the structure.

    'MakeSureDirectoryPathExists strDest
    With fileop
        .hwnd = frmPolicy.hwnd
        .wFunc = FO_MOVE
        .pFrom = strSource
        .pTo = strDest

        If bDlg = True Then
            .fFlags = FOF_NOCONFIRMMKDIR + FOF_NOCONFIRMATION
        Else
            .fFlags = FOF_SILENT + FOF_NOCONFIRMMKDIR + FOF_NOCONFIRMATION
        End If
                          
        .lpszProgressTitle = "Please Be Patient...working... " & vbNullChar & vbNullChar
    End With

    'Call the CopyMemory procedure and copy the
    'structure into a byte array(??)

    Call CopyMemory(foBuf(1), fileop, lenFileop)
    Call CopyMemory(foBuf(19), foBuf(21), 12)
    
    'This is used for Error Handling - 0 means everything was
    'fine, any other number means their was an error...
    Result = SHFileOperation(foBuf(1))
    
    'ERROR HANDLER....
        
    If Result = 0 Then
        'MsgBox " Successful"

        CopySource = True
    ElseIf Result = 5 Then
        'MsgBox " Access Denied"
        CopySource = False
    ElseIf Result = 112 Then
        'MsgBox " Disk Full"
        CopySource = False
    ElseIf Result = 3 Then
        'MsgBox " Path Not Found"
        CopySource = False
    ElseIf Result <> 0 Then
        'MsgBox " Failed"
        CopySource = False
    End If

End Function

