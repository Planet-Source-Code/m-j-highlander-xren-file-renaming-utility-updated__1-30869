Attribute VB_Name = "Shell_File_Operation"
Option Explicit

'SHFileOperation declarations
Const FO_MOVE = 1
Const FO_COPY = 2
Const FO_DELETE = 3
Const FO_RENAME = 4

Const FOF_MULTIDESTFILES = &H1      'Destination specifies multiple files
Const FOF_SILENT = &H4              'Don't display progress dialog
Const FOF_RENAMEONCOLLISION = &H8   'Rename if destination already exists
Const FOF_NOCONFIRMATION = &H10     'Don't prompt user
Const FOF_WANTMAPPINGHANDLE = &H20  'Fill in hNameMappings member
Const FOF_ALLOWUNDO = &H40          'Store undo information if possible
Const FOF_FILESONLY = &H80          'On *.*, don't copy directories
Const FOF_SIMPLEPROGRESS = &H100    'Don't show name of each file
Const FOF_NOCONFIRMMKDIR = &H200    'Don't confirm making any needed dirs

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String 'Used only if FOF_SIMPLEPROGRESS specified
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Sub SHRename(sOldFileName As String, sNewFileNAme As String)
    Dim FileOp As SHFILEOPSTRUCT
    
    'Parent window of dialog box--just use 0
    FileOp.hwnd = 0
    
    'Operation to perform
    FileOp.wFunc = FO_RENAME
    
  
        FileOp.pFrom = sOldFileName & Chr$(0)
        FileOp.pTo = sNewFileNAme & Chr$(0)
        FileOp.fFlags = FOF_ALLOWUNDO
        'FileOp.fFlags = FileOp.fFlags Or FOF_RENAMEONCOLLISION
   
    'Perform SHFileOperation
    If SHFileOperation(FileOp) <> 0 Then
                    MsgBox "Did not complete operation successfully!"
    End If

End Sub

