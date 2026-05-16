Attribute VB_Name = "mdlMain"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1    'bring to top and stay there
Public Const SWP_NOMOVE = &H2    'don't move window
Public Const SWP_NOSIZE = &H1    'don't size window

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CreateDirectoryEx Lib "kernel32.dll" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDListA Lib "shell32.dll" (pidl As Any, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ppidl As ITEMIDLIST) As Long

Public Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Enum BROWSE_FLAGS
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_EDITBOX = &H10
    BIF_RETURNFSANCESTORS = &H8
    BIF_RETURNONLYFSDIRS = &H1
    BIF_STATUSTEXT = &H4
    BIF_USENEWUI = &H40
    BIF_VALIDATE = &H20
End Enum

Public Const MAX_PATH = 260
Public Const MAX_NAME = 40

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function FileExists(Filename As String) As Boolean
    Dim FileNum As Integer
    On Error GoTo FileDoesNotExist
    FileNum = FreeFile
    Open Filename For Input As #FileNum
    Close #1
    FileExists = True
    Exit Function
FileDoesNotExist:     FileExists = False
End Function
