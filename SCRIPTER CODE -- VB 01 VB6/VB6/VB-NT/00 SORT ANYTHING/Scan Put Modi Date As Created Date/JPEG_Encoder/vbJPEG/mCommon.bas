Attribute VB_Name = "mCommon"

Option Explicit
Public FileName As String

'Public Procedures
'   FileDialog()             As String
'   BrowseFolders()          As String
'   RidFile()
'   FileExists()             As Boolean
'   DirectoryExists()        As Boolean
'   FileTitleOnly()          As String
'   GetShortName()           As String
'   CreateTempFile()         As String
'   CreatePath()             As Boolean

'Public Properties [Get Only]
'   WindowsDirectory()       As String
'   WindowsTempDirectory()   As String
'   WindowsSystemDirectory() As String
'   AppPath()                As String

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type
Private Type BrowseInfo
    hwndOwner         As Long
    pIDLRoot          As Long
    pszDisplayName    As Long
    lpszTitle         As Long
    ulFlags           As Long
    lpfnCallback      As Long
    lParam            As Long
    iImage            As Long
End Type

Private Const OFN_READONLY             As Long = &H1
Private Const OFN_OVERWRITEPROMPT      As Long = &H2
Private Const OFN_HIDEREADONLY         As Long = &H4
Private Const OFN_NOCHANGEDIR          As Long = &H8
Private Const OFN_SHOWHELP             As Long = &H10
Private Const OFN_ENABLEHOOK           As Long = &H20
Private Const OFN_ENABLETEMPLATE       As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Private Const OFN_NOVALIDATE           As Long = &H100
Private Const OFN_ALLOWMULTISELECT     As Long = &H200
Private Const OFN_EXTENSIONDIFFERENT   As Long = &H400
Private Const OFN_PATHMUSTEXIST        As Long = &H800
Private Const OFN_FILEMUSTEXIST        As Long = &H1000
Private Const OFN_CREATEPROMPT         As Long = &H2000
Private Const OFN_SHAREAWARE           As Long = &H4000
Private Const OFN_NOREADONLYRETURN     As Long = &H8000
Private Const OFN_NOTESTFILECREATE     As Long = &H10000
Private Const OFN_NONETWORKBUTTON      As Long = &H20000
Private Const OFN_NOLONGNAMES          As Long = &H40000
Private Const OFN_EXPLORER             As Long = &H80000
Private Const OFN_NODEREFERENCELINKS   As Long = &H100000
Private Const OFN_LONGNAMES            As Long = &H200000

Private Const OFN_SHAREFALLTHROUGH     As Long = 2
Private Const OFN_SHARENOWARN          As Long = 1
Private Const OFN_SHAREWARN            As Long = 0

Private Const BrowseForFolders         As Long = &H1
Private Const BrowseForComputers       As Long = &H1000
Private Const BrowseForPrinters        As Long = &H2000
Private Const BrowseForEverything      As Long = &H4000

Private Const CSIDL_BITBUCKET          As Long = 10
Private Const CSIDL_CONTROLS           As Long = 3
Private Const CSIDL_DESKTOP            As Long = 0
Private Const CSIDL_DRIVES             As Long = 17
Private Const CSIDL_FONTS              As Long = 20
Private Const CSIDL_NETHOOD            As Long = 18
Private Const CSIDL_NETWORK            As Long = 19
Private Const CSIDL_PERSONAL           As Long = 5
Private Const CSIDL_PRINTERS           As Long = 4
Private Const CSIDL_PROGRAMS           As Long = 2
Private Const CSIDL_RECENT             As Long = 8
Private Const CSIDL_SENDTO             As Long = 9
Private Const CSIDL_STARTMENU          As Long = 11

Private Const MAX_PATH                 As Long = 260

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ListId As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long


'==============================================================================================
'                                Selecting Files / Directories
'==============================================================================================
Public Function FileDialog(FormObject As Form, SaveDialog As Boolean, ByVal Title As String, ByVal Filter As String, Optional ByVal FileName As String, Optional ByVal Extention As String, Optional ByVal InitDir As String) As String
    Dim OFN   As OPENFILENAME
    Dim R     As Long

    If Len(FileName) > MAX_PATH Then Call MsgBox("Filename Length Overflow", vbExclamation, App.Title + " - FileDialog Function"): Exit Function

    FormObject.Enabled = False
    FileName = FileName + String(MAX_PATH - Len(FileName), 0)

    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = FormObject.hwnd
        .hInstance = App.hInstance
        .lpstrFilter = Replace(Filter, "|", vbNullChar)
        .lpstrFile = FileName
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = Space$(MAX_PATH - 1)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = InitDir
        .lpstrTitle = Title
        .flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
        .lpstrDefExt = Extention
    End With

Dim l As Long
l = GetTickCount

    If SaveDialog Then R = GetSaveFileName(OFN) Else R = GetOpenFileName(OFN)

If GetTickCount - l < 20 Then
OFN.lpstrFile = ""
If SaveDialog Then R = GetSaveFileName(OFN) Else R = GetOpenFileName(OFN)
End If

    If R = 1 Then FileDialog = Left$(OFN.lpstrFile, InStr(1, OFN.lpstrFile + vbNullChar, vbNullChar) - 1)
    FormObject.Enabled = True

End Function
Public Function BrowseFolders(FormObject As Form, sMessage As String) As String
    Dim B As BrowseInfo
    Dim R As Long
    Dim l As Long
    Dim f As String

    FormObject.Enabled = False
    With B
        .hwndOwner = FormObject.hwnd
        .lpszTitle = lstrcat(sMessage, "")
        .ulFlags = BrowseForFolders
    End With

    SHGetSpecialFolderLocation FormObject.hwnd, CSIDL_DRIVES, B.pIDLRoot
    R = SHBrowseForFolder(B)

    If R <> 0 Then     'A zero would mean cancel was pressed
        f = String(MAX_PATH, vbNullChar)
        SHGetPathFromIDList R, f
        CoTaskMemFree R
        l = InStr(1, f, vbNullChar) - 1
        If l < 0 Then l = 0
        f = Left(f, l)
        AddSlash f
    End If

    BrowseFolders = f
    FormObject.Enabled = True

End Function



'==============================================================================================
'                       Getting System Directories / Executable Filename
'==============================================================================================
Public Property Get WindowsDirectory() As String
    Static R As String
    If Len(R) = 0 Then
        Dim l As Long
        l = MAX_PATH
        R = String(l, 0)
        l = GetWindowsDirectory(R, l)
        If l > 0 Then
            R = Left$(R, l)
            AddSlash R
        Else
            R = ""
        End If
    End If
    WindowsDirectory = R
End Property
Public Property Get WindowsTempDirectory() As String
    Static m_WindowsTempDirectory As String
    If Len(m_WindowsTempDirectory) = 0 Then
        Dim Buffer As String
        Dim Length As Long
        Buffer = String(MAX_PATH, 0)
        Length = GetTempPath(MAX_PATH, Buffer)
        If Length > 0 Then
            m_WindowsTempDirectory = Left$(Buffer, Length)
            AddSlash m_WindowsTempDirectory
        End If
    End If
    WindowsTempDirectory = m_WindowsTempDirectory
End Property
Public Property Get WindowsSystemDirectory() As String
    Static m_WindowsSystemDirectory As String
    If Len(m_WindowsSystemDirectory) = 0 Then
        Dim Buffer As String
        Dim Length As Long
        Buffer = String(MAX_PATH, 0)
        Length = GetSystemDirectory(Buffer, MAX_PATH)
        If Length > 0 Then
            m_WindowsSystemDirectory = Left$(Buffer, Length)
            AddSlash m_WindowsSystemDirectory
        End If
    End If
    WindowsSystemDirectory = m_WindowsSystemDirectory
End Property
Public Property Get AppPath() As String
    Static m_AppPath As String 'Returns Program EXE File Name
    If Len(m_AppPath) = 0 Then
        Dim ret As Long
        Dim Length As Long
        Dim FilePath As String
        Dim FileHandle As Long
        FilePath = String(MAX_PATH, 0)
        FileHandle = GetModuleHandle(App.EXEName)
        ret = GetModuleFileName(FileHandle, FilePath, MAX_PATH)
        Length = InStr(1, FilePath, vbNullChar) - 1
        If Length > 0 Then m_AppPath = Left$(FilePath, Length)
    End If
    AppPath = m_AppPath
End Property




'==============================================================================================
'                                      File Name Procedures
'==============================================================================================
Public Function FileExists(FileName As String) As Boolean
    If Len(FileName) > 0 Then FileExists = (Len(Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function
Public Function DirectoryExists(ByVal Directory As String) As Boolean
    AddSlash Directory
    DirectoryExists = Len(Directory) > 0 And Len(Dir(Directory + "*.*", vbDirectory)) > 0
End Function
Public Function FileTitleOnly(FileName As String, Optional ReturnDirectory As Boolean) As String
    If ReturnDirectory Then
        FileTitleOnly = Left$(FileName, InStrRev(FileName, "\"))
    Else
        FileTitleOnly = Right$(FileName, Len(FileName) - InStrRev(FileName, "\"))
    End If
End Function
Public Sub AddSlash(Directory As String)
    If InStrRev(Directory, "\") <> Len(Directory) Then Directory = Directory + "\"
End Sub
Public Sub RemoveSlash(Directory As String)
    If Len(Directory) > 3 And InStrRev(Directory, "\") = Len(Directory) Then Directory = Left$(Directory, Len(Directory) - 1)
End Sub
Public Sub RidFile(FileName As String)
    If FileExists(FileName) Then
        SetAttr FileName, vbNormal
        Kill FileName
    End If
End Sub
Public Function GetShortName(ByVal FileName As String) As String
    Dim Buffer As String
    Dim Length As Long
    Buffer = String(MAX_PATH, 0)
    Length = GetShortPathName(FileName, Buffer, MAX_PATH)
    If Length > 0 Then GetShortName = Left$(Buffer, Length)
End Function
Public Function CreateTempFile(Optional ByVal Prefix As String, Optional Directory As String) As String
    Dim Buffer As String 'This code will CREATE a new temp file with a unique filename
    Dim Length As Long
    Buffer = String(MAX_PATH, 0)
    If Len(Prefix) = 0 Then Prefix = Left$(App.Title + "TMP", 3)
    If Not DirectoryExists(Directory) Then Directory = WindowsTempDirectory
    If GetTempFileName(Directory, Prefix, 0&, Buffer) = 0 Then Exit Function
    Length = InStr(1, Buffer, vbNullChar) - 1
    If Length > 0 Then CreateTempFile = Left$(Buffer, Length)
End Function
Public Function CreatePath(ByVal Path As String) As Boolean
    On Error GoTo Fail
    Dim i As Integer
    Dim s As String

    AddSlash Path
    Do
        i = InStr(i + 1, Path, "\")
        If i = 0 Then Exit Do
        s = Left$(Path, i - 1)
        If Not DirectoryExists(s) Then MkDir s
    Loop Until i = Len(Path)

    If DirectoryExists(Path) Then
        CreatePath = True
        Exit Function
    End If

Fail:
Call MsgBox(IIf(Err.Number = 0, "", "Error " + CStr(Err.Number) + ": " + Err.Description + vbCrLf) + "Could Not Create/Access Directory:" + vbCrLf + vbCrLf + Chr$(34) + Path + Chr$(34), vbExclamation, App.Title + " - CreatePath Function")

End Function

