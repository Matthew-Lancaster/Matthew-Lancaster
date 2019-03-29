Attribute VB_Name = "modFileSystem"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' AUTHOR:       Nathan Moschkin                               '''''''
''''''' MODULE:       modCDFileSystem                               '''''''
''''''' DESCRIPTION:  Functions for interfacing with the Windows    '''''''
'''''''               file system.                                  '''''''
''''''' COPYRIGHT:    Datavations, LLC  2003                        '''''''
'''''''                                                             '''''''
'''''''                                                             '''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Option Base 0

Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type OPENFILENAMENT5
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    lpReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type

Public FILESYS_UNICODE_CHECKED As Boolean
Public FILESYS_UNICODE As Boolean

Public Const CDM_SETCONTROLTEXT = (&H400& + 104&)
Public Const CDM_HIDECONTROL = (&H400& + 105&)

Public Const IDOK = 1&
Public Const IDCANCEL = 2&

Public Const MAX_PATH = 260
Public Const WM_USER = &H400

Public Const CREATE_ALWAYS = 2
Public Const CREATE_NEW = 1

Public Const OPEN_ALWAYS = 4
Public Const OPEN_EXISTING = 3

Public Const GENERIC_ALL = &H10000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

Public Const TRUNCATE_EXISTING = 5

Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Const BIF_RETURNONLYFSDIRS = &H1   ' For finding a folder to start document searching
Public Const BIF_DONTGOBELOWDOMAIN = &H2 ' For starting the Find Computer
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_EDITBOX = &H10
Public Const BIF_VALIDATE = &H20 ' insist on valid result (or CANCEL)

Public Const BIF_BROWSEFORCOMPUTER = &H1000   ' Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000 ' Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000  ' Browsing for Everything

' message from browser

Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const BFFM_VALIDATEFAILED = 3  ' lParam:szPath ret:1(cont),0(EndDialog)
' Public Const BFFM_VALIDATEFAILEDW = 4  ' lParam:wzPath ret:1(cont),0(EndDialog)

' messages to browser

Public Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
Public Const BFFM_ENABLEOK = (WM_USER + 101)
Public Const BFFM_SETSELECTIONA = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW = (WM_USER + 103)
Public Const BFFM_SETSTATUSTEXTW = (WM_USER + 104)

Public Const SHGFI_ICON = &H100                        '' get icon
Public Const SHGFI_DISPLAYNAME = &H200                 '' get display name
Public Const SHGFI_TYPENAME = &H400                    '' get type name
Public Const SHGFI_ATTRIBUTES = &H800                  '' get attributes
Public Const SHGFI_ICONLOCATION = &H1000               '' get icon location
Public Const SHGFI_EXETYPE = &H2000                    '' return exe type
Public Const SHGFI_SYSICONINDEX = &H4000               '' get system icon index
Public Const SHGFI_LINKOVERLAY = &H8000                '' put a link overlay on icon
Public Const SHGFI_SELECTED = &H10000                  '' show icon in selected state
Public Const SHGFI_ATTR_SPECIFIED = &H20000            '' get only specified attributes
Public Const SHGFI_LARGEICON = &H0                     '' get large icon
Public Const SHGFI_SMALLICON = &H1                     '' get small icon
Public Const SHGFI_OPENICON = &H2                      '' get open icon
Public Const SHGFI_SHELLICONSIZE = &H4                 '' get shell size icon
Public Const SHGFI_PIDL = &H8                          '' pszPath is a pidl
Public Const SHGFI_USEFILEATTRIBUTES = &H10            '' use passed dwFileAttribute

'Public Const SFGAO_CANCOPY           DROPEFFECT_COPY '' Objects can be copied
'Public Const SFGAO_CANMOVE           DROPEFFECT_MOVE '' Objects can be moved
'Public Const SFGAO_CANLINK           DROPEFFECT_LINK '' Objects can be linked

Public Const SFGAO_CANRENAME = &H10&                  '' Objects can be renamed
Public Const SFGAO_CANDELETE = &H20&                  '' Objects can be deleted
Public Const SFGAO_HASPROPSHEET = &H40&               '' Objects have property sheets
Public Const SFGAO_DROPTARGET = &H100&                '' Objects are drop target
Public Const SFGAO_CAPABILITYMASK = &H177&
Public Const SFGAO_LINK = &H10000                     '' Shortcut (link)
Public Const SFGAO_SHARE = &H20000                    '' shared
Public Const SFGAO_READONLY = &H40000                 '' read-only
Public Const SFGAO_GHOSTED = &H80000                  '' ghosted icon
Public Const SFGAO_HIDDEN = &H80000                   '' hidden object
Public Const SFGAO_DISPLAYATTRMASK = &HF0000
Public Const SFGAO_FILESYSANCESTOR = &H10000000       '' It contains file system folder
Public Const SFGAO_FOLDER = &H20000000                '' It's a folder.
Public Const SFGAO_FILESYSTEM = &H40000000            '' is a file system thing (file/folder/root)
Public Const SFGAO_HASSUBFOLDER = &H80000000          '' Expandable in the map pane
Public Const SFGAO_CONTENTSMASK = &H80000000
Public Const SFGAO_VALIDATE = &H1000000               '' invalidate cached information
Public Const SFGAO_REMOVABLE = &H2000000              '' is this removeable media?
Public Const SFGAO_COMPRESSED = &H4000000             '' Object is compressed (use alt color)
Public Const SFGAO_BROWSABLE = &H8000000              '' is in-place browsable
Public Const SFGAO_NONENUMERATED = &H100000           '' is a non-enumerated object
Public Const SFGAO_NEWCONTENT = &H200000              '' should show bold in explorer tree


Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_INTERNET = &H1
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PRINTHOOD = &H1B
Public Const CSIDL_ALTSTARTUP = &H1D                          '' DBCS
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E                   '' DBCS
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22

Public Enum SpecialFolderConstants
    spfldDesktop = &H0
    spfldInternet = &H1
    spfldPrograms = &H2
    spfldControls = &H3
    spfldPrinters = &H4
    spfldPersonal = &H5
    spfldFavorites = &H6
    spfldStartUp = &H7
    spfldRecent = &H8
    spfldSendTo = &H9
    spfldBitBucket = &HA
    spfldStartMenu = &HB
    spfldDesktopDirectory = &H10
    spfldDrives = &H11
    spfldNetwork = &H12
    spfldNethood = &H13
    spfldFonts = &H14
    spfldTemplates = &H15
    spfldCommonStartMenu = &H16
    spfldCommonPrograms = &H17
    spfldCommonStartup = &H18
    spfldCommonDesktopDirectory = &H19
    spfldAppData = &H1A
    spfldPrintHood = &H1B
    spfldAltStartup = &H1D                          '' DBCS
    spfldCommonAltStartup = &H1E                   '' DBCS
    spfldCommonFavorites = &H1F
    spfldInternetCache = &H20
    spfldCookies = &H21
    spfldHistory = &H22
End Enum

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type


Public Type BROWSEINFO
    hWndOwner As Long
    pidlRoot As Long
    DisplayName As String
    lpszTitle As String
    ulFlags As Integer
    lpfn As Long
    lParam As Long
    iImage As Integer
End Type

Public Type FILETIME
        hi As Long
        lo As Long
End Type

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type


Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFilename(0 To (MAX_PATH - 1)) As Byte
        cAlternate(0 To 13) As Byte
End Type

Public Type WIN32_FIND_DATA_UNICODE
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFilename(0 To ((MAX_PATH * 2) - 1)) As Byte
        cAlternate(0 To 27) As Byte
End Type

Public Type MODFSFINDDATA
        Attributes As Long
        CreationTime As Date
        LastAccessTime As Date
        LastWriteTime As Date
        SizeHigh As Long
        SizeLow As Long
        FileName As String
        ShortName As String
End Type

Public Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal bLen As Long, ByVal lpszBuffer As String) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFilename As Long, lpFindFileData As WIN32_FIND_DATA_UNICODE) As Long
Public Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA_UNICODE) As Long

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFilename As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilename As String) As Long

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function pGetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFilename As String) As Long
Public Declare Function pSetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFilename As String, ByVal dwFileAttributes As Long) As Long

Public Declare Function pGetFileTime Lib "kernel32" Alias "GetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function pSetFileTime Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Public Declare Function pGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function pGetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Public Declare Function pGetFileType Lib "kernel32" Alias "GetFileType" (ByVal hFile As Long) As Long

Public Declare Function SHBrowseForFolder Lib "shell32" (bInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal lpPath As String, ByVal nFolder As Long, ByVal fCreate As Boolean) As Long

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Public Declare Function pMoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFilename As String, ByVal lpNewFilename As String) As Long
Public Declare Function pDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFilename As String) As Long
Public Declare Function pCopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFilename As String, ByVal lpNewFilename As String, ByVal bFailIfExists As Long) As Long

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As Any) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As Any) As Long

Public mSetDir As Boolean
Public mBrowseCurrent As String
Public BrowserLastFolder As Long

' SHBrowseForFolder callback function

' Since we need the absolute path, we take the path of the most recently
' selected Item.  The only way to do that is through processing of the
' BFFM_SELCHANGED event generated by SHBrowseForFolder, and retrieve the
' full path using the SHGetPathFromIDList and wParam

Public Function BrowseCallback(ByVal hwnd As Long, ByVal uMsg As Integer, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim u As Long, _
        aOut As String
    
    If (mSetDir = True) Then
    
        aOut = StrConv(mBrowseCurrent + String(MAX_PATH - Len(mBrowseCurrent), 0), vbFromUnicode)
        SendMessage hwnd, BFFM_SETSELECTIONA, 1&, ByVal StrPtr(aOut)
        mSetDir = False
    End If

    If uMsg = BFFM_SELCHANGED Then
        mBrowseCurrent = String(260, 0)
        SHGetPathFromIDList wParam, mBrowseCurrent
        BrowserLastFolder = wParam
    End If
        
End Function

Public Function GethIcon(ByVal FileName As String) As Long

    Dim sh As SHFILEINFO
    
    FileName = FileName + Chr(0)
    
    SHGetFileInfo FileName, 0&, sh, Len(sh), SHGFI_ICON
    
    GethIcon = sh.hIcon

End Function

' Opens the folder browser

Public Function FolderBrowse(ByVal hwnd As Long, Optional ByVal Title As String = "Browse for Folder", Optional ByVal InitialDirectory As String) As String
    Dim bInfo As BROWSEINFO, i As Long
    bInfo.ulFlags = &H30 + 13
    
    bInfo.hWndOwner = hwnd
    bInfo.lpszTitle = Title
    bInfo.DisplayName = String(260, 0)
    
    ' Get the address of the callback function
    ' must be passed indirectly, see notes above (GetProcAddress() function)
    
    i = GetProcAddress(AddressOf BrowseCallback)
    
    bInfo.lpfn = i
    
    If (InitialDirectory <> "") Then
        mBrowseCurrent = InitialDirectory
        mSetDir = True
    Else
        mSetDir = False
        mBrowseCurrent = ""
    End If
    
    i = SHBrowseForFolder(bInfo)
    If (i = 0) Then
        FolderBrowse = ""
        Exit Function
    End If
    
    i = InStr(1, mBrowseCurrent, Chr(0))
    FolderBrowse = CleanPath(GetAbsolutePathname(Mid(mBrowseCurrent, 1, i - 1)))
    
End Function

Public Function GetFileSize(ByVal FileName As String, lpSizeHigh As Long) As Long

    Dim hFile As Long, _
        lpSA As SECURITY_ATTRIBUTES
        
    Dim SL As Long, _
        lpSH As Long
    
    If Not FileExists(FileName) Then Exit Function
        
    lpSA.nLength = Len(lpSA)
    
    hFile = CreateFile(FileName, GENERIC_READ, 0&, lpSA, OPEN_EXISTING, 0&, 0&)
    If hFile = -1& Then Exit Function
        
    SL = pGetFileSize(hFile, lpSH)
    
    lpSizeHigh = lpSH
    GetFileSize = SL
    
    CloseHandle hFile

End Function


Public Function FileSizeDecimal(ByVal FileName As String) As Variant

    Dim decOut As Variant, _
        fSize1 As Long, _
        fSize2 As Long
    
    Dim mult
    
    mult = CDec(4294967296#)
    
    decOut = CDec(0)
    
    fSize1 = GetFileSize(FileName, fSize2)
    
    decOut = CDec(fSize1) + (mult * CDec(fSize2))
    FileSizeDecimal = decOut
    
End Function



Public Function GetTypeNameOf(ByVal FileName As String) As String

    Dim sfi As SHFILEINFO, _
        s As String
    
    Dim i As Long
    
    If FileExists(FileName) = False Then Exit Function
    
    SHGetFileInfo FileName, 0&, sfi, Len(sfi), SHGFI_TYPENAME
    
    s = sfi.szTypeName
    i = InStr(s, Chr(0))
    
    If (i - 1) Then
        s = Mid(s, 1, i - 1)
    Else
        s = ""
    End If
    
    GetTypeNameOf = Trim(s)
    
End Function

' Recursively makes a directory into existance.

' Creates every directory that does not exist, descending from root
' to the path argument passed, as needed.

Public Function CreatePath(ByVal Path As String)
    
    Dim j As Integer, i As Integer, s As String
    
    i = 4
    j = 1
    Do While j <> 0
        j = InStr(i, Path, "\")
        If j = 0 Then j = Len(Path)
        
        s = Mid(Path, 1, j)
        If (PathExists(s) = False) Then
            MkDir s
        End If
        If j >= Len(Path) Then Exit Do
        i = j + 1
    Loop

End Function

' Add a backslash if there is none, or remove double backslash

Public Function CleanPath(Path As String) As String
    Dim s As String, _
        t As String
    
    Dim i As Long, _
        j As Long
    
    s = Path
    
    If Len(Path) = 2& Then
        CleanPath = Path
        Exit Function
    End If

    If (Mid(s, 2, 1) = ":") And (Mid(s, 3, 2) = "\\") Then
        
        t = Mid(s, 1, 3)
        t = t + Mid(s, 5)
        s = t
    
    End If

    If Mid(s, Len(s) - 1) = "\\" Then
        t = Mid(s, 1, Len(s) - 2)
    ElseIf Mid(s, Len(s)) = "\" Then

        t = Mid(s, 1, Len(s) - 1)
    Else
        t = s
    End If

    j = 3
    i = InStr(j, t, "\\")
    
    Do While i <> 0
        
        s = Mid(t, 1, i)
        s = s + Mid(t, i + 2)
        
        t = s
        j = i + 1
        
        i = InStr(j, t, "\\")
    Loop

    CleanPath = t
    Path = t
    
End Function


' Get the absolute file name of a file

Public Function GetAbsolutePathname(ByVal FileName As String) As String

    Dim s As String * 261, t As String * 261
    Dim i As Integer
    
    GetFullPathName FileName, 260, s, t
    
    i = InStr(1, s, Chr(0))
    GetAbsolutePathname = Mid(s, 1, i - 1)
    
End Function


Public Function IsNT5() As Boolean
    Dim lpOS As OSVERSIONINFO
    
    lpOS.OSVersionInfoSize = Len(lpOS)
    
    GetVersion lpOS
    
    If (lpOS.PlatformId = VER_PLATFORM_WIN32_NT) And _
        (lpOS.MajorVersion >= 5&) Then
        
        IsNT5 = True
    End If

End Function

Public Function GetUnicode() As Boolean
    Dim lpOS As OSVERSIONINFO
    
    lpOS.OSVersionInfoSize = Len(lpOS)
    
    GetVersion lpOS
    
    If (lpOS.PlatformId = VER_PLATFORM_WIN32_NT) Then
        GetUnicode = True
    End If
    
    
    FILESYS_UNICODE_CHECKED = True
    FILESYS_UNICODE = GetUnicode

End Function

Public Sub SysTimeToVBTime(lpSysTime As SYSTEMTIME, lpVBTime As Date)
        
    With lpSysTime
        
        lpVBTime = DateSerial(.wYear, .wMonth, .wDay) + _
            TimeSerial(.wHour, .wMinute, .wSecond)

    End With

End Sub

Public Sub VBTimeToSysTime(lpVBTime As Date, lpSysTime As SYSTEMTIME)
        
    With lpSysTime
        
        .wMilliseconds = 0&
        .wSecond = Second(lpVBTime)
        .wMinute = Minute(lpVBTime)
        .wHour = Hour(lpVBTime)
        
        .wDay = Day(lpVBTime)
        .wMonth = Month(lpVBTime)
        .wYear = Year(lpVBTime)

    End With

End Sub

Public Sub FileTimeToVBTime(lpFileTime As FILETIME, lpVBTime As Date)
        
    Dim lpSysTime As SYSTEMTIME, _
        lpLocal As FILETIME
    
    FileTimeToLocalFileTime lpFileTime, lpLocal
    
    FileTimeToSystemTime lpLocal, lpSysTime
    SysTimeToVBTime lpSysTime, lpVBTime
    
End Sub

Public Sub VBTimeToFileTime(lpVBTime As Date, lpFileTime As FILETIME)
        
    Dim lpSysTime As SYSTEMTIME, _
        lpLocal As FILETIME
                
    VBTimeToSysTime lpVBTime, lpSysTime
    SystemTimeToFileTime lpSysTime, lpLocal
    
    LocalFileTimeToFileTime lpLocal, lpFileTime
        
End Sub

Public Function GetProcAddress(ByVal Proc As Long) As Long
    GetProcAddress = Proc
End Function


' Return all the files in a directory in a string array that match a search expression

Public Function EnumFiles(ByVal SearchExpression As String, Count As Long) As String()

    Dim wfind As WIN32_FIND_DATA
    
    Dim b() As String, _
        c As String, _
        z As Long, i As Long
        
    Dim l As Long
    
    Count = 0
    l = FindFirstFile(SearchExpression, wfind)
    If l = -1& Then Exit Function
    
    i = 1
    Do While i <> 0
        
        If (wfind.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
            ReDim Preserve b(z)
        
            c = StrConv(wfind.cFilename, vbUnicode)
            b(z) = Mid(c, 1, InStr(1, c, Chr(0)) - 1)
            z = z + 1
        End If
        
        i = FindNextFile(l, wfind)
        
    Loop
    FindClose l
    
    Count = z
    EnumFiles = b
    
End Function

' Return all the sub-directories in a directory in a string array that match a search expression

Public Function EnumSubFolders(ByVal SearchExpression As String, Count As Long) As String()

    Dim wfind As WIN32_FIND_DATA
    
    Dim b() As String, _
        c As String, _
        z As Long, i As Long
        
    Dim l As Long
    
    Count = 0
    
    l = FindFirstFile(SearchExpression, wfind)
    If l = -1& Then Exit Function
    
    i = 1
    Do While i <> 0
        
        If (wfind.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
        
            c = StrConv(wfind.cFilename, vbUnicode)
            c = Mid(c, 1, InStr(1, c, Chr(0)) - 1)
            
            If (c <> ".") And (c <> "..") Then
                ReDim Preserve b(z)
                b(z) = c
                z = z + 1
            End If
        End If
        
        i = FindNextFile(l, wfind)
        
    Loop
    FindClose l
    
    Count = z
    EnumSubFolders = b
    
End Function


' Returns a boolean indicating weather the file exists or not

Public Function FileExists(ByVal FileName As String) As Boolean
    Dim wf As WIN32_FIND_DATA
    
    Dim i As Long, _
        e As Long
    
    
    i = FindFirstFile(FileName, wf)

    If i <> -1& Then
        
        If (wf.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0& Then
            FileExists = True
        Else
            e = FindNextFile(i, wf)
            
            Do While e <> 0&
                If (wf.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0& Then
                    FileExists = True
                    Exit Do
                End If
                
                e = FindNextFile(i, wf)
            Loop
        
        End If
                        
        FindClose i
    End If

End Function

Public Function DeleteFile(ByVal FileName As String) As Boolean
    DeleteFile = CBool(pDeleteFile(FileName))
End Function

Public Function CopyFile(ByVal FileName As String, ByVal Destination As String, Optional ByVal FailIfExists As Boolean) As Boolean
    CopyFile = CBool(pCopyFile(FileName, Destination, FailIfExists))
End Function

Public Function MoveFile(ByVal FileName As String, ByVal Destination As String) As Boolean
    MoveFile = CBool(pMoveFile(FileName, Destination))
End Function

Public Function GetFNOnly(ByVal InName As String) As String
    Dim vOut As String, _
        i As Long
        
    i = InStrRev(InName, "\")
    If i = 0 Then
        GetFNOnly = InName
    Else
        GetFNOnly = Mid(InName, i + 1)
    End If
    
End Function

Public Function GetPathOnly(ByVal InName As String) As String
    Dim vOut As String, _
        i As Long
        
    i = InStrRev(InName, "\")
    If i = 0 Then
        vOut = String(1025, " ")
        GetCurrentDirectory 1024, vOut
        vOut = Trim(vOut)
        
        GetPathOnly = vOut
    Else
        GetPathOnly = Mid(InName, 1, i - 1)
    End If
    
End Function


' Returns a boolean indicating weather the folder exists or not

Public Function PathExists(ByVal Path As String) As Boolean

    Dim i As Long, _
        e As Long, _
        s As String
                
    Dim wf As WIN32_FIND_DATA, _
        df As Boolean
    
    s = Path
    i = Len(s)
    If Len(s) = 0 Then Exit Function
    
    If (Mid(s, i, 1) = "\") And (i <> 3) Then
        s = Mid(s, 1, i - 1)
    ElseIf (i = 3) And (Mid(s, i - 1, 2) = ":\") Then
        s = s + "*.*"
        df = True
    End If
    
    i = FindFirstFile(s, wf)

    If i <> -1& Then
        
        If (wf.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Or (df = True) Then
            PathExists = True
            
        Else
            e = FindNextFile(i, wf)
            
            Do While e <> 0&
            
                If (wf.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    PathExists = True
                    Exit Do
                End If
                
                e = FindNextFile(i, wf)
            Loop
        
        End If
        
        FindClose i
    End If

End Function


Public Function GetSpecialFolder(ByVal hwnd As Long, ByVal fSpcPath As SpecialFolderConstants) As String
    Dim lpStr As String, _
        i As Long
    
    lpStr = String(512, 0)
    SHGetSpecialFolderPath hwnd, lpStr, fSpcPath, False
    
    i = InStr(1, lpStr, Chr(0))
    If i = 0 Then
        GetSpecialFolder = lpStr
    Else
        GetSpecialFolder = Mid(lpStr, 1, i - 1)
    End If

End Function


