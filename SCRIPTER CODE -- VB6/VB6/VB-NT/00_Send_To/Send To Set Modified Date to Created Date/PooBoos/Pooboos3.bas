Attribute VB_Name = "Pooboos1100_3"
' Jamal 331 Pooboos 1100 Section III

' Contact as Csjh331@hotmail.com

Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DPtoLP Lib "gdi32" (ByVal HDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal HDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal HDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal HDC As Long, ByVal iMode As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpStringDest As String, ByVal lpStringSrc As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByVal lpszUrlSearchPattern As String, ByVal lpFirstCacheEntryInfo As Long, ByRef lpdwFirstCacheEntryInfoBufferSize As Long) As Long
Private Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, ByVal lpNextCacheEntryInfo As Long, ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, Buffer As Long, ByVal pbSize As Long, pbSizeNeeded As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Private Declare Function Escape Lib "gdi32" (ByVal HDC As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal HDC As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal HModule As Long, ByVal lpType As resType, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function StrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function MonitorFromPoint Lib "user32.dll" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromRect Lib "user32.dll" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal HDC As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal HDC As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Private Declare Function WSAStartupInfo Lib "WSOCK32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSADATA As WSADataInfo) As Long
Private Declare Function WSACleanup Lib "WSOCK32" () As Long
Private Declare Function WSAGetLastError Lib "WSOCK32" () As Long
Private Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Private Declare Function gethostname Lib "WSOCK32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32" (ByVal szHost As String) As Long
Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
Private Declare Function waveOutGetVolume Lib "Winmm.dll" (ByVal wDeviceID As Integer, dwVolume As Long) As Integer
Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal Handle As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function Netbios Lib "netapi32.dll" (pncb As NET_CONTROL_BLOCK) As Byte
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal HDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal HDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hwinsta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetProcessWindowStation Lib "user32" () As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal HDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal HDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function GetOpenFileNamePreview Lib "msvfw32.dll" (ByRef lpofn As OPENFILENAME) As Long
Private Declare Function GetSaveFileNamePreview Lib "msvfw32.dll" Alias "GetSaveFileNamePreviewA" (ByRef lpofn As OPENFILENAME) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal HDC As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal HDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)
Private Declare Sub FindCloseUrlCache Lib "wininet.dll" (ByVal hEnumHandle As Long)
Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Const FW_HEAVY = 900
Private Const ANSI_CHARSET = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0
Private Const TRANSPARENT = 1
Private Const RGN_COPY = 5
Private Const OPAQUE = 2
Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RESET = &H80000
Private Const MEM_TOP_DOWN = &H100000
Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_EXECUTE = &H10
Private Const PAGE_EXECUTE_READ = &H20
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const PAGE_GUARD = &H100
Private Const PAGE_NOACCESS = &H1
Private Const PAGE_NOCACHE = &H200
Private Const ENABLES_TRUE = 1
Private Const ENABLES_FALSE = 0
Private Const LPERRSTRING_ON = 1
Private Const LPERRSTRING_OFF = 0
Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Const SW_HIDE = 0
Private Const EM_SETSEL = &HB1
Private Const EM_REPLACESEL = &HC2
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const NEWFRAME = 1
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&
Private Const TMPF_FIXED_PITCH = &H1
Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4
Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
Private Const TRUETYPE_FONTTYPE = &H4
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Const DIFFERENCE = 11
Private Const LOAD_LIBRARY_AS_DATAFILE = &H2
Private Const MONITORINFOF_PRIMARY = &H1
Private Const MONITOR_DEFAULTTONEAREST = &H2
Private Const MONITOR_DEFAULTTONULL = &H0
Private Const MONITOR_DEFAULTTOPRIMARY = &H1
Private Const PRINTER_ENUM_LOCAL = &H2
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const SOCKET_ERROR As Long = -1
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const WN_SUCCESS = 0
Private Const WN_NET_ERROR = 2
Private Const WN_BAD_PASSWORD = 6
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const MAX_PATH = 260
Private Const HKEY_CURRENT_USER = &H80000001
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Const REG_FORCE_RESTORE As Long = 8&
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000
Private Const INTERNET_CONNECTION_CONFIGURED = &H40
Private Const INTERNET_CONNECTION_LAN = &H2
Private Const INTERNET_CONNECTION_MODEM = &H1
Private Const INTERNET_CONNECTION_OFFLINE = &H20
Private Const INTERNET_CONNECTION_PROXY = &H4
Private Const INTERNET_RAS_INSTALLED = &H10
Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const RASTERCAPS As Long = 38
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const AC_SRC_OVER = &H0
Private Const FLOODFILLBORDER = 0
Private Const FLOODFILLSURFACE = 1
Private Const crNewColor = &HFFFF80
Private Const OFN_EXPLORER = &H80000
Private Const ALTERNATE = 1
Private Const WINDING = 2
Private Const BLACKBRUSH = 4

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type INTERNET_CACHE_ENTRY_INFO
    dwStructSize As Long
    lpszSourceUrlName As Long
    lpszLocalFileName As Long
    CacheEntryType As Long
    dwUseCount As Long
    dwHitRate As Long
    dwSizeLow As Long
    dwSizeHigh As Long
    LastModifiedTime As FILETIME
    ExpireTime As FILETIME
    LastAccessTime As FILETIME
    LastSyncTime As FILETIME
    lpHeaderInfo As Long
    dwHeaderInfoSize As Long
    lpszFileExtension As Long
    dwReserved As Long
    dwExemptDelta As Long
End Type

Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges As LUID_AND_ATTRIBUTES
End Type

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Type PRINTER_INFO_2
   pServerName As String
   pPrinterName As String
   pShareName As String
   pPortName As String
   pDriverName As String
   pComment As String
   pLocation As String
   pDevMode As Long
   pSepFile As String
   pPrintProcessor As String
   pDatatype As String
   pParameters As String
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   JobsCount As Long
   AveragePPM As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type PAINTSTRUCT
        HDC As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type

Private Type Size
        cx As Long
        cy As Long
End Type

'Private Type POINTAPI
'        x As Long
'        y As Long
'End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Private Type STARTUPINFO
  cb As Long
  lpReserved As Long
  lpDesktop As Long
  lpTitle As Long
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Byte
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Type PRINTER_DEFAULTS
  pDatatype As String
  pDevMode As DEVMODE
  DesiredAccess As Long
End Type

Enum Printer_Status
   PRINTER_STATUS_READY = &H0
   PRINTER_STATUS_PAUSED = &H1
   PRINTER_STATUS_ERROR = &H2
   PRINTER_STATUS_PENDING_DELETION = &H4
   PRINTER_STATUS_PAPER_JAM = &H8
   PRINTER_STATUS_PAPER_OUT = &H10
   PRINTER_STATUS_MANUAL_FEED = &H20
   PRINTER_STATUS_PAPER_PROBLEM = &H40
   PRINTER_STATUS_OFFLINE = &H80
   PRINTER_STATUS_IO_ACTIVE = &H100
   PRINTER_STATUS_BUSY = &H200
   PRINTER_STATUS_PRINTING = &H400
   PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
   PRINTER_STATUS_NOT_AVAILABLE = &H1000
   PRINTER_STATUS_WAITING = &H2000
   PRINTER_STATUS_PROCESSING = &H4000
   PRINTER_STATUS_INITIALIZING = &H8000
   PRINTER_STATUS_WARMING_UP = &H10000
   PRINTER_STATUS_TONER_LOW = &H20000
   PRINTER_STATUS_NO_TONER = &H40000
   PRINTER_STATUS_PAGE_PUNT = &H80000
   PRINTER_STATUS_USER_INTERVENTION = &H100000
   PRINTER_STATUS_OUT_OF_MEMORY = &H200000
   PRINTER_STATUS_DOOR_OPEN = &H400000
   PRINTER_STATUS_SERVER_UNKNOWN = &H800000
   PRINTER_STATUS_POWER_SAVE = &H1000000
End Enum

Private Type NET_CONTROL_BLOCK
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte
   ncb_event      As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Type POINT
    x As Long
    y As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type

Private Type PRINTER_INFO_1
        flags As Long
        pDescription As String
        pName As String
        pComment As String
End Type

Private Type PORT_INFO_2
    pPortName As String
    pMonitorName As String
    pDescription As String
    fPortType As Long
    Reserved As Long
End Type

Private Type API_PORT_INFO_2
    pPortName As Long
    pMonitorName As Long
    pDescription As Long
    fPortType As Long
    Reserved As Long
End Type

Private Type MODULEENTRY32
  dwSize As Long
  th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  HModule As Long
  szModule As String * 256
  szExePath As String * 260
End Type

Private Type NAME_BUFFER
   Name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Private Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type

Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    Data As Long
    Options As IP_OPTION_INFORMATION
End Type

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Type WSADataInfo
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type

Private Type NEWTEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
   ntmFlags As Long
   ntmSizeEM As Long
   ntmCellHeight As Long
   ntmAveWidth As Long
End Type

Enum resType
    RT_FIRST = 1&
    RT_CURSOR = 1&
    RT_BITMAP = 2&
    RT_ICON = 3&
    RT_MENU = 4&
    RT_DIALOG = 5&
    RT_STRING = 6&
    RT_FONTDIR = 7&
    RT_FONT = 8&
    RT_ACCELERATOR = 9&
    RT_RCDATA = 10&
    RT_MESSAGETABLE = (11)
    RT_GROUP_CURSOR = (RT_CURSOR + DIFFERENCE)
    RT_GROUP_ICON = (RT_ICON + DIFFERENCE)
    RT_VERSION = (16)
    RT_LAST = (16)
End Enum

Dim mBrush As Long
Dim Ports(0 To 100) As PORT_INFO_2
Dim sClasses As String
Dim hMemoryDC As Long
Dim m_VirtualMem As Long, lLength As Long

Dim LP_RESULT_CALLBACK As Long

Dim IntMain As Integer

Dim duNUM(0 To 128) As Byte
Dim dvNUM(0 To 256) As Byte
Property Get Handle() As Long
    Handle = m_VirtualMem
End Property
Function InitFill(Picture As PictureBox, Color As Long)
    'Initializes the Flood Up
    mBrush = CreateSolidBrush(Color)
    SelectObject Picture.HDC, mBrush
    Picture.ScaleMode = vbPixels
End Function
Function FillToPictureBox(Picture As PictureBox)
    'Floods up the PictureBox
    Dim x As Integer, y As Integer
    ExtFloodFill Picture.HDC, x, y, GetPixel(Picture.HDC, x, y), FLOODFILLSURFACE
    
    'For example:

    'Private Sub Form_Load()
    'InitFill Picture1, vbBlue
    'End Sub

    'Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'FillToPictureBox Picture1
    'End Sub
End Function
Function CrashedErrors()
    FatalAppExit 0, "Error in module POOBOOS 1100 at 0x2659AFF. The program will now terminate."
End Function
Public Sub Allocate(lCount As Long)
    EndMemory
    m_VirtualMem = VirtualAlloc(ByVal 0&, lCount, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    VirtualLock m_VirtualMem, lCount
End Sub
Public Sub ReadFrom(hWritePointer As Long, lLength As Long)
    If IsBadWritePtr(hWritePointer, lLength) = 0 And IsBadReadPtr(Handle, lLength) = 0 Then
        CopyMemory hWritePointer, Handle, lLength
    End If
End Sub
Public Sub WriteTo(hReadPointer As Long, lLength As Long)
    If IsBadReadPtr(hReadPointer, lLength) = 0 And IsBadWritePtr(Handle, lLength) = 0 Then
        CopyMemory Handle, hReadPointer, lLength
    End If
End Sub
Public Function ExtractString(hStartPointer As Long, lMax As Long) As String
    Dim Length As Long
    If IsBadStringPtr(hStartPointer, lMax) = 0 Then
        ExtractString = Space(lMax)
        lstrcpy ExtractString, hStartPointer
        Length = lstrlen(hStartPointer)
        If Length >= 0 Then ExtractString = Left$(ExtractString, Length)
    End If
End Function
Public Sub EndMemory()
' Releases Memory
    If m_VirtualMem <> 0 Then
        VirtualUnlock m_VirtualMem, lLength
        VirtualFree m_VirtualMem, lLength, MEM_DECOMMIT
        VirtualFree m_VirtualMem, 0, MEM_RELEASE
        m_VirtualMem = 0
    End If
End Sub
Function AviFilePreviewDialog(ByVal hWnd As Long, duTitle As String, dvFilter As String, dwExtension As String)
    Dim OFN As OPENFILENAME, Ret As Long
    With OFN
        .lStructSize = Len(OFN)
        .hInstance = App.hInstance
        .hwndOwner = hWnd
        .lpstrTitle = duTitle
        .lpstrFilter = dvFilter + Chr$(0) + dwExtension + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*"
        .lpstrFile = String(255, 0)
        .nMaxFile = 255
        .flags = OFN_EXPLORER
    End With
    Ret = GetOpenFileNamePreview(OFN)
End Function
Function StarShapedForm(ByVal hWnd As Long, duSize As Long) As Long
    Dim PentaPoints(1 To 5) As POINTAPI, Cnt As Long
    NumCoords = 5
    PentaPoints(1).x = duSize / 2
    PentaPoints(1).y = 0
    PentaPoints(2).x = (duSize / 4) * 3
    PentaPoints(2).y = duSize
    PentaPoints(3).x = 0
    PentaPoints(3).y = duSize / 3
    PentaPoints(4).x = duSize
    PentaPoints(4).y = duSize / 3
    PentaPoints(5).x = duSize / 4
    PentaPoints(5).y = duSize
    hRgn = CreatePolygonRgn(PentaPoints(1), NumCoords, WINDING)
    LP_RESULT_CALLBACK = SetWindowRgn(hWnd, hRgn, True)
    
    ' For Example,

    ' Private Sub Form_Load()
    ' StarShapedForm Me.hWnd, Me.Height And Me.Width < = To use the size same as the Form.
    ' End Sub

    ' Private Sub Form_Load()
    ' StarShapedForm Me.hWnd, 250 < = Specified Size (at least 150)
    ' End Sub

End Function
Function IsFileAlreadyOpen(Filename As String) As Boolean
    Dim hFile As Long
    Dim lastErr As Long
    hFile = -1
    lastErr = 0
    hFile = lOpen(Filename, &H10)
    If hFile = -1 Then
        lastErr = Err.LastDllError
    Else
        lClose (hFile)
    End If
    sFileAlreadyOpen = (hFile = -1) And (lastErr = 32)
End Function
Public Sub BlendPicture(aPict As PictureBox, bPict As PictureBox)
    Dim BF As BLENDFUNCTION, lBF As Long
    aPict.AutoRedraw = True
    bPict.AutoRedraw = True
    
    aPict.ScaleMode = vbPixels
    bPict.ScaleMode = vbPixels
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = 128
        .AlphaFormat = 0
    End With
    RtlMoveMemory lBF, BF, 4
    AlphaBlend bPict.HDC, 0, 0, bPict.ScaleWidth, bPict.ScaleHeight, aPict.HDC, 0, 0, aPict.ScaleWidth, aPict.ScaleHeight, lBF
End Sub
Private Sub NullifyInputs(ByVal duTime As Long)
    DoEvents
    BlockInput True
    Sleep duTime
    BlockInput False
End Sub
Function StripNulls(OriginalStr As String) As String
    'Removes NullStrings.
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
Function FindFilesAPI(Path As String, SearchStr As String, FileCount As Integer, DirCount As Integer, aList As ListBox)

    Dim Filename As String
    Dim DirName As String
    Dim dirNames() As String
    Dim nDir As Integer
    Dim i As Integer
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(Path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        If (DirName <> ".") And (DirName <> "..") Then
            If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD)
        Loop
        Cont = FindClose(hSearch)
    End If
    hSearch = FindFirstFile(Path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            Filename = StripNulls(WFD.cFileName)
            If (Filename <> ".") And (Filename <> "..") Then
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                aList.AddItem Path & Filename
            End If
            Cont = FindNextFile(hSearch, WFD)
        Wend
        Cont = FindClose(hSearch)
    End If
    If nDir > 0 Then
        For i = 0 To nDir - 1
            FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount, aList)
        Next i
    End If
End Function
Function FindFilesProc(List As ListBox, aText As TextBox, bText As TextBox, cText As TextBox, dText As TextBox)
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer, NumDirs As Integer
    Screen.MousePointer = vbHourglass
    List1.Clear
    SearchPath = aText.Text
    FindStr = bText.Text
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs, List)
    cText.Text = NumFiles & " Files found in " & NumDirs + 1 & " Directories"
    dText.Text = "Total size of files found under " & SearchPath & " = " & Format(FileSize, "#,###,###,##0") & " Bytes"
    Screen.MousePointer = vbDefault
End Function
Function FindDesktop(ByVal lpszDesktop As Long, ByVal lParam As Long, ByVal Form As Form) As Boolean
    Dim Buffer As String
    Buffer = Space(lstrlen(lpszDesktop))
    CopyMemory Buffer, lpszDesktop, Len(Buffer)
    Form.Print Buffer
    FindDesktop = True
End Function
Public Function FindShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    strPath = String$(165, 0)
    lngRes = GetShortPathName(strFileName, strPath, 164)
    FindShortPath = Left$(strPath, lngRes)
End Function
Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim r As Long, Pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    Set CreateBitmapPicture = IPic
End Function
Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, r As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE

    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    r = DeleteDC(hDCMemory)
    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)
End Function
Function PrintScreenOntoForm(ByVal Form As Form)
    Set Form.Picture = hDCToPicture(GetDC(0), 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
End Function
Function GetMACAddress() As String
   Dim tmp As String
   Dim pASTAT As Long
   Dim NCB As NET_CONTROL_BLOCK
   Dim AST As ASTAT
   NCB.ncb_command = NCBRESET
   Call Netbios(NCB)
   NCB.ncb_callname = "*               "
   NCB.ncb_command = NCBASTAT

   NCB.ncb_lana_num = 0
   NCB.ncb_length = Len(AST)
   pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, NCB.ncb_length)
   If pASTAT = 0 Then
      Debug.Print "Memory Allocation Failed!"
      Exit Function
   End If
   NCB.ncb_buffer = pASTAT
   Call Netbios(NCB)
   tmp = Format$(Hex(AST.adapt.adapter_address(0)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(1)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(2)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(3)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(4)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(5)), "00")
   HeapFree GetProcessHeap(), 0, pASTAT
   GetMACAddress = tmp
End Function
Function FindNetworkAddress()
   MsgBox "Network adapter address: " + GetMACAddress(), vbInformation, "Network Address"
End Function
Function TrimStr(strName As String) As String
    Dim x As Integer
    x = InStr(strName, vbNullChar)
    If x > 0 Then TrimStr = Left(strName, x - 1) Else TrimStr = strName
End Function
Function LpstrToString(ByVal lngPointer As Long) As String
    Dim lngLength As Long
    lngLength = lstrlenW(lngPointer) * 2
    LpstrToString = String(lngLength, 0)
    CopyMem ByVal StrPtr(LpstrToString), ByVal lngPointer, lngLength
    LpstrToString = TrimStr(StrConv(LpstrToString, vbUnicode))
End Function
Public Function GetAvailablePorts(ServerName As String) As Long
    Dim Ret As Long
    Dim PortsStruct(0 To 100) As API_PORT_INFO_2
    Dim pcbNeeded As Long
    Dim pcReturned As Long
    Dim TempBuff As Long
    Dim i As Integer
    Ret = EnumPorts(ServerName, 2, TempBuff, 0, pcbNeeded, pcReturned)
    TempBuff = HeapAlloc(GetProcessHeap(), 0, pcbNeeded)
    Ret = EnumPorts(ServerName, 2, TempBuff, pcbNeeded, pcbNeeded, pcReturned)
    If Ret Then
        CopyMem PortsStruct(0), ByVal TempBuff, pcbNeeded
        For i = 0 To pcReturned - 1
            Ports(i).pDescription = LpstrToString(PortsStruct(i).pDescription)
            Ports(i).pPortName = LpstrToString(PortsStruct(i).pPortName)
            Ports(i).pMonitorName = LpstrToString(PortsStruct(i).pMonitorName)
            Ports(i).fPortType = PortsStruct(i).fPortType
        Next
    End If
    GetAvailablePorts = pcReturned
    If TempBuff Then HeapFree GetProcessHeap(), 0, TempBuff
End Function
Function FindAvailablePorts()
    Dim NumPorts As Long
    Dim i As Integer
    NumPorts = GetAvailablePorts("")
    For i = 0 To NumPorts - 1
        MsgBox "Port:" + " " + Ports(i).pPortName, vbInformation, "Available Ports For Use"
    Next
End Function
Function Ping(HostName As String)
    Dim hFile As Long, lpWSADATA As WSAData
    Dim hHostent As HOSTENT, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Call WSAStartup(&H101, lpWSADATA)
    If gethostbyname(HostName + String(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal gethostbyname(HostName + String(64 - Len(HostName), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        MsgBox "Unable to Create File Handle"
        Exit Function
    End If
    OptInfo.TTL = 255
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
        MsgBox "Timeout"
    End If
    If EchoReply.Status = 0 Then
        MsgBox "Reply from " + HostName + " (" + rIP + ") recieved after " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms"
    Else
        MsgBox "Failure ..."
    End If
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
End Function
Function EnumeratingConnectionsStates()
    Dim Ret As Long
    InternetGetConnectedState Ret, 0&
    If (Ret And INTERNET_CONNECTION_CONFIGURED) = INTERNET_CONNECTION_CONFIGURED Then MsgBox "Local system has a valid connection to the Internet, but it may or may not be currently connected.", vbInformation, "Connection State"
    If (Ret And INTERNET_CONNECTION_LAN) = INTERNET_CONNECTION_LAN Then MsgBox "Local system uses a local area network to connect to the Internet.", vbInformation, "Connection State"
    If (Ret And INTERNET_CONNECTION_MODEM) = INTERNET_CONNECTION_MODEM Then MsgBox "Local system uses a modem to connect to the Internet.", vbInformation, "Connection State"
    If (Ret And INTERNET_CONNECTION_OFFLINE) = INTERNET_CONNECTION_OFFLINE Then MsgBox "Local system is in offline mode.", vbInformation, "Connection State"
    If (Ret And INTERNET_CONNECTION_PROXY) = INTERNET_CONNECTION_PROXY Then MsgBox "Local system uses a proxy server to connect to the Internet.", vbInformation, "Connection State"
    If (Ret And INTERNET_RAS_INSTALLED) = INTERNET_RAS_INSTALLED Then MsgBox "Local system has RAS installed.", vbInformation, "Connection State"
    End
End Function
Function IsConnected() As Boolean
    If InternetGetConnectedState(0&, 0&) = 1 Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function
Function EnumeratingModules(aList As ListBox)
    
    ' Information about module's process in a list.
    
    Dim uProcess As MODULEENTRY32
    lProcessID = GetCurrentProcessId
    hSnapShot = CreateToolhelp32Snapshot(8, 0)
    uProcess.dwSize = Len(uProcess)
    n = Module32First(hSnapShot, uProcess)
    Do While n
        aList.AddItem Left(uProcess.szModule, InStr(uProcess.szModule, Chr(0)) - 1)
        n = Module32Next(hSnapShot, uProcess)
    Loop
End Function
Function EnumeratingProcesses(aList As ListBox)

    'Processes List

    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    Do While r
        aList.AddItem Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
        r = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot
End Function
Function EnablePrivilege(seName As String) As Boolean
    Dim p_lngRtn As Long
    Dim p_lngToken As Long
    Dim p_lngBufferLen As Long
    Dim p_typLUID As LUID
    Dim p_typTokenPriv As TOKEN_PRIVILEGES
    Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES
    p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
    If p_lngRtn = 0 Then
        Exit Function
    ElseIf Err.LastDllError <> 0 Then
        Exit Function
    End If
    p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)  'Used to look up privileges LUID.
    If p_lngRtn = 0 Then
        Exit Function
    End If
    
    p_typTokenPriv.PrivilegeCount = 1
    p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
    p_typTokenPriv.Privileges.pLuid = p_typLUID
    EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function
Public Function RestoreKey(ByVal sKeyName As String, ByVal sFilename As String, lPredefinedKey As Long) As Boolean
    If EnablePrivilege(SE_RESTORE_NAME) = False Then Exit Function
    Dim hKey As Long, lRetVal As Long
    Call RegOpenKeyEx(lPredefinedKey, sKeyName, 0&, KEY_ALL_ACCESS, hKey)
    Call RegRestoreKey(hKey, sFilename, REG_FORCE_RESTORE)
    RegCloseKey hKey
End Function
Public Function SaveKey(ByVal sKeyName As String, ByVal sFilename As String, lPredefinedKey As Long) As Boolean
    If EnablePrivilege(SE_BACKUP_NAME) = False Then Exit Function
    Dim hKey As Long, lRetVal As Long
    Call RegOpenKeyEx(lPredefinedKey, sKeyName, 0&, KEY_ALL_ACCESS, hKey)
    If Dir(sFilename) <> "" Then Kill sFilename
    Call RegSaveKey(hKey, sFilename, ByVal 0&)
    RegCloseKey hKey
End Function
Function SpecialFoldersAre()
    LP_RESULT_CALLBACK = MsgBox("Start menu folder is located at : " + GetSpecialfolder(CSIDL_STARTMENU), vbInformation, "Start Menu Folder")
    LP_RESULT_CALLBACK = MsgBox("Favorites folder is located at : " + GetSpecialfolder(CSIDL_FAVORITES), vbInformation, "Favorites Folder")
    LP_RESULT_CALLBACK = MsgBox("Programs folder is located at : " + GetSpecialfolder(CSIDL_PROGRAMS), vbInformation, "Programs Menu Folder")
    LP_RESULT_CALLBACK = MsgBox("Desktop folder is located at : " + GetSpecialfolder(CSIDL_DESKTOP), vbInformation, "Desktop Folder")
    End
End Function
Function GetSpecialfolder(ByVal CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = NOERROR Then
        Path$ = Space$(512)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function
Function LastModifiedToCurrent(duFile As String)
    Dim lngHandle As Long
    lngHandle = CreateFile(duFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If TouchFileTimes(lngHandle, ByVal 0&) = 0 Then
        'MsgBox "Error while changing file dates. Access Denied." + vbCrLf + duFile, vbCritical, "Modify Error"
    LastModifiedToCurrent = False
    Else
        'MsgBox "File date changed successfully.", vbInformation, "Modify Successful"
    LastModifiedToCurrent = True
    End If
    CloseHandle lngHandle
End Function
Sub IsLanguage(ByVal lpWord As Long)
    Dim lpBuffer As String
    lpBuffer = String(255, 0)
    VerLanguageName lpWord, lpBuffer, Len(lpBuffer)
    lpBuffer = Left$(lpBuffer, InStr(1, lpBuffer, Chr$(0)) - 1)
    MsgBox lpBuffer
End Sub
Function FindWaveVolume()
    Dim a, i As Long
    Dim tmp As String
    a = waveOutGetVolume(0, i)
    tmp = "&h" & Right(Hex$(i), 4)
    MsgBox "Wave Volume Is At: " & ((CSng(tmp) / 65536) * 100) & "  %", vbInformation, "Volume Is"
End Function
Function AddConnection(MyShareName As String, MyPWD As String, UseLetter As String) As Integer
    On Local Error GoTo AddConnection_Err
    AddConnection = WNetAddConnection(MyShareName, MyPWD, UseLetter)
AddConnection_End:
    Exit Function
AddConnection_Err:
    AddConnection = Err
    MsgBox error$
    Resume AddConnection_End
End Function
Function CancelConnection(DriveLetter As String, Force As Integer) As Integer
    On Local Error GoTo CancelConnection_Err
    CancelConnection = WNetCancelConnection(DriveLetter, Force)
CancelConnection_End:
    Exit Function
CancelConnection_Err:
    CancelConnection = Err
    MsgBox error$
    Resume CancelConnection_End
End Function
Public Function GetIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim HOST As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        MsgBox "Windows Sockets error " & str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    If lpHost = 0 Then
        GetIPAddress = ""
        MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    CopyMemoryIP HOST, lpHost, Len(HOST)
    CopyMemoryIP dwIPAddr, HOST.h_addr_list, 4
    ReDim tmpIPAddr(1 To HOST.h_length)
    CopyMemoryIP tmpIPAddr(1), dwIPAddr, HOST.h_length
    For i = 1 To HOST.h_length
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function
Public Function GetIPHostName() As String
    Dim sHostName As String * 256
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup
End Function
Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function
Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
Public Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
End Sub
Public Function SocketsInitialize() As Boolean

    'For Example,
    'Private Sub Command1_Click()
    'MsgBox "Sockets Initialized: " & SocketsInitialize
    'End Sub

    Dim WSAD As WSAData
    Dim sLoByte As String
    Dim sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        MsgBox "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If
    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
    End If
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        MsgBox "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
        SocketsInitialize = False
        Exit Function
    End If
    SocketsInitialize = True
End Function
Public Sub FindIPAddress()
    MsgBox "Your IP-Address: " + GetIPAddress, vbInformation, "IP Address"
    End
End Sub
Function EnumeratingDesktops()
    MsgBox "Available Desktops:   " & LP_RESULT_CALLBACK, vbInformation, " "
    LP_RESULT_CALLBACK = EnumDesktops(GetProcessWindowStation, AddressOf FindDesktop, 0)
End Function
Function EnumeratingPrinters()
    Dim longbuffer() As Long
    Dim printinfo() As PRINTER_INFO_1
    Dim numbytes As Long
    Dim numneeded As Long
    Dim numprinters As Long
    Dim C As Integer, Retval As Long
    numbytes = 3076
    ReDim longbuffer(0 To numbytes / 4) As Long
    Retval = EnumPrinters(PRINTER_ENUM_LOCAL, "", 1, longbuffer(0), numbytes, numneeded, numprinters)
    If Retval = 0 Then
        numbytes = numneeded
        ReDim longbuffer(0 To numbytes / 4) As Long
        Retval = EnumPrinters(PRINTER_ENUM_LOCAL, "", 1, longbuffer(0), numbytes, numneeded, numprinters)
        If Retval = 0 Then
            MsgBox "Could not successfully enumerate the printes.", vbCritical, "Error"
        End
    End If
    End If

    If numprinters <> 0 Then ReDim printinfo(0 To numprinters - 1) As PRINTER_INFO_1
    For C = 0 To numprinters - 1
        printinfo(C).flags = longbuffer(4 * C)
        printinfo(C).pDescription = Space(lstrlen(longbuffer(4 * C + 1)))
        Retval = lstrcpy(printinfo(C).pDescription, longbuffer(4 * C + 1))
        printinfo(C).pName = Space(lstrlen(longbuffer(4 * C + 2)))
        Retval = lstrcpy(printinfo(C).pName, longbuffer(4 * C + 2))
        printinfo(C).pComment = Space(lstrlen(longbuffer(4 * C + 3)))
        Retval = lstrcpy(printinfo(C).pComment, longbuffer(4 * C + 3))
    Next C
    
    For C = 0 To numprinters - 1
        MsgBox "Name of Printer" & C & 1 & " is: " + printinfo(C).pName, vbInformation, "Printers"
    Next C
End Function
Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim sSave As String
    sSave = Space$(GetWindowTextLength(hWnd) + 1)
    GetWindowText hWnd, sSave, Len(sSave)
    sSave = Left$(sSave, Len(sSave) - 1)
    If sSave <> "" Then MsgBox sSave, vbInformation, "Childs"
    EnumChildProc = 1
End Function
Sub EnumeratingChilds()
    ' Get all active objects...
    
    On Error Resume Next
    EnumChildWindows GetDesktopWindow, AddressOf EnumChildProc, ByVal 0&
End Sub
Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, ByVal dwData As Long, ByVal hWnd As Long) As Long
    Dim MI As MONITORINFO, r As RECT
    MsgBox "Moitor handle: " + CStr(hMonitor), vbInformation, "Monitors"
    MI.cbSize = Len(MI)
    GetMonitorInfo hMonitor, MI
    MsgBox "Monitor Width/Height: " + CStr(MI.rcMonitor.Right - MI.rcMonitor.Left) + "x" + CStr(MI.rcMonitor.Bottom - MI.rcMonitor.Top), vbInformation, "Montiors"
    MsgBox "Primary monitor: " + CStr(CBool(MI.dwFlags = MONITORINFOF_PRIMARY)), vbInformation, "Monitors"
    If MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST) = hMonitor Then
        MsgBox " The Form is located on this monitor", vbInformation, "Monitors"
    End If
    If MonitorFromPoint(0, 0, MONITOR_DEFAULTTONEAREST) = hMonitor Then
        MsgBox "The point (0, 0) lies within the range of this monitor...", vbInformation, "Monitors"
    End If
    GetWindowRect hWnd, r
    If MonitorFromRect(r, MONITOR_DEFAULTTONEAREST) = hMonitor Then
        MsgBox "The rectangle of Form lies within this monitor", vbInformation, "Monitors"
    End If
    MsgBox "Metrics Complete", vbInformation, "Monitors"
    MonitorEnumProc = 1
End Function
Function EnumeratingMonitors()

    'Monitor Metrics
    
    EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
End Function
Function EnumThreadWndProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, Ret)
    sClasses = sClasses + sText + vbCrLf
    EnumThreadWndProc = 1
End Function
Function EnumeratingThreads(ByVal hWnd As Long)
    Dim ThreadID As Long, ProcessID As Long
    ThreadID = GetWindowThreadProcessId(hWnd, ProcessID)
    EnumThreadWindows ThreadID, AddressOf EnumThreadWndProc, 0
    MsgBox sClasses, vbInformation, "Threads"
End Function
Function EnumResNameProc(ByVal HModule As Long, ByVal lpszType As resType, ByVal lpszName As Long, ByVal lParam As Long) As Long
    Dim Num As String, IsNum As Boolean
    If (lpszName > &HFFFF&) Or (lpszName < 0) Then
        Num = PtrToVBString(lpszName)
        IsNum = False
    Else
        Num = CStr(lpszName)
        IsNum = True
    End If
    If IsNum Then
        MsgBox Num & " "
    Else
        MsgBox """" & Num & """ "
    End If
    EnumResNameProc = 1
End Function
Function PtrToVBString(ByVal lpszBuffer As Long) As String
    Dim Buffer As String, LenBuffer As Long
    LenBuffer = StrLen(lpszBuffer)
    Buffer = String(LenBuffer + 1, 0)
    StrCpy Buffer, lpszBuffer
    PtrToVBString = Left(Buffer, LenBuffer)
End Function
Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, lParam As Long) As Long
   Dim FaceName As String
   FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
   MsgBox Left$(FaceName, InStr(FaceName, vbNullChar) - 1), vbInformation, "Font Families"
   EnumFontFamProc = 1
End Function
Function EnumeratingFonts(ByVal HDC As Long) As Long
   Dim LF As LOGFONT
   EnumFontFamiliesEx HDC, LF, AddressOf EnumFontFamProc, ByVal 0&, 0
End Function
Function PrintPictureBox(ByVal dvPic As PictureBox)
    
    'Prints off the Picture Box
    
    dvPic.ScaleMode = vbPixels
    Printer.ScaleMode = vbPixels
    Printer.Print ""
    
    hMemoryDC = CreateCompatibleDC(dvPic.HDC)
    hOldBitMap = SelectObject(hMemoryDC, dvPic.Picture)
    StretchBlt Printer.HDC, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight, hMemoryDC, 0, 0, dvPic.ScaleWidth, dvPic.ScaleHeight, vbSrcCopy
    hOldBitMap = SelectObject(hMemoryDC, hOldBitMap)
    DeleteDC hMemoryDC
    
    Escape Printer.HDC, NEWFRAME, 0, 0&, 0&
    Printer.EndDoc
    
End Function
Function PrintData(Name As String, Data As String)
    If LenB(Data) > 0 Then
        Debug.Print Name + ": " + Data
    End If
End Function
Function Redirect(cmdLine As String, objTarget As Object)

  'Gives output parameters of (Dos) programs.

  Dim i%, t$
  Dim pa As SECURITY_ATTRIBUTES
  Dim pra As SECURITY_ATTRIBUTES
  Dim tra As SECURITY_ATTRIBUTES
  Dim pi As PROCESS_INFORMATION
  Dim sui As STARTUPINFO
  Dim hRead As Long
  Dim hWrite As Long
  Dim bRead As Long
  Dim lpBuffer(1024) As Byte
  pa.nLength = Len(pa)
  pa.lpSecurityDescriptor = 0
  pa.bInheritHandle = True
  
  pra.nLength = Len(pra)
  tra.nLength = Len(tra)

  If CreatePipe(hRead, hWrite, pa, 0) <> 0 Then
    sui.cb = Len(sui)
    GetStartupInfo sui
    sui.hStdOutput = hWrite
    sui.hStdError = hWrite
    sui.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    sui.wShowWindow = SW_HIDE
    If CreateProcess(vbNullString, cmdLine, pra, tra, True, 0, Null, vbNullString, sui, pi) <> 0 Then
      SetWindowText objTarget.hWnd, ""
      Do
        Erase lpBuffer()
        If ReadFile(hRead, lpBuffer(0), 1023, bRead, ByVal 0&) Then
          SendMessage objTarget.hWnd, EM_SETSEL, -1, 0
          SendMessage objTarget.hWnd, EM_REPLACESEL, False, lpBuffer(0)
          DoEvents
        Else
          CloseHandle pi.hThread
          CloseHandle pi.hProcess
          Exit Do
        End If
        CloseHandle hWrite
      Loop
      CloseHandle hRead
    End If
  End If
End Function
Function LimitCursorMovementDx(AttnScrX As Long, AttnScrY As Long)
    
    'Parameters should be numbers only.
    'For Instance LimitCursorMovementDx (320, 200)
    
    On Error Resume Next
        
    Dim dx As RECT
    Dim dy As POINT
    Dim lResult As Long
    
    With dx   'Right/Bottom side of Rectangle
         .Right = AttnScrX
         .Bottom = AttnScrY
    End With
    
    With dy   'Left/Top side of Rectangle
        .x = dx.Left
        .y = dx.Top
    End With
    
    If dx.Bottom <= 0 Then
        MsgBox "Cursor Limit cannot be 0 pixels in the Y axis.", vbCritical, "Error"
        Exit Function
        Resume Next
    End If

    If dx.Right <= 0 Then
        MsgBox "Cursor Limit cannot be 0 pixels in the X axis.", vbCritical, "Error"
        Exit Function
        Resume Next
    End If
                
    lResult = ClipCursor(dx)
End Function
Function EndClipping()
    On Error Resume Next
    ClipCursor ByVal 0&  'Release Cursor Clip.
End Function
Function Suspend(ByVal Frm As Form)

    'Nullifies and ignores input of Form.
    
    Dim Enables  As Long
    Dim Retval   As Long

    EndMemory   'Memory Release
    
    Enables = IsWindowEnabled(Frm.hWnd)
    Retval = EnableWindow(Frm.hWnd, ENABLES_FALSE)
    
    Debug.Print "Application Suspended"
End Function
Function EndSuspend(ByVal Frm As Form)

    'Stops nullification.
    
    Dim Enables As Long
    Dim Retval As Long
    
    EndMemory   'Memory Release
    
    Enables = IsWindowEnabled(Frm.hWnd)
    Retval = EnableWindow(Frm.hWnd, ENABLES_TRUE)
    
    Debug.Print "Application no longer halted"
End Function
Function DelTreeCacheDx(aList As ListBox)
    Dim ICEI As INTERNET_CACHE_ENTRY_INFO, Ret As Long
    Dim hEntry As Long, Msg As VbMsgBoxResult
    FindFirstUrlCacheEntry vbNullString, ByVal 0&, Ret
    If Ret > 0 Then
        Allocate Ret
        hEntry = FindFirstUrlCacheEntry(vbNullString, Handle, Ret)
        ReadFrom VarPtr(ICEI), LenB(ICEI)
        If ICEI.lpszSourceUrlName <> 0 Then aList.AddItem ExtractString(ICEI.lpszSourceUrlName, Ret)
    End If
    Do While hEntry <> 0
        Ret = 0
        FindNextUrlCacheEntry hEntry, ByVal 0&, Ret
        If Ret > 0 Then
            Allocate Ret
            FindNextUrlCacheEntry hEntry, Handle, Ret
            ReadFrom VarPtr(ICEI), LenB(ICEI)
            If ICEI.lpszSourceUrlName <> 0 Then aList.AddItem ExtractString(ICEI.lpszSourceUrlName, Ret)
        Else
            Exit Do
        End If
    Loop
    FindCloseUrlCache hEntry
    Msg = MsgBox("Do you wish to delete the Internet Explorer cache?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If Msg = vbYes Then
        For Ret = 0 To aList.ListCount - 1
            DeleteUrlCacheEntry aList.List(Ret)
        Next Ret
        MsgBox "Cache deleted..."
    End If
    
    EndMemory
    
End Function
Function DrawTextOnForm(ByVal hWnd As Long, duText As String, dvFont As String)

    'Should compile project first and then run the EXE file.

    'For Eample,
    'Private Sub Form_Load()
    'Refresh
    'DrawTextOnForm Me.hWnd, "Bird Fiharitz", "Comic Sans MS"
    'End Sub
    
Rem  Should have bform's backcolor besides the colors that are the same as the system colors. (Gray)
    
    Dim PAINTSTRUCT As PAINTSTRUCT
    Dim REC As RECT
    Dim sizeText As Size
    Dim ptText As POINTAPI
    Dim hFont As Long
    Dim a As Double, d As Double, r As Double
    Dim GT As Long
    Dim LINNE As Long
    Dim HDC As Long
    
    EndMemory
    
    HDC = BeginPaint(hWnd, PAINTSTRUCT)

    GetClientRect hWnd, REC
    DPtoLP HDC, ptText, 2

    hFont = CreateFont((REC.Bottom - REC.Top) / 5, (REC.Right - REC.Left) / 26, 0, 0, FW_HEAVY, False, False, False, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, dvFont)

    SelectObject HDC, hFont

    GT = GetTextExtentPoint32(HDC, duText, (Len(duText)), sizeText)

    ptText.x = (REC.Left + REC.Right - sizeText.cx) / 2
    ptText.y = (REC.Top + REC.Bottom - sizeText.cy) / 2
        
    SetBkMode HDC, TRANSPARENT
    BeginPath HDC
    TextOut HDC, ptText.x, ptText.y, duText, (Len(duText))
    EndPath HDC
    SelectClipPath HDC, RGN_COPY

    d = Sqr((sizeText.cx * sizeText.cx + sizeText.cy * sizeText.cy))

    For r = 0 To r <= 90

        a = r / 180 * 3.14159265359

        MoveToEx HDC, ptText.x, ptText.y, ptText
            
        LINNE = LineTo(HDC, ptText.x + (d * Cos(a)), ptText.y + (d * Sin(a)))
    Next
        EndPaint hWnd, PAINTSTRUCT
End Function
