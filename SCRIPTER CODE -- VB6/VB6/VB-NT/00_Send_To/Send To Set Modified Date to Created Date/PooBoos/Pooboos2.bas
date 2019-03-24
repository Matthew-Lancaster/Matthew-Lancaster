Attribute VB_Name = "Pooboos_1100_2"
'Jamal 331 Pooboos_1100 Section II (Continuation of Jamal 331 Poobooos 1100 Modules) Functions for Forms

'Contact me at Csjh331@hotmail.com or at Csjh218@yahoo.com

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function AnimateWindow Lib "user32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean
Private Declare Function SearchTreeForFile Lib "imagehlp" (ByVal RootPath As String, ByVal InputPathName As String, ByVal OutputPathBuffer As String) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal HDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As BOOL
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal strTempFilePath As String, ByVal strFilePrefix As String, ByVal lngUniqueValue As Long, ByVal ReturnBuffer As String) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal lngAccessType As Long, ByVal blnInheritHandle As BOOL, ByVal lngProcessID As Long) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptions As IP_OPTION_INFORMATION, ReplyBuffer As icmp_echo_reply, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function mciSendString Lib "Winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal uReserved As Integer) As Integer
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" (ByVal lprasdialextensions As Long, ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByVal dword As Long, lpvoid As Any, ByRef lphrasconn As Long) As Long
Private Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal Reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long
Private Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByRef lpbool As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32.DLL" (ByVal HDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32.DLL" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32.DLL" (ByVal HDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32.DLL" (ByVal HDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32.DLL" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32.DLL" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32.DLL" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32.DLL" (ByVal HDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function SelectPalette Lib "GDI32.DLL" (ByVal HDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32.DLL" (ByVal HDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal HDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As String, ByVal ByteLen As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal lngMilliseconds As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MemoryStatus)

Global MouseDown As Boolean
Global MouseOver As Boolean
Global S(52) As String
Global pi As Long
Global NumLinesOnPageToPrint As Integer
Global FirstPageNum As Integer
Global NextPageNum As Integer
Global LineNum As Integer
Global CheckThisLineNum As Integer
Global NumLines As Integer
Global TotalPageCount As Integer

Global Const EW_REBOOTSYSTEM = &H43
Global Const EW_RESTARTWINDOWS = &H42
Global Const EW_EXITWINDOWS = 0

Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Const IMAGE_ENHMETAFILE = 3
Private Const CF_BITMAP = 2
Private Const SHFD_CAPACITY_DEFAULT = 0
Private Const SHFD_CAPACITY_360 = 3
Private Const SHFD_CAPACITY_720 = 5
Private Const SHFD_FORMAT_QUICK = 0
Private Const SHFD_FORMAT_FULL = 1
Private Const SHFD_FORMAT_SYSONLY = 2
Private Const RAS_MAXENTRYNAME = 256
Private Const RAS_MAXDEVICETYPE = 16
Private Const RAS_MAXDEVICENAME = 128
Private Const RAS_RASCONNSIZE = 412
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_MENU = 4
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_WINDOWTEXT = 8
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_2NDACTIVECAPTION = 27
Private Const COLOR_2NDINACTIVECAPTION = 28
Private Const AW_HOR_POSITIVE = &H1
Private Const AW_HOR_NEGATIVE = &H2
Private Const AW_VER_POSITIVE = &H4
Private Const AW_VER_NEGATIVE = &H8
Private Const AW_CENTER = &H10
Private Const AW_HIDE = &H10000
Private Const AW_ACTIVATE = &H20000
Private Const AW_SLIDE = &H40000
Private Const AW_BLEND = &H80000
Private Const MAX_PATH = 260
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const PING_TIMEOUT = 200
Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYSSTATUS_LEN = 256
Private Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Private Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Private Const SOCKET_ERROR = -1
Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)
Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)
Private Const SPI_SCREENSAVERRUNNING = 97
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSEEVENTF_MOVE = &H1
Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Const STATUS_WAIT_0 = &H0
Private Const STATUS_ABANDONED_WAIT_0 = &H80
Private Const STATUS_USER_APC = &HC0
Private Const STATUS_TIMEOUT = &H102
Private Const STATUS_PENDING = &H103
Private Const STATUS_SEGMENT_NOTIFICATION = &H40000005
Private Const STATUS_GUARD_PAGE_VIOLATION = &H80000001
Private Const STATUS_DATATYPE_MISALIGNMENT = &H80000002
Private Const STATUS_BREAKPOINT = &H80000003
Private Const STATUS_SINGLE_STEP = &H80000004
Private Const STATUS_ACCESS_VIOLATION = &HC0000005
Private Const STATUS_IN_PAGE_ERROR = &HC0000006
Private Const STATUS_INVALID_HANDLE = &HC0000008
Private Const STATUS_NO_MEMORY = &HC0000017
Private Const STATUS_ILLEGAL_INSTRUCTION = &HC000001D
Private Const STATUS_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const STATUS_INVALID_DISPOSITION = &HC0000026
Private Const STATUS_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const STATUS_FLOAT_DENORMAL_OPERAND = &HC000008D
Private Const STATUS_FLOAT_DIVIDE_BY_ZERO = &HC000008E
Private Const STATUS_FLOAT_INEXACT_RESULT = &HC000008F
Private Const STATUS_FLOAT_INVALID_OPERATION = &HC0000090
Private Const STATUS_FLOAT_OVERFLOW = &HC0000091
Private Const STATUS_FLOAT_STACK_CHECK = &HC0000092
Private Const STATUS_FLOAT_UNDERFLOW = &HC0000093
Private Const STATUS_INTEGER_DIVIDE_BY_ZERO = &HC0000094
Private Const STATUS_INTEGER_OVERFLOW = &HC0000095
Private Const STATUS_PRIVILEGED_INSTRUCTION = &HC0000096
Private Const STATUS_STACK_OVERFLOW = &HC00000FD
Private Const STATUS_CONTROL_C_EXIT = &HC000013A
Private Const STATUS_FLOAT_MULTIPLE_FAULTS = &HC00002B4
Private Const STATUS_FLOAT_MULTIPLE_TRAPS = &HC00002B5
Private Const STATUS_REG_NAT_CONSUMPTION = &HC00002C9
Private Const DBG_TERMINATE_THREAD = &H40010003
Private Const DBG_TERMINATE_PROCESS = &H40010004
Private Const DBG_CONTROL_C = &H40010005
Private Const DBG_CONTROL_BREAK = &H40010008
Private Const DBG_EXCEPTION_NOT_HANDLED = &H80010001
Private Const DBG_CONTINUE = &H10002
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const LP_HT_CAPTION = 2
Private Const RAS95_MaxEntryName = 256
Private Const RAS_MaxPhoneNumber = 128
Private Const RAS_MaxCallbackNumber = RAS_MaxPhoneNumber
Private Const UNLEN = 256
Private Const PWLEN = 256
Private Const DNLEN = 12

Dim sConnType As String * 255

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Type RASDIALPARAMS
   dwSize As Long
   szEntryName(RAS95_MaxEntryName) As Byte
   szPhoneNumber(RAS_MaxPhoneNumber) As Byte
   szCallbackNumber(RAS_MaxCallbackNumber) As Byte
   szUserName(UNLEN) As Byte
   szPassword(PWLEN) As Byte
   szDomain(DNLEN) As Byte
End Type

Type RASENTRYNAME95
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type IP_OPTION_INFORMATION
    TTL             As Byte
    Tos             As Byte
    flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LARGE_INTEGER
    ullAvailPhys As LARGE_INTEGER
    ullTotalPageFile As LARGE_INTEGER
    ullAvailPageFile As LARGE_INTEGER
    ullTotalVirtual As LARGE_INTEGER
    ullAvailVirtual As LARGE_INTEGER
    ullAvailExtendedVirtual As LARGE_INTEGER
End Type

Type icmp_echo_reply
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As IP_OPTION_INFORMATION
    Data            As String * 250
End Type

Type tagWSAData
    wVersion            As Integer
    wHighVersion        As Integer
    szDescription       As String * WSADESCRIPTION_LEN_1
    szSystemStatus      As String * WSASYSSTATUS_LEN_1
    iMaxSockets         As Integer
    iMaxUdpDg           As Integer
    lpVendorInfo        As String * 200
End Type

Type PicBmp
  Size As Long
  Type As Long
  hBmp As Long
  hPal As Long
  Reserved As Long
End Type

Type PALETTEENTRY
  peRed As Byte
  peGreen As Byte
  peBlue As Byte
  peFlags As Byte
End Type

Type LOGPALETTE
  palVersion As Integer
  palNumEntries As Integer
  palPalEntry(255) As PALETTEENTRY
End Type

Type MemoryStatus
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Enum BOOL
   FALSE_ = 0
   TRUE_ = 1
End Enum

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type Mouse
    x As Long
    y As Long
End Type

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

'Don't have to use these below. But, don't delete them, they are used by some functions.

Type LP_RAM     'for inputing parmeters
    LP_RAM As Long
End Type

Type LP_RESULT  'for output of functions
    LP_RESULT As Long
End Type

Type LP_LONG
    LP_LONG As Long
End Type

Type LP_UNIT
    LP_UNIT As Long
End Type

Type LP_HWND
    LP_HWND As Long
End Type

Type LP_HHOOK
    LP_HHOOK As Long
End Type

Type LP_WORD
    LP_WORD As Long
End Type

Type LP_CHAR
    LP_CHAR As Byte
End Type

Type LP_BYTE
    LP_BYTE As Byte
End Type

Type LP_UCHAR
    LP_UCHAR As String
End Type

Type LP_DWORD
    LP_DWORD As String
End Type

Type LP_UINT
    LP_UINT As Integer
End Type

Type LP_INT
    LP_INT As Integer
End Type

Type LP_ITEM
    LP_ITEM As Object
End Type

Type LP_NUM
    LP_NUM As Single
End Type

Type LP_UARE
    LP_UARE As Single
End Type

Type LP_DNUM
    LP_DNUM As Double
End Type

Type LP_VARE
    LP_VARE As Double
End Type

Type LP_WARE
    LP_WARE As Variant
End Type

Type LP_BOOL
    LP_BOOL As Boolean
End Type

Public Function ProgressBar(ByRef ThePictureBox As PictureBox, ByVal Min As Long, ByVal Max As Long, ByVal Value As Long, Optional ByVal ShowProgressCaption As Boolean = False, Optional ByVal ForeColor As Long = 16777215, Optional ByVal BackColor As Long = 16711680, Optional ByVal FillColor As Long = vbButtonFace, Optional ByVal Alignment As AlignmentConstants = vbCenter, Optional ByVal ByPassChecks As Boolean = False)
On Error Resume Next
  
  Dim TheCaption As String
  Dim RangeDiff As Long
  
  If ThePictureBox Is Nothing Then Exit Function
  If ByPassChecks = False Then
  ThePictureBox.AutoRedraw = True
    
    ThePictureBox.AutoSize = False
    Set ThePictureBox.Picture = Nothing
    ThePictureBox.Visible = True
    
  End If
  
  RangeDiff = Max - Min
  If RangeDiff = 0 Then
    TheCaption = "0.0%"
  Else
    TheCaption = Format((Value - Min) / RangeDiff, "0.0%")
  End If
  
  If RangeDiff = 0 Then
    ThePictureBox.Line (0, 0)-(0, ThePictureBox.ScaleHeight), BackColor, BF
    ThePictureBox.Line (0, 0)-(ThePictureBox.ScaleWidth, ThePictureBox.ScaleHeight), FillColor, BF
  Else
    ThePictureBox.Line (0, 0)-((((Value - Min) / RangeDiff) * ThePictureBox.ScaleWidth), ThePictureBox.ScaleHeight), BackColor, BF
    ThePictureBox.Line ((((Value - Min) / RangeDiff) * ThePictureBox.ScaleWidth), 0)-(ThePictureBox.ScaleWidth, ThePictureBox.ScaleHeight), FillColor, BF
  End If
  
  If ShowProgressCaption = False Then
    ThePictureBox.Refresh
    Exit Function
  End If

  If Alignment = vbCenter Then
    ThePictureBox.CurrentX = (ThePictureBox.ScaleWidth / 2 - ThePictureBox.TextWidth(TheCaption) / 2)
  ElseIf Alignment = vbLeftJustify Then
    ThePictureBox.CurrentX = 1
  ElseIf Alignment = vbRightJustify Then
    ThePictureBox.CurrentX = (ThePictureBox.ScaleWidth - ThePictureBox.TextWidth(TheCaption)) - 1
  End If
  ThePictureBox.CurrentY = (ThePictureBox.ScaleHeight - ThePictureBox.TextHeight(TheCaption)) / 2
  
  ThePictureBox.ForeColor = ForeColor
  ThePictureBox.Print TheCaption
  ThePictureBox.Refresh
  
End Function
Public Function LoWord(ByVal dwValue As Long) As Integer
    LoWord = Val("&H" & Right("0000" & Hex(dwValue), 4))
End Function
Public Function HiWord(ByVal dwValue As Long) As Integer
    HiWord = Val("&H" & Left(Right("00000000" & Hex(dwValue), 8), 4))
End Function
Public Function LoByte(ByVal wValue As Integer) As Byte
    LoByte = Val("&H" & Right("00" & Hex(wValue), 2))
End Function
Public Function HiByte(ByVal wValue As Integer) As Byte
    HiByte = Val("&H" & Left(Right("0000" & Hex(wValue), 4), 2))
End Function
Public Function MakeWord(ByVal bHigh As Byte, ByVal bLow As Byte) As Integer
    MakeWord = Val("&H" & Right("00" & Hex(bHigh), 2) & Right("00" & Hex(bLow), 2))
End Function
Public Function MakeLong(ByVal wHigh As Integer, ByVal wLow As Integer) As Long
    MakeLong = Val("&H" & Right("0000" & Hex(wHigh), 4) & Right("0000" & Hex(wLow), 4))
End Function
Public Sub FileSave(Text As String, FilePath As String)
    Dim Directory As String
              Directory$ = FilePath
       Open Directory$ For Output As #1
           Print #1, Text
       Close #1
    Exit Sub
End Sub
Function FileOpen(FilePath As String)
    Dim Directory As String
    Directory$ = FilePath
    Dim MyString As String
       Open Directory$ For Input As #1
       While Not EOF(1)
           Input #1, FileOpen
           Wend
           Close #1
    Exit Function
End Function
Public Sub ListSave(List As ListBox, FilePath As String)
    Dim i As Integer
    Dim Directory As String
              Directory$ = FilePath
       Open Directory$ For Output As #1
       For i = 0 To List.ListCount - 1
           Print #1, List.List(i)
       Next i
       Close #1
    Exit Sub
End Sub
Public Sub ListOpen(List As ListBox, FilePath As String)
    Directory$ = FilePath
    Dim MyString As String
       Open Directory$ For Input As #1
       While Not EOF(1)
           Input #1, MyString$
           DoEvents
               List.AddItem MyString$
           Wend
           Close #1
    Exit Sub
End Sub
Public Sub MakeDir(DirPath As String)
    MkDir DirPath$
    Exit Sub
End Sub
Public Sub DeleteDir(DirPath As String)
    RmDir DirPath$
    Exit Sub
End Sub
Public Sub DelFilesInDir(DirPath As String, DelDir As Boolean)
    Kill DirPath$ & "*.*"
    If DelDir = True Then
    RmDir DirPath$
    End If
Exit Sub
End Sub
Public Sub MoveFile(StartPath As String, EndPath As String)
    FileCopy StartPath$, EndPath$
    Kill StartPath$
    Exit Sub
End Sub
Public Sub CopyFile(StartPath As String, EndPath As String)
    FileCopy StartPath$, EndPath$
    Exit Sub
End Sub
Public Sub DeleteFile(FilePath As String)
    Kill FilePath$
    Exit Sub
End Sub
Public Sub ExecuteFile(FilePath As String)
    Ret = Shell("rundll32.exe url.dll,FileProtocolHandler " & (FilePath))
    Exit Sub
End Sub
Public Sub DisableCtrlAltDel()
    Dim Ret As Integer
    Dim pOld As Boolean
    Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
    Exit Sub
End Sub
Public Sub EnableCtrlAltDel()
    Dim Ret As Integer
    Dim pOld As Boolean
    Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
    Exit Sub
End Sub
Public Sub HideCtrlAltDel()
    App.TaskVisible = False
    Exit Sub
End Sub
Public Sub ShowCtrlAltDel()
    App.TaskVisible = True
    Exit Sub
End Sub
Public Sub OpenCD()
    retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
    Exit Sub
End Sub
Public Sub CloseCD()
    retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
    Exit Sub
End Sub
Public Sub PrintBlankPage()
    Printer.NewPage
    Exit Sub
End Sub
Public Sub PrintText(Text As String, MarginSize As Integer, AmountOfChrsInOneLine As Integer, JustUseDefault As Boolean)
    On Error Resume Next
    Screen.MousePointer = 11
    If JustUseDefault = True Then
    MarginSize = 10
    AmountOfChrsInOneLine = 65
    End If
    NumLinesOnPageToPrint = 60
    If NextPageNum% > 0 Then NextPageNum% = 0
    NextPageNum% = FirstPageNum% + NextPageNum% + 1
    TotalPageCount% = 1
    Call PrintPage(Text$, MarginSize, AmountOfChrsInOneLine)
    PrintEndOfLastPage
    Screen.MousePointer = 0
    Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub PrintPage(TextString, Margin As Integer, Length_ChrsInlineOfText As Integer)
    On Error Resume Next
    Dim ChrPosition
    Dim AllChrsInThisLineOfText
    Dim PlaceInLineOfText As Integer
    ChrPosition = 1
    Printer.FontSize = 18
    Printer.Print Tab(MarginSize%);
    LineNum% = 1
    Do While ChrPosition < Len(TextString)
    AllChrsInThisLineOfText = Mid$(TextString, ChrPosition, Length_ChrsInlineOfText%)
    If ChrPosition + Len(AllChrsInThisLineOfText) < Len(TextString) Then
    For PlaceInLineOfText% = Len(AllChrsInThisLineOfText) To 1 Step -1
    If Mid$(AllChrsInThisLineOfText, PlaceInLineOfText%, 1) = Chr$(32) Then
    CheckThisLineNum% = 1
    PrintNewPage
    If InStr(1, AllChrsInThisLineOfText, Chr$(10), 1) > 0 Then
    CheckThisLineNum% = 1
    PrintNewPage
    PlaceInLineOfText% = InStr(1, AllChrsInThisLineOfText, Chr$(10), 1)
    LineNum% = LineNum% + 1
    End If
    If Mid$(TextString, ChrPosition, PlaceInLineOfText%) <> Chr$(13) + Chr$(10) Then
    Printer.Print Tab(MarginSize%);
    Printer.Print Mid$(TextString, ChrPosition, PlaceInLineOfText%)
    LineNum% = LineNum% + 1
    Else
    LineNum% = LineNum% - 1
    End If
    ChrPosition = ChrPosition + PlaceInLineOfText%
    PlaceInLineOfText% = 0
    End If
    Next
    Else
    CheckThisLineNum% = 1
    PrintNewPage
    Printer.Print Tab(Margin%);
    Printer.Print AllChrsInThisLineOfText
    ChrPosition = Len(TextString)
    LineNum% = LineNum% + 1
    End If
    Loop
    Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub PrintNewPage()
    On Error Resume Next
    If LineNum% + CheckThisLineNum% >= NumLinesOnPageToPrint% Then
    Printer.Print ""
    Printer.Print Tab(MarginSize%);
    Printer.Print "(continued on page " + CStr(NextPageNum%) + ")"
    Printer.NewPage
    TotalPageCount% = TotalPageCount% + 1
    Printer.Print Tab(MarginSize%);
    Printer.Print "Page " + CStr(NextPageNum%)
    Printer.Print ""
    Printer.Print ""
    NextPageNum% = NextPageNum% + 1
    LineNum% = 3
    End If
    CheckThisLineNum% = 0
    Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub PrintEndOfLastPage()
    On Error Resume Next
    If LineNum% + 2 > NumLinesOnPageToPrint% Then
    Printer.NewPage
    TotalPageCount% = TotalPageCount% + 1
    Printer.Print Tab(MarginSize%);
    Printer.Print "Page " + CStr(NextPageNum%)
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(MarginSize%);
    Else
    Printer.Print ""
    Printer.Print Tab(MarginSize%);
    End If
    Printer.EndDoc
    Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Public Sub MakeStartupReg(AppTitle As String)
    a = MakeRegFile(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AppTitle$, App.Path & "\" & App.EXEName & ".exe")
    Exit Sub
End Sub
Public Sub AddToStartupDir()
    FileCopy App.Path & "\" & App.EXEName & ".EXE", Mid$(App.Path, 1, 3) & "WINDOWS\START MENU\PROGRAMS\STARTUP\" & App.EXEName & ".EXE"
    Exit Sub
End Sub
Private Function MakeRegFile(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    Dim lCreate As Long
    RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, _
    KEY_ALL_ACCESS, SA, phkResult, lCreate
    lResult = RegSetValueEx(phkResult, sSetValue, 0, 1, sValue, _
    CLng(Len(sValue) + 1))
    RegCloseKey phkResult
    MakeRegFile = (lResult = ERROR_SUCCESS)
    Exit Function
End Function
Public Sub ExecuteNewProgram()
    Ret = Shell("rundll32.exe url.dll,FileProtocolHandler " & App.Path & "\" & App.EXEName & ".EXE")
    Exit Sub
End Sub
Public Sub AlwaysOnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
    Exit Sub
End Sub
Public Sub NotAlwaysOnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags)
    Exit Sub
End Sub
Public Sub InvisibleForm(Frm As Form)
    Dim rctClient As RECT, rctFrame As RECT
    Dim hClient As Long, hFrame As Long
    GetWindowRect Frm.hWnd, rctFrame
    GetClientRect Frm.hWnd, rctClient
    Dim lpTL As POINTAPI, lpBR As POINTAPI
    lpTL.x = rctFrame.Left
    lpTL.y = rctFrame.Top
    lpBR.x = rctFrame.Right
    lpBR.y = rctFrame.Bottom
    ScreenToClient Frm.hWnd, lpTL
    ScreenToClient Frm.hWnd, lpBR
    rctFrame.Left = lpTL.x
    rctFrame.Top = lpTL.y
    rctFrame.Right = lpBR.x
    rctFrame.Bottom = lpBR.y
    rctClient.Left = Abs(rctFrame.Left)
    rctClient.Top = Abs(rctFrame.Top)
    rctClient.Right = rctClient.Right + Abs(rctFrame.Left)
    rctClient.Bottom = rctClient.Bottom + Abs(rctFrame.Top)
    rctFrame.Right = rctFrame.Right + Abs(rctFrame.Left)
    rctFrame.Bottom = rctFrame.Bottom + Abs(rctFrame.Top)
    rctFrame.Top = 0
    rctFrame.Left = 0
    hClient = CreateRectRgn(rctClient.Left, rctClient.Top, rctClient.Right, rctClient.Bottom)
    hFrame = CreateRectRgn(rctFrame.Left, rctFrame.Top, rctFrame.Right, rctFrame.Bottom)
    CombineRgn hFrame, hClient, hFrame, RGN_XOR
    SetWindowRgn Frm.hWnd, hFrame, True
    Exit Sub
End Sub
Public Sub ClipboardCopy(Text As String)
    Clipboard.Clear
    Clipboard.SetText Text$
    Exit Sub
End Sub
Function ClipboardGet()
    ClipboardGet = Clipboard.GetText
    Exit Function
End Function
Public Sub ClearClipboard()
    Clipboard.Clear
    Exit Sub
End Sub
Function FractionToDecimal(numerator As Integer, denominator As Integer) As Long
    FractionToDecimal = numerator / denominator
    Exit Function
End Function
Function DecimalToPercentage(DecimalNum As Long) As String
    DecimalToPercentage = (DecimalNum * 100) & "%"
    Exit Function
End Function
Function PercentageToDeciaml(PercentNum As String) As Long
    If Mid$(PercentNum$, Len(PercentNum$), 1) = "%" Then
    PercentNum$ = Mid$(PercentNum$, 2, Len(PercentNum$) - 1)
    End If
    PercentageToDecimal = Val(PercentNum$) / 100
    Exit Function
End Function
Function AreaOfCircle(radius As Long)
    pi = 3.141592654
    AreaOfCircle = pi * (radius ^ 2)
    Exit Function
End Function
Function Circumference(radius As Long)
    pi = 3.141592654
    Circumference = pi * 2 * radius
    Exit Function
End Function
Function AreaOfSquare(side As Long)
    AreaOfSquare = side ^ 2
    Exit Function
End Function
Function PerimeterOfSquare(side As Long)
    PerimeterOfSquare = 4 * side
    Exit Function
End Function
Function PerimeterOfRectangle(Length As Long, Width As Long)
    PerimeterOfRectangle = (2 * Length) + (2 * Width)
    Exit Function
End Function
Function AreaOfRectangle(Length As Long, Width As Long)
    AreaOfRectangle = Length * Width
    Exit Function
End Function
Function AreaOfTriangle(base As Long, Height As Long)
    AreaOfTriangle = (1 / 2) * base * Height
    Exit Function
End Function
Function PerimeterOfTriangle(side1 As Long, side2 As Long, side3 As Long)
    PerimeterOfTriangle = side1 + side2 + side3
    Exit Function
End Function
Function VolumeOfCube(edge As Long)
    VolumeOfCube = edge ^ 3
    Exit Function
End Function
Function VolumeOfPrism(base As Long, Height As Long)
    VolumeOfPrism = base * Height
    Exit Function
End Function
Function VolumeOfSphere(radius As Long)
    pi = 3.141592654
    VolumeOfSphere = (4 / 3) * pi * (radius ^ 3)
    Exit Function
End Function
Function VolumeOfPyramid(base As Long, Height As Long)
    VolumeOfPyramid = (1 / 3) * base * Height
    Exit Function
End Function
Function VolumeOfCone(radius As Long, Height As Long)
    pi = 3.141592654
    VolumeOfCone = (1 / 3) * pi * (radius ^ 2) * Height
    Exit Function
End Function
Function VolumeOfCylinder(radius As Long, Height As Long)
    pi = 3.141592654
    VolumeOfCylinder = pi * Height * (radius ^ 2)
    Exit Function
End Function
Function FadeThreeColorHTML(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$)
    textlen% = Len(TheText)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, textlen% - fstlen%)
    textlen% = Len(part1$)
    For i = 1 To textlen%
    TextDone$ = Left(part1$, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
    colorx2 = RGBToHex(ColorX)
    Faded1$ = Faded1$ + "<FONT COLOR=" & colorx2 & ">" + LastChr$ + "</FONT>"
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
    TextDone$ = Left(part2$, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBToHex(ColorX)
    Faded2$ = Faded2$ + "<FONT COLOR=" & colorx2 & ">" + LastChr$ + "</FONT>"
    Next i
    FadeThreeColorHTML = Faded1$ + Faded2$
    Exit Function
End Function
Private Function FadeTwoColorHTML(R1%, G1%, B1%, R2%, G2%, B2%, TheText$)
    On Error GoTo error
    textlen$ = Len(TheText)
    For i = 1 To textlen$
    TextDone$ = Left(TheText, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
    colorx2 = RGBToHex(ColorX)
    Faded$ = Faded$ + "<FONT COLOR=" & colorx2 & ">" + LastChr$ + "</FONT>"
    Next i
    FadeTwoColorHTML = Faded$
    Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
Function FadeThreeColorYahoo(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$)
    On Error GoTo error
    textlen% = Len(TheText)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, textlen% - fstlen%)
    textlen% = Len(part1$)
    For i = 1 To textlen%
    TextDone$ = Left(part1$, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
    colorx2 = RGBToHex(ColorX)
    Faded1$ = Faded1$ + "<#" & colorx2 & ">" + LastChr$
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
    TextDone$ = Left(part2$, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
    colorx2 = RGBToHex(ColorX)
    Faded2$ = Faded2$ + "<#" & colorx2 & ">" + LastChr$
    Next i
    FadeThreeColorYahoo = Faded1$ + Faded2$
    Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
Private Function FadeTwoColorYahoo(R1%, G1%, B1%, R2%, G2%, B2%, TheText$)
    On Error GoTo error
    textlen$ = Len(TheText)
    For i = 1 To textlen$
    TextDone$ = Left(TheText, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
    colorx2 = RGBToHex(ColorX)
    Faded$ = Faded$ + "<#" & colorx2 & ">" + LastChr$
    Next i
    FadeTwoColorYahoo = Faded$
    Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
Function FadeThreeColorANSI(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$)
    textlen% = Len(TheText)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, textlen% - fstlen%)
    textlen% = Len(part1$)
    For i = 1 To textlen%
    TextDone$ = Left(part1$, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
    colorx2 = RGBToHex(ColorX)
    Faded1$ = Faded1$ + "[#" & colorx2 & LastChr$
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
    TextDone$ = Left(part2$, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
    colorx2 = RGBToHex(ColorX)
    Faded2$ = Faded2$ + "[#" & colorx2 & LastChr$
    Next i
    FadeThreeColorANSI = Faded1$ + Faded2$
    Exit Function
End Function

Private Function FadeTwoColorANSI(R1%, G1%, B1%, R2%, G2%, B2%, TheText$)
    textlen$ = Len(TheText)
    For i = 1 To textlen$
    TextDone$ = Left(TheText, i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
    colorx2 = RGBToHex(ColorX)
    Faded$ = Faded$ + "[#" & colorx2 & LastChr$
    Next i
    FadeTwoColorANSI = Faded$
    Exit Function
End Function
Function AltCaps(Text As String)
    Dim i As Integer
    Dim S As String
    S = ""
    For i = 1 To Len(Text$)
      keyval = Asc(Mid$(Text$, i, 1))
      If (keyval >= 96 And keyval < 96 + 26) Or (keyval >= 64 And keyval < 64 + 26) Then
        If (i And 1) = 1 Then
          If keyval < 96 Then
            S = S + Chr$(96 + keyval - 64)
          Else
            S = S + Chr$(keyval)
          End If
        Else
          If keyval >= 96 Then
            S = S + Chr$(64 + keyval - 96)
          Else
            S = S + Chr$(keyval)
          End If
        End If
      Else
        S = S + Chr$(keyval)
      End If
    Next i
    Text$ = S
    AltCaps = Text$
Exit Function
End Function
Function GetAppVersion()
    AppVersion = App.Major & "." & App.Minor & "." & App.Revision
    Exit Function
End Function
Function GetAppName(ShowEXE As Boolean)
    GetAppName = App.EXEName
    If ShowEXE = True Then
    GetAppName = GetAppName & ".exe"
    End If
    Exit Function
End Function
Function GetAppPath()
    GetAppPath = App.Path
    Exit Function
End Function
Function GetAppDescription()
    GetAppDescription = App.FileDescription
    Exit Function
End Function
Function GetAppCopyRight()
    GetAppCopyRight = App.LegalCopyright
    Exit Function
End Function
Function GetAppComment()
    GetAppComment = App.Comments
    Exit Function
End Function
Function GetAppTitle()
    GetAppTitle = App.Title
    Exit Function
End Function
Function GetAppCompanyName()
    GetAppCompanyName = App.CompanyName
    Exit Function
End Function
Function GetAppProductName()
    GetAppProductName = App.ProductName
    Exit Function
End Function
Public Sub LeftClick()
    LeftDown
    LeftUp
    Exit Sub
End Sub
Public Sub LeftDown()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Exit Sub
End Sub
Public Sub LeftUp()
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Exit Sub
End Sub
Public Sub MiddleClick()
    MiddleDown
    MiddleUp
    Exit Sub
End Sub
Public Sub MiddleDown()
    mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
    Exit Sub
End Sub
Public Sub MiddleUp()
    mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
    Exit Sub
End Sub
Public Sub RightClick()
    RightDown
    RightUp
    Exit Sub
End Sub
Public Sub RightDown()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    Exit Sub
End Sub
Public Sub RightUp()
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Exit Sub
End Sub
Public Sub HideMouse()
    ShowCursor (bShow = False)
    Exit Sub
End Sub
Public Sub ShowMouse()
    ShowCursor (bShow = True)
    Exit Sub
End Sub
Public Sub DrawSquareOnForm(Frm As Form, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, Red As Integer, Green As Integer, Blue As Integer, Solid As Boolean)
    If Solid = True Then
    Frm.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), BF
    Else
    Frm.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), B
    End If
    Exit Sub
End Sub
Public Sub DrawLineOnForm(Frm As Form, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, Red As Integer, Green As Integer, Blue As Integer)
    Frm.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue)
    Exit Sub
End Sub
Public Sub DrawSquareOnPictureBox(Picture As PictureBox, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, Red As Integer, Green As Integer, Blue As Integer, Solid As Boolean)
    If Solid = True Then
    Picture.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), BF
    Else
    Picture.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), B
    End If
    Exit Sub
End Sub

Public Sub DrawLineOnPictureBox(Picture As PictureBox, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, Red As Integer, Green As Integer, Blue As Integer)
    Picture.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue)
    Exit Sub
End Sub
Function ConvertRGBToHex(Red As Double, Green As Double, Blue As Double)
    ConvertRGBToHex = RGBToHex(RGB(Blue, Green, Red))
    Exit Function
End Function
Function RGBToHex(RGB)
    Dim a As String
    Dim B As Integer
    a$ = Hex(RGB)
    B% = Len(a$)
    If B% = 5 Then a$ = "0" & a$
    If B% = 4 Then a$ = "00" & a$
    If B% = 3 Then a$ = "000" & a$
    If B% = 2 Then a$ = "0000" & a$
    If B% = 1 Then a$ = "00000" & a$
    RGBToHex = a$
Exit Function
End Function
Function ConvertHexToRGB(HexCode As String)
    HexCode$ = Mid$(HexCode$, 1, 6)
    ConvertHexToRGB = HexToRGB(HexCode$)
    Exit Function
End Function
Function HexToRGB(H As String) As Currency
    Dim tmp$
    Dim lo1 As Integer, lo2 As Integer
    Dim hi1 As Long, hi2 As Long
    Const Hx = "&H"
    Const BigShift = 65536
    Const LilShift = 256, Two = 2
    tmp = H
    If UCase(Left$(H, 2)) = "&H" Then tmp = Mid$(H, 3)
    tmp = Right$("0000000" & tmp, 8)
    If IsNumeric(Hx & tmp) Then
    lo1 = CInt(Hx & Right$(tmp, Two))
    hi1 = CLng(Hx & Mid$(tmp, 5, Two))
    lo2 = CInt(Hx & Mid$(tmp, 3, Two))
    hi2 = CLng(Hx & Left$(tmp, Two))
    HexToRGB = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
    End If
    Exit Function
End Function
Public Sub WebPage(Address As String)
    On Error GoTo error
    Ret = Shell("Start.exe " & Address, 0)
    Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Function RandomNumber(Max As Double, Min As Double)
    Randomize Timer
    RandomNumber = Int((Max - Min + 1) * Rnd + Min)
    Exit Function
End Function
Function MakeInputBox(DefaultText As String, Question As String, Title As String)
    MakeInputBox = InputBox(Question$, Title$, DefaultText$)
    Exit Function
End Function
Function LengthOfString(Text As String) As Long
    LengthOfString = Len(Text$)
    Exit Function
End Function
Function FindAsciiOfChr(Chr As String)
    FindAsciiOfChr = Asc(Mid$(Chr$, 1, 1))
    Exit Function
End Function
Function MakeChrFromAscii(Ascii As Integer)
    MakeChrFromAscii = Chr$(Ascii)
    Exit Function
End Function
Function MakeRndChrString(Length As Integer, Numbers As Boolean, Letters As Boolean, Symbols As Boolean, Other As Boolean) As String
    Dim chrasc As Integer
    Dim i As Integer
    Dim endstr As String
    Randomize Timer
    If Length > 100 Then
    Length = 100
    ElseIf Length < 1 Then
    Length = 1
    End If
    For i = 1 To Length
Start:
chrasc = Int((255 - 1 + 1) * Rnd + 1)
  If chrasc < 34 Then
    If Other = False Then
    GoTo Start
    End If
  ElseIf chrasc > 33 And chrasc < 49 Then
    If Symbols = False Then
    GoTo Start
    End If
  ElseIf chrasc > 48 And chrasc < 58 Then
    If Numbers = False Then
    GoTo Start
    End If
  ElseIf chrasc > 57 And chrasc < 65 Then
    If Symbols = False Then
    GoTo Start
    End If
  ElseIf chrasc > 64 And chrasc < 91 Then
    If Letters = False Then
    GoTo Start
    End If
  ElseIf chrasc > 90 And chrasc < 97 Then
    If Symbols = False Then
    GoTo Start
    End If
  ElseIf chrasc > 96 And chrasc < 123 Then
    If Letters = False Then
    GoTo Start
    End If
  ElseIf chrasc > 122 And chrasc < 127 Then
    If Symbols = False Then
    GoTo Start
    End If
  Else
    If Other = False Then
    GoTo Start
    End If
  End If
endstr$ = endstr$ & Chr$(chrasc)
Next i
MakeRndChrString = endstr$
Exit Function
End Function
Public Sub DoSendKeys(AppToActivate As String, AppActivateDelay As Integer, TextToSend As String, SendKeysDelay As Integer)
    AppActivate AppToActivate$, AppActivateDelay
    SendKeys TextToSend$, SendKeysDelay
    Exit Sub
End Sub
Function GetTextFromListBox(ListB As ListBox, ListIndex As Long) As String
    GetTextFromListBox = ListB.List(ListIndex)
    Exit Function
End Function
Function GetTextFromComboBox(ComboB As ComboBox, ListIndex As Long) As String
    GetTextFromComboBox = ComboB.List(ListIndex)
    Exit Function
End Function
Function PasswordLock(Password As String)
    Dim xtra As String
Start:
    xtra$ = InputBox("Please enter the password.", "Password Lock")
    If xtra$ = Password$ Then
    MsgBox "Correct Password!", vbExclamation, "Password Lock"
Else
  If MsgBox("Incorrect Password!  Would you like to try again?", 48 + vbYesNo, "Password Lock") = vbYes Then
  GoTo Start
  Else
  End
  End If
End If
    Exit Function
End Function
Public Sub ChangeDefaultDir(NewDirPath As String)
    ChDir NewDirPath$
    Exit Sub
End Sub
Public Sub ChangeDefaultDrive(NewDrive As String)
    ChDrive NewDrive$
    Exit Sub
End Sub
Public Sub MakeRegistrySetting(RegPath As String, Title As String, Data As String)
    a = MakeRegFile(&H80000002, RegPath$, Title$, Data$)
    Exit Sub
End Sub
Function ArcCosh(u As Double) As Double
    ArcCosh = Log(u + Sqr(u * u - 1))
End Function
Function ArcSinh(u As Double) As Double
    ArcSinh = Log(u + Sqr(u * u + 1))
End Function
Function ArcTanh(u As Double) As Double
    ArcTanh = Log((1 + u) / (1 - u)) / 2
End Function
Function ArcSech(u As Double) As Double
    ArcSech = Log((Sqr(-u * u + 1) + 1) / u)
End Function
Function ArcCsch(u As Double) As Double
    ArcCsch = Log((Sgn(u) * Sqr(u * u + 1) + 1) / u)
End Function
Function ArcCsc(u As Double) As Double
    ArcCsc = Atn(u / Sqr(u * u - 1)) + (Sgn(u) - 1) * (2 * Atn(1))
End Function
Function ArcCoth(u As Double) As Double
    ArcCoth = Log((u + 1) / (u - 1)) / 2
End Function
Function Cosh(u As Double) As Double
    Cosh = (Exp(u) + Exp(-u)) / 2
End Function
Function ArcCos(u As Double) As Double
    ArcCos = Atn(-u / Sqr(-u * u + 1)) + (2 * Atn(1))
End Function
Function Sech(u As Double) As Double
    Sech = 2 / (Exp(u) + Exp(-u))
End Function
Function ArcSec(u As Double) As Double
    ArcSec = Atn(u / Sqr(u * u - 1)) + Sgn(u - 1) * (2 * Atn(1))
End Function
Function ArcTan(u As Double) As Double
    ArcTan = Atn(u)
End Function
Function ArcCot(u As Double) As Double
    ArcCot = Atn(u) + 2 * Atn(1)
End Function
Function Cot(u As Double) As Double
    Cot = 1 / Tan(u)
End Function
Function Coth(u As Double) As Double
    Coth = (Exp(u) + Exp(-u)) / (Exp(u) - Exp(-u))
End Function
Function Csc(u As Double) As Double
    Csc = 1 / Sin(u)
End Function
Function Sec(u As Double) As Double
    Sec = 1 / Cos(u)
End Function
Function Csch(u As Double) As Double
    Csch = 2 / (Exp(u) - Exp(-u))
End Function
Function Sinh(u As Double) As Double
    Sinh = (Exp(u) - Exp(-u)) / 2
End Function
Function ArcSin(u As Double) As Double
    ArcSin = Atn(u / Sqr(-u * u + 1))
End Function
Function Tanh(u As Double) As Double
    Tanh = (Exp(u) - Exp(-u)) / (Exp(u) + Exp(-u))
End Function
Function XInverse(u As Double) As Double
    XInverse = u ^ -1
End Function
Function Log10(u As Double) As Double
    Log10 = Log(u) / Log(10)
End Function
Function XSuqare(u As Double) As Double
    Xsquare = u ^ 2
End Function
Function XtoY(u As Double, v As Double) As Double
    XtoY = u ^ v
End Function
Function XRoot(uNumber As Double, vRoot As Double) As Double
    XRoot = uNumber ^ (1 / vRoot)
End Function
Function LogX(u As Double, v As Double) As Double
    LogX = Log(u) / Log(v)
End Function
Function Cbr(u As Double) As Double
    Cbr = u ^ (1 / 3)
End Function
Function XCubed(u As Double) As Double
    XCubed = u ^ 3
End Function
Function Deg2Rad(u As Double) As Double
    On Error Resume Next
    Deg2Rad = u * (3.14159265358979 / 180)
End Function
Function Rad2Deg(u As Double) As Double
    On Error Resume Next
    Rad2Deg = u * (180 / 3.14159265358979)
End Function
Function Dial(ByVal Connection As String, ByVal UserName As String, ByVal Password As String) As Boolean
    Dim rp As RASDIALPARAMS, H As Long, resp As Long
    rp.dwSize = Len(rp) + 6
    ChangeBytes Connection, rp.szEntryName
    ChangeBytes "", rp.szPhoneNumber
    ChangeBytes "*", rp.szCallbackNumber
    ChangeBytes UserName, rp.szUserName
    ChangeBytes Password, rp.szPassword
    ChangeBytes "*", rp.szDomain
    resp = RasDial(ByVal 0, ByVal 0, rp, 0, ByVal 0, H)
    Dial = (resp = 0)
End Function
Function ChangeToStringUni(Bytes() As Byte) As String
    Dim temp As String
    temp = StrConv(Bytes, vbUnicode)
    ChangeToStringUni = Left(temp, InStr(temp, Chr(0)) - 1)
End Function
Function ChangeBytes(ByVal str As String, Bytes() As Byte) As Boolean
    Dim lenBs As Long
    Dim lenStr As Long
    lenBs = UBound(Bytes) - LBound(Bytes)
    lenStr = LenB(StrConv(str, vbFromUnicode))
    If lenBs > lenStr Then
        CopyMemory Bytes(0), str, lenStr
        ZeroMemory Bytes(lenStr), lenBs - lenStr
    ElseIf lenBs = lenStr Then
        CopyMemory Bytes(0), str, lenStr
    Else
        CopyMemory Bytes(0), str, lenBs
        ChangeBytes = True
    End If
End Function
Function DragObject(ByVal hWnd As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, LP_HT_CAPTION, 0
End Function
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    On Error Resume Next
      
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
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
      OleCreatePictureIndirect Pic, IID_IDispatch, 1, IPic
      Set CreateBitmapPicture = IPic
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    On Error Resume Next
      
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
      
    If Client Then
        hDCSrc = GetDC(hWndSrc)
    Else
        hDCSrc = GetWindowDC(hWndSrc)
      End If
      hDCMemory = CreateCompatibleDC(hDCSrc)
      
      hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
      hBmpPrev = SelectObject(hDCMemory, hBmp)
      
      RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
      HasPaletteScrn = RasterCapsScrn And RC_PALETTE
      PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
      
      If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        GetSystemPaletteEntries hDCSrc, 0, 256, LogPal.palPalEntry(0)
        hPal = CreatePalette(LogPal)
        
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        RealizePalette hDCMemory
    End If
    
      BitBlt hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy
    
      hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
      
      DeleteDC hDCMemory
      ReleaseDC hWndSrc, hDCSrc
    
      Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CaptureForm(frmSrc As Form) As Picture
    On Error Resume Next
    Set CaptureForm = CaptureWindow(frmSrc.hWnd, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function
Public Function CaptureClient(frmSrc As Form) As Picture
    On Error Resume Next
    Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function
Public Function CaptureArea(frmSrc As Form, Left As Long, Top As Long, Width As Long, Height As Long) As Picture
    On Error Resume Next
    Set CaptureArea = CaptureWindow(frmSrc.hWnd, True, Left, Top, Width, Height)
End Function
Public Function CaptureActiveWindow() As Picture
    On Error Resume Next
    
    Dim hWndActive As Long
    Dim RectActive As RECT
    
    hWndActive = GetForegroundWindow()
    GetWindowRect hWndActive, RectActive
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
End Function
Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    On Error Resume Next
      
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double
    
    PicRatio = Pic.Width / Pic.Height
    
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    PrnRatio = PrnWidth / PrnHeight
    
    If PicRatio >= PrnRatio Then
      
      PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
      PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
      
    Else
      
      PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
      PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
      
    End If
    
    Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
    DoEvents
    Prn.EndDoc
End Sub
Public Function CaptureScreen() As Picture
    On Error Resume Next
    Dim hWndScreen As Long
    hWndScreen = GetDesktopWindow()
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
Public Function EnumeratingWindows(ByVal hWnd As Long, ByVal lParam As Long, ByVal Form As Form) As Boolean
    Dim sSave As String, Ret As Long
    Ret = GetWindowTextLength(hWnd)
    sSave = Space(Ret)
    GetWindowText hWnd, sSave, Ret + 1
    Form.Print str$(hWnd) + " " + sSave
    EnumeratingWindows = True
End Function
Function TerminateProcess()
    On Error Resume Next
    ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
End Function
Function TerminateThreads()
    ExitThread GetExitCodeThread(GetCurrentThread, 0)
End Function
Function DrawAnIcon(ByVal HDC As Long, ByVal uIcon As String) As Long
    Dim mIcon As Long
    mIcon = ExtractAssociatedIcon(App.hInstance, uIcon, 2)
    DrawIconEx HDC, 0, 0, mIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon mIcon
End Function
Function DisableClose(ByVal hWnd As Long)
    Dim hSysMenu As Long, nCnt As Long
    hSysMenu = GetSystemMenu(hWnd, False)
    If hSysMenu Then
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
            DrawMenuBar hWnd
        End If
    End If
End Function
Function SearchFiles(ByVal uDrive As String, ByVal vFile As String)
    Dim tempStr As String, Ret As Long
    tempStr = String(MAX_PATH, 0)
    Ret = SearchTreeForFile(uDrive, vFile, tempStr)
    If Ret <> 0 Then
        MsgBox "Located file at " + Left$(tempStr, InStr(1, tempStr, Chr$(0)) - 1)
    Else
        MsgBox "File not found!"
    End If
End Function
Function AnimateWindows(ByVal Form As Form, ByVal hWnd As Long) As Long
    On Error Resume Next
    Form.AutoRedraw = True
    AnimateWindow hWnd, 200, AW_VER_POSITIVE Or AW_HOR_NEGATIVE Or AW_SLIDE
End Function
Function ProcessTimings()
    Dim FT0 As FILETIME, FT1 As FILETIME, ST As SYSTEMTIME
    GetProcessTimes GetCurrentProcess, FT1, FT0, FT0, FT0
    FileTimeToLocalFileTime FT1, FT1
    FileTimeToSystemTime FT1, ST
    MsgBox "This program was started at " + CStr(ST.wHour) + ":" + CStr(ST.wMinute) + "." + CStr(ST.wSecond) + " on " + CStr(ST.wMonth) + "/" + CStr(ST.wDay) + "/" + CStr(ST.wYear)
End Function
Function ChangeActiveCaptionBarColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_ACTIVECAPTION)
    t& = SetSysColors(1, COLOR_ACTIVECAPTION, uColor)
End Function
Function ChangeScrollBarColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_SCROLLBAR)
    t& = SetSysColors(1, COLOR_SCROLLBAR, uColor)
End Function
Function ChangeBackgroundColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_BACKGROUND)
    t& = SetSysColors(1, COLOR_BACKGROUND, uColor)
End Function
Function ChangeInactiveCaptionBarColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_INACTIVECAPTION)
    t& = SetSysColors(1, COLOR_INACTIVECAPTION, uColor)
End Function
Function ChangeMenuColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_MENU)
    t& = SetSysColors(1, COLOR_MENU, uColor)
End Function
Function ChangeWindowColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_WINDOW)
    t& = SetSysColors(1, COLOR_WINDOW, uColor)
End Function
Function ChangeWindowFrameColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_WINDOWFRAME)
    t& = SetSysColors(1, COLOR_WINDOWFRAME, uColor)
End Function
Function ChangeMenuTextColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_MENUTEXT)
    t& = SetSysColors(1, COLOR_MENUTEXT, uColor)
End Function
Function ChangeWindowTextColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_WINDOWTEXT)
    t& = SetSysColors(1, COLOR_WINDOWTEXT, uColor)
End Function
Function ChangeCaptionTextColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_CAPTIONTEXT)
    t& = SetSysColors(1, COLOR_CAPTIONTEXT, uColor)
End Function
Function ChangeActiveBorderColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_ACTIVEBORDER)
    t& = SetSysColors(1, COLOR_ACTIVEBORDER, uColor)
End Function
Function ChangeInactiveBorderColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_INACTIVEBORDER)
    t& = SetSysColors(1, COLOR_INACTIVEBORDER, uColor)
End Function
Function ChangeAppWorkspaceColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_APPWORKSPACE)
    t& = SetSysColors(1, COLOR_APPWORKSPACE, uColor)
End Function
Function ChangeHighlightColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_HIGHLIGHT)
    t& = SetSysColors(1, COLOR_HIGHLIGHT, uColor)
End Function
Function ChangeHighlightTextColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_HIGHLIGHTTEXT)
    t& = SetSysColors(1, COLOR_HIGHLIGHTTEXT, uColor)
End Function
Function ChangeButtonFaceColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_BTNFACE)
    t& = SetSysColors(1, COLOR_BTNFACE, uColor)
End Function
Function ChangeButtonShadowColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_BTNSHADOW)
    t& = SetSysColors(1, COLOR_BTNSHADOW, uColor)
End Function
Function ChangeGrayTextColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_GRAYTEXT)
    t& = SetSysColors(1, COLOR_GRAYTEXT, uColor)
End Function
Function ChangeButtonTextColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_BTNTEXT)
    t& = SetSysColors(1, COLOR_BTNTEXT, uColor)
End Function
Function ChangeInactiveCaptionTextColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
    t& = SetSysColors(1, COLOR_INACTIVECAPTIONTEXT, uColor)
End Function
Function ChangeButtonHighlightColor(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_BTNHIGHLIGHT)
    t& = SetSysColors(1, COLOR_BTNHIGHLIGHT, uColor)
End Function
Function ChangeInactiveCaptionBar2Color(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_2NDINACTIVECAPTION)
    t& = SetSysColors(1, COLOR_ACTIVECAPTION, uColor)
End Function
Function ChangeActiveCaptionBarColor2(ByVal uColor As Long) As Long
    col& = GetSysColor(COLOR_2NDACTIVECAPTION)
    t& = SetSysColors(1, COLOR_2NDACTIVECAPTION, uColor)
End Function
Function MemoryStatus() As Long
    Dim MemStat As MemoryStatus
    GlobalMemoryStatus MemStat
    MsgBox "You have" + str$(MemStat.dwTotalPhys / 1024) + " KB total memory and" + str$(MemStat.dwAvailPageFile / 1024) + " KB Left."
End Function
Function IsInternetConnected()
    Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 1 Then
        MsgBox "You are connected to Internet via a " & sConnType, vbInformation
    Else
        MsgBox "You are not connected to internet", vbInformation
    End If
End Function
Function PrinterProps(hWnd As LP_HWND)
    Dim Result As LP_RESULT
    Dim hPrinter As Long
    OpenPrinter Printer.DeviceName, hPrinter, ByVal 0&
    PrinterProperties hWnd.LP_HWND, hPrinter
    Result.LP_RESULT = ClosePrinter(hPrinter)
End Function
Function EnumerationJobs()
    Dim hPrinter As Long, lNeeded As Long, lReturned As Long
    Dim lJobCount As Long
    OpenPrinter Printer.DeviceName, hPrinter, ByVal 0&
    EnumJobs hPrinter, 0, 99, 1, ByVal 0&, 0, lNeeded, lReturned
    If lNeeded > 0 Then
        ReDim byteJobsBuffer(lNeeded - 1) As Byte
        EnumJobs hPrinter, 0, 99, 1, byteJobsBuffer(0), lNeeded, lNeeded, lReturned
        If lReturned > 0 Then
            lJobCount = lReturned
        Else
            lJobCount = 0
        End If
    Else
        lJobCount = 0
    End If
    ClosePrinter hPrinter
    MsgBox "Jobs in printer queue: " + CStr(lJobCount), vbInformation
End Function
Function DisconnectFromInternet()
    Dim i As Long, lpRasConn(255) As RasConn, lpcb As Long
    Dim lpcConnections As Long, hRasConn As Long
    lpRasConn(0).dwSize = RAS_RASCONNSIZE
    lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
    lpcConnections = 0
    returncode = RasEnumConnections(lpRasConn(0), lpcb, lpcConnections)

    If returncode = 0 Then
        For i = 0 To lpcConnections - 1
            hRasConn = lpRasConn(i).hRasConn
            returncode = RasHangUp(ByVal hRasConn)
        Next i
    End If
End Function
Function ChangeComputerName()
    Dim sNewName As LP_RAM
    Dim Result As LP_RESULT
    sNewName.LP_RAM = InputBox("Please enter a new computer name.")
    Result.LP_RESULT = SetComputerName(sNewName.LP_RAM)
    MsgBox "Computername set to " + sNewName.LP_RAM
End Function
Function QuickFormatDisk(ByVal hWnd As Long) As Long
    Dim Result As LP_RESULT
    Result.LP_RESULT = SHFormatDrive(hWnd, 0, SHFD_CAPACITY_DEFAULT, SHFD_FORMAT_QUICK)
End Function
Function FullFormatDisk(ByVal hWnd As Long) As Long
    Dim Result As LP_RESULT
    Result.LP_RESULT = SHFormatDrive(hWnd, 0, SHFD_CAPACITY_DEFAULT, SHFD_FORMAT_FULL)
End Function
Function FormatCopySYSFilesOnly(ByVal hWnd As Long) As Long
    Dim Result As LP_RESULT
    Result.LP_RESULT = SHFormatDrive(hWnd, 0, SHFD_CAPACITY_DEFAULT, SHFD_FORMAT_SYSONLY)
End Function
Function CreateDir(uDir As String) As String
    Dim Security As SECURITY_ATTRIBUTES
    Dim Result As LP_RESULT
    Result.LP_RESULT = CreateDirectory(uDir, Security)
    MsgBox "Directory " & uDir & " has been created.", vbInformation, "Directory Created!"
End Function
Function KillForms()
    Dim Result As LP_RESULT
    Dim Form As Form
    For Each Form In Forms
    Unload Form
    Set Form = Nothing
    Next Form
    Result.LP_RESULT = MsgBox("Forms Terminated!", vbExclamation, "Forms")
End Function
Function PlacePictureOnForm(Picture As String, Width As Long, Height As Long, Form As Form, hWnd As Long)
    Dim HDC As Long, hBitmap As Long
    hBitmap = LoadImage(App.hInstance, Picture, IMAGE_BITMAP, Width, Height, LR_LOADFROMFILE)
    If hBitmap = 0 Then
        MsgBox "There was an error while loading the bitmap"
        Exit Function
    End If
    OpenClipboard hWnd
    EmptyClipboard
    SetClipboardData CF_BITMAP, hBitmap
    If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then
        MsgBox "There was an error while pasting the bitmap to the clipboard!"
    End If
    CloseClipboard
    Form.Picture = Clipboard.GetData(vbCFBitmap)
    
    ' For Example,
    ' Private Sub Form_Load()
    ' Example: PlacePictureOnForm "C:\windows\setup.bmp", 640, 480, Me, Me.hWnd
    ' End Sub
End Function
