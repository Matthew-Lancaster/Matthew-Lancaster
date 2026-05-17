Attribute VB_Name = "Pooboos_1100"
'Jamal 331 Pooboos_1100 Funcitonal LPBI Module for Forms.

'Contact at Csjh218@yahoo.com

Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function lstrCat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long 'BOOL
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetIDListFromPath Lib "shell32.dll" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long  'MODULE 1171
Private Declare Function GetTickCount Lib "kernel32" () As Long   'MODULE 1117
Private Declare Function CloseClipboard Lib "user32" () As Long   'MODULE 1116
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer  'MODULE 1118
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long 'Parameter: form.(property) = Messagebox(0,"Text","Caption",(vbcritical,vbinformation))  'MODULE 1119
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long  'Parameter: form.(property) = SetCursorPos(x,y) in Pixels.  'MODULE 1121
Private Declare Function sndPlaySound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long  'MODULE 1126
Private Declare Function CreateSolidBrush Lib "GDI32.DLL" (ByVal crColor As Long) As Long  'MODULE 1127
Private Declare Function FillRect Lib "user32.dll" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long  'MODULE 1128
Private Declare Function DeleteObject Lib "GDI32.DLL" (ByVal hObject As Long) As Long  'MODULE 1129
Private Declare Function GetDiskFreeSpace Lib "kernel32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long  'MODULE 1164
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long  'MODULE 1131
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long  'MODULE 1133
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long  'MODULE 1134
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long  'MODULE 1135
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long  'MODULE 1136
Private Declare Function BeginPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long  'MODULE 1140

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142

Private Declare Function EndPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long  'MODULE 1143
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long  'MODULE 1144
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long  'MODULE 1145
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal HDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long  'MODULE 1146
Private Declare Function SetBkMode Lib "GDI32.DLL" (ByVal HDC As Long, ByVal nBkMode As Long) As Long  'MODULE 1148
Private Declare Function CreateFontIndirect Lib "GDI32.DLL" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long  'MODULE 1149
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function SetTextColor Lib "GDI32.DLL" (ByVal HDC As Long, ByVal crColor As Long) As Long  'MODULE 1150
Private Declare Function CreateRectRgn Lib "GDI32.DLL" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long  'MODULE 1151
Private Declare Function CreateRoundRectRgn Lib "GDI32.DLL" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long  'MODULE 1152
Private Declare Function CreatePolygonRgn Lib "GDI32.DLL" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long  'MODULE 1153
Private Declare Function CombineRgn Lib "GDI32.DLL" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long  'MODULE 1154
Private Declare Function FillRgn Lib "GDI32.DLL" (ByVal HDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long  'MODULE 1155
Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long  'MODULE 1156
Private Declare Function FrameRgn Lib "GDI32.DLL" (ByVal HDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long  'MODULE 1157
Private Declare Function GetStockObject Lib "GDI32.DLL" (ByVal nIndex As Long) As Long  'MODULE 1158
Private Declare Function SelectObject Lib "GDI32.DLL" (ByVal HDC As Long, ByVal hObject As Long) As Long  'MODULE 1159
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long  'MODULE 1166
Private Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean  'MODULE 1167
Private Declare Function ChangeDisplaySettings Lib "user32.dll" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long  'MODULE 1168
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function InitializeFlatSB Lib "comctl32" (ByVal hWnd As Long) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32" (ByVal hWnd As Long) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32" (ByVal hWnd As Long, ByVal Index As Long, ByVal newValue As Long, ByVal fRedraw As Boolean) As Boolean
Private Declare Function FlatSB_EnableScrollBar Lib "comctl32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32" (ByVal hWnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO) As Boolean
Private Declare Function FlatSB_GetScrollProp Lib "comctl32" (ByVal hWnd As Long, ByVal Index As Long, pValue As Long) As Boolean
Private Declare Function FlatSB_GetScrollRange Lib "comctl32" (ByVal hWnd As Long, ByVal code As Long, lpMinPos As Long, lpMaxPos As Long) As Boolean
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32" (ByVal hWnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollPos Lib "comctl32" (ByVal hWnd As Long, ByVal code As Long, ByVal nPos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollRange Lib "comctl32" (ByVal hWnd As Long, ByVal code As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32" (ByVal hWnd As Long, ByVal code As Long, ByVal fShow As Boolean) As Boolean
Private Declare Function FlatSB_GetScrollPos Lib "comctl32" (ByVal hWnd As Long, ByVal code As Long) As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal HDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Private Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function mciSendString Lib "Winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function mciGetErrorString Lib "Winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal hMem As Long)
Private Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)
Private Declare Sub ExitThread Lib "kernel32.dll" (ByVal dwExitCode As Long)

Const E = 2.7182818284
Const pi = 3.141592648
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const shrdNoMRUString = &H2
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Const LF_FACESIZE = 32
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
Const GRADIENT_FILL_RECT_H As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1
Const GRADIENT_FILL_TRIANGLE As Long = &H2
Const WM_NCLBUTTONDOWN = &HA1
Const LP_HT_CAPTION = 2
Const GWL_WNDPROC = -4
Const SWP_HIDEWINDOW = &H80
Const GW_CHILD = 5


Const GW_HWNDNEXT = 2
Const GWL_STYLE = (-16)
Const WS_BORDER = &H800000
Const FW_NORMAL = 400
Const FW_HEAVY = 900
Const FW_SEMIBOLD = 600
Const FW_BLACK = FW_HEAVY
Const FW_BOLD = 700
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_DONTCARE = 0
Const FW_EXTRABOLD = 800
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_MEDIUM = 500
Const FW_REGULAR = FW_NORMAL
Const FW_THIN = 100
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
Const TRANSPARENT = 1
Const ALTERNATE = 1
Const BLACK_BRUSH = 4
Const DKGRAY_BRUSH = 3
Const DT_SINGLELINE = &H20
Const DT_NOCLIP = &H100
Const DT_CENTER = &H1
Const DT_VCENTER = &H4
Const DT_WORDBREAK = &H10
Const DT_CALCRECT = &H400
Const CW_USEDEFAULT = &H80000000
Const TTS_ALWAYSTIP = 1
Const TTF_IDISHWND = 1
Const TTF_CENTERTIP = 2
Const TTF_RTLREADING = 4
Const TTF_SUBCLASS = &H10
Const TTF_TRACK = &H20
Const TTF_ABSOLUTE = &H80
Const TTF_TRANSPARENT = &H100
Const TTF_DI_SETITEM = &H8000
Const WM_USER = &H400
Const WM_PAINT = &HF
Const WM_PRINT = &H317
Const TTM_ACTIVATE = WM_USER + 1
Const TTM_SETDELAYTIME = WM_USER + 3
Const TTM_ADDTOOL = WM_USER + 4
Const TTM_DELTOOL = WM_USER + 5
Const TTM_NEWTOOLRECT = WM_USER + 6
Const TTM_RELAYEVENT = WM_USER + 7
Const COLOR_INFOTEXT = 23
Const COLOR_INFOBK = 24
Const COLOR_GRAYTEXT = 17
Const COLOR_3DLIGHT = 22
Const RGN_OR = 2
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const iOffset = 8
Const FontType = "Comic Sans MS" & vbNullChar
Const FontSize = 13
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000
Const WS_VSCROLL = &H200000
Const WS_HSCROLL = &H100000
Const WSB_PROP_CYVSCROLL = &H1
Const WSB_PROP_CXHSCROLL = &H2
Const WSB_PROP_CYHSCROLL = &H4
Const WSB_PROP_CXVSCROLL = &H8
Const WSB_PROP_CXHTHUMB = &H10
Const WSB_PROP_CYVTHUMB = &H20
Const WSB_PROP_VBKGCOLOR = &H40
Const WSB_PROP_HBKGCOLOR = &H80
Const WSB_PROP_VSTYLE = &H100
Const WSB_PROP_HSTYLE = &H200
Const WSB_PROP_WINSTYLE = &H400
Const WSB_PROP_PALETTE = &H800
Const WSB_PROP_MASK = &HFFF
Const FSB_FLAT_MODE = 2
Const FSB_ENCARTA_MODE = 1
Const FSB_REGULAR_MODE = 0
Const SB_HORZ = 0
Const SB_VERT = 1
Const SB_BOTH = 3
Const ESB_ENABLE_BOTH = &H0
Const ESB_DISABLE_BOTH = &H3
Const ESB_DISABLE_LEFT = &H1
Const ESB_DISABLE_RIGHT = &H2
Const ESB_DISABLE_UP = &H1
Const ESB_DISABLE_DOWN = &H2
Const ESB_DISABLE_LTUP = ESB_DISABLE_LEFT
Const ESB_DISABLE_RTDN = ESB_DISABLE_RIGHT
Const SIF_RANGE = &H1
Const SIF_PAGE = &H2
Const SIF_POS = &H4
Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS)
Const WS_CHILD = &H40000000
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const RSP_SIMPLE_SERVICE = 1
Const RSP_UNREGISTER_SERVICE = 0
Const GWL_EXSTYLE = (-20)
Const WS_EX_TRANSPARENT = &H20&
Const SWP_FRAMECHANGED = &H20
Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
Const EXIT_LOGOFF = 0
Const EXIT_SHUTDOWN = 1
Const EXIT_REBOOT = 2
Const SWP_NOZORDER = &H4
Const SWP_NOREPOSITION = &H200
Const WM_DESTROY = &H2
Const STARTF_NULL = 0
Const STARTF_FORCEONFEEDBACK = &H40
Const STARTF_FORCEOFFFEEDBACK = &H80
Const STARTF_RUNFULLSCREEN = &H20
Const STARTF_USECOUNTCHARS = &H8
Const STARTF_USEFILLATTRIBUTE = &H10
Const STARTF_USEPOSITION = &H4
Const STARTF_USESHOWWINDOW = &H1
Const STARTF_USESIZE = &H2
Const STARTF_USESTDHANDLES = &H100
Const CREATE_DEFAULT_ERROR_MODE = &H4000000
Const CREATE_NEW_CONSOLE = &H10
Const CREATE_NEW_PROCESS_GROUP = &H200
Const CREATE_SUSPENDED = &H4
Const CREATE_UNICODE_ENVIRONMENT = &H400
Const DETACHED_PROCESS = &H8
Const DEBUG_ONLY_THIS_PROCESS = &H2
Const DEBUG_PROCESS = &H1
Const REALTIME_PRIORITY_CLASS = &H100
Const HIGH_PRIORITY_CLASS = &H80
Const NORMAL_PRIORITY_CLASS = &H20
Const IDLE_PRIORITY_CLASS = &H40
Const MAX_PATH = 260
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREDRAW = &H8
Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const MF_BYCOMMAND = &H0&
Const BIF_BROWSEFORCOMPUTER = &H1000
Const BIF_BROWSEFORPRINTER = &H2000
Const BIF_BROWSEINCLUDEFILES = &H4000
Const BIF_RETURNONLYFSDIRS = &H1
Const BIF_DONTGOBELOWDOMAIN = &H2
Const BIF_STATUSTEXT = &H4
Const BIF_RETURNFSANCESTORS = &H8
Const BIF_VALIDATE = &H20
Const BIF_EDITBOX = &H10
Const OF_READ = &H0
Const OF_WRITE = &H1
Const OF_READWRITE = &H2
Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40
Const OF_PARSE = &H100
Const OF_DELETE = &H200
Const OF_VERIFY = &H400
Const OF_CANCEL = &H800
Const OF_CREATE = &H1000
Const OF_PROMPT = &H2000
Const OF_EXIST = &H4000
Const OF_REOPEN = &H8000

Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

'Type POINTAPI
'   x As Long
'   y As Long
'End Type

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'Public Declare Function ShellExecute _
'        Lib "shell32.dll" _
'        Alias "ShellExecuteA" _
'        (ByVal hWnd As Long, _
'         ByVal lpOperation As String, _
'         ByVal lpFile As String, _
'         ByVal lpParameters As String, _
'         ByVal lpDirectory As String, _
'         ByVal nShowCmd As Long) As Long

Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Type TOOLINFO
   cbSize As Long
   uFlags As Long
   hWnd As Long
   uId As Long
   hInst As Long
   lpszText As String
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type DEVMODE
   dmDeviceName As String * CCDEVICENAME
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
   dmFormName As String * CCFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Type PAINTSTRUCT
   HDC As Long
   fErase As Long
   rcPaint As RECT
   fRestore As Long
   fIncUpdate As Long
   rgbReserved(32) As Byte
End Type

Type TOldWndProc
   hWnd As Long
   lPrevWndProc As Long
End Type

Type OPENFILENAME
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

Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type

Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type LOGFONT
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
        lfFaceName As String * 31
End Type

Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long
        HDC As Long
        lpLogFont As Long
        iPointSize As Long
        flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
End Type

Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    HDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Type DEVMODE_TYPE
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

Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Type ULARGE_INTEGER
  lowpart As Long
  highpart As Long
End Type

Type SHQUERYRBINFO
  cbSize As Long
  i64Size As ULARGE_INTEGER
  i64NumItems As ULARGE_INTEGER
End Type

Type BROWSEINFO
  hwndOwner      As Long
  pidlRoot       As Long
  pszDisplayName As String
  lpszTitle      As String
  ulFlags        As Long
  lpfnCALLBACK   As Long
  lParam         As Long
  iImage         As Long
End Type

Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Type SECURITY_ATTRIBUTES
  nLength              As Long
  lpSecurityDescriptor As Long
  bInheritHandle       As Long
End Type

Type PROCESS_INFORMATION
  hProcess    As Long
  hThread     As Long
  dwProcessId As Long
  dwThreadId  As Long
End Type

Type STARTUPINFO
  cb              As Long
  lpReserved      As String
  lpDesktop       As String
  lpTitle         As String
  dwX             As Long
  dwY             As Long
  dwXSize         As Long
  dwYSize         As Long
  dwXCountChars   As Long
  dwYCountChars   As Long
  dwFillAttribute As Long
  dwFlags         As Long
  wShowWindow     As Integer
  cbReserved2     As Integer
  lpReserved2     As Long
  hStdInput       As Long
  hStdOutput      As Long
  hStdError       As Long
End Type

Enum Priorities
  p_RealTime = &H100
  p_Hight = &H80
  p_Normal = &H20
  p_Idle = &H40
End Enum

Enum WindowStates
  SW_HIDE = 0
  SW_SHOWNORMAL = 1
  SW_NORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_SHOWMAXIMIZED = 3
  SW_MAXIMIZE = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  SW_RESTORE = 9
  SW_SHOWDEFAULT = 10
  SW_FORCEMINIMIZE = 11
  SW_MAX = 11
End Enum

Public WndProc() As TOldWndProc
Public DevChg As DEVMODE
Public NumTips As Long
Public ExitingProgram As Boolean
Public TaskBarhWnd As Long

Dim error As Long
Dim TheStat As String * 128
Dim Word As Integer
Dim dword As Long
Dim DByte As Byte
Dim HShort As Integer
Dim Hlong As Long
Dim HInt As Integer
Dim Degrees As Single
Dim Radians As Single
Dim strResult As String
Dim Char As Byte
Dim OFName As OPENFILENAME
Dim CustomColors() As Byte

Dim r As RECT

Dim tWnd As Long, bWnd As Long, sSave As String * 250
Function FreeDiskSpace(ByVal Drive As String)  'MODULE 1100
   Dim SectorsPerCluster As Long
   Dim BytesPerSector As Long
   Dim FreeClusters As Long
   Dim NumberOfClusters As Long
   Dim Ret As Long
   Ret = GetDiskFreeSpace(Drive & ":\", SectorsPerCluster, BytesPerSector, FreeClusters, NumberOfClusters)
   FreeDiskSpace = BytesPerSector * SectorsPerCluster * FreeClusters
End Function
Sub MakeDirectory(ByVal Dir As String)  'MODULE 1101
    'Create Path. Parameter: MakeDir (object.property) with an object (text)
    MkDir (Dir)
End Sub
Sub DelTree(ByVal Dir As String)  'MODULE 1102
    'Deletes directory and subdirectories. Parameter: DelTree (object.property) with an object (text: "C:\")
    MsgBox "Delete Directory?", vbYesNo
    If vbYes Then
        RemoveDirectory (Dir)
    End If
    If vbNo Then
        Exit Sub
    End If
End Sub
Sub TimeDateSettingsDialog()  'MODULE 1103
    Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", vbNormalFocus
End Sub
Sub ModemsSettingsDialog()  'MODULE 1104
    Shell "rundll32.exe shell32.dll,Control_RunDLL modem.cpl", vbNormalFocus
End Sub
Sub ChangePasswordDialog()  'MODULE 1105
    Shell "rundll32.exe shell32.dll,Control_RunDll password.cpl", vbNormalFocus
End Sub
Sub DesktopSettingsDialog()
    Shell "rundll32.exe shell32.dll,Control_Rundll Desk.cpl", vbNormalFocus
End Sub
Function ArcSin(u As Double) As Double  'MODULE 1107
    ArcSin = Atn(u / Sqr(-u * u + 1))
End Function
Function ArcCos(u As Double) As Double  'MODULE 1108)
    ArcCos = Atn(-u / Sqr(-u * u + 1)) + 2 * Atn(1)
End Function
Function ArcTan(u As Double) As Double  'MODULE 1109
    ArcTan = Atn(u)
End Function
Function Ln(u As Double) As Double  'MODULE 1110
    Ln = Log(u)
End Function
Function Exp10(u As Double) As Double  'MODULE 1111
    '10^x Parameter: object.property = Ten(object.property)
    Ten = 10 ^ (u)
End Function
Function Sinh(u As Double) As Double  'MODULE 1112
    'Sinh Parameter: object.property = Sinh(object.property)
    Sinh = (Exp(u) - Exp(-u)) / 2
End Function
Function Cosh(u As Double) As Double  'MODULE 1113
    'Cosh Parameter: object.property = Cosh(object.property)
    Cosh = (Exp(u) + Exp(-u)) / 2
End Function
Function PlayWav(ByVal u As String)  'MODULE 1114
    soundfile1 = u
    sndPlaySound u, 1
End Function
Function Csc(u As Double) As Double  'MODULE 1122
    'Csc Parameter: object.property = Csc(object.property)
    Csc = 1 / Sin(u)
End Function
Function Sec(u As Double) As Double  'MODULE 1123
    'Sec Parameter: object.property = Sec(object.property)
    Sec = 1 / Cos(u)
End Function
Function Cot(u As Double) As Double  'MODULE 1124
    'Cot Parameter: object.property = Cot(object.property)
    Cot = 1 / Tan(u)
End Function
Function Csch(u As Double) As Double 'MODULE 1175
    'Csch Parameter: object.property = Csch(object.property)
    Csch = 1 / Sinh(u)
End Function
Function Sech(u As Double) As Double  'MODULE 1176
    'Sech Paramete: object.property = Sech(object.property)
    Sech = 1 / Cosh(u)
End Function
Function Tanh(u As Double)
    Tanh = Sinh(u) / Cosh(u)
End Function
Function Log10(u As Double)
    Log10 = Ln(u) / Ln(10)
End Function
Function LogX(u As Double, V As Double)
    LogBase = Ln(u) / Ln(V)
End Function
Function ArcCsc(u As Double) As Double
    ArcCsc = 1 / ArcSin(u)
End Function
Function ArcSec(u As Double) As Double
    ArcSec = 1 / ArcCos(u)
End Function
Function ArcCot(u As Double) As Double
    ArcCot = 1 / ArcTan(u) 'Or ArcCot = 1 / Atn(u) < = The standard math function.
End Function
Function Deg2Rad(Degrees)
    Deg2Rad = Degrees * (pi / 180)
End Function
Function Rad2Deg(Radians)
    Rad2Deg = Radians * (180 / pi)
End Function
Sub GradientColor(Frm As Form, colStart As Long, colEnd As Long) 'Create a gradient back color
   
   'Example: Me.BackColor = GradientColor(Me, RGB(0,0,0), RGB(255,255,255))
   
   Dim Red As Single
   Dim Green As Single
   Dim Blue As Single
   Dim redStep As Single
   Dim greenStep As Single
   Dim blueStep As Single
   Dim StepInterval As Single
   Dim x As Long
   Dim Ret As Long
   Dim OldMode As Long
   Dim FillArea As RECT
   Dim rTop As Single
   Dim hBrush As Long
   OldMode = Frm.ScaleMode
   Frm.ScaleMode = vbPixels
   StepInterval = Frm.ScaleHeight / 64
   Blue = (colStart \ &H10000) And &HFF
   blueStep = (Blue - ((colEnd \ &H10000) And &HFF)) / 64
   Green = (colStart \ &H100) And &HFF
   greenStep = (Green - ((colEnd \ &H100) And &HFF)) / 64
   Red = (colStart And &HFF)
   redStep = (Red - (colEnd And &HFF)) / 64
   rTop = 0
   FillArea.Left = 0
   FillArea.Right = Frm.ScaleWidth
   FillArea.Top = 0
   FillArea.Bottom = StepInterval
   For x = 1 To 64
       hBrush = CreateSolidBrush(RGB(Red, Green, Blue))
       Ret = FillRect(Frm.HDC, FillArea, hBrush)
       Ret = DeleteObject(hBrush)
       Red = Red - redStep
       Green = Green - greenStep
       Blue = Blue - blueStep
       rTop = rTop + StepInterval
       FillArea.Top = rTop
       FillArea.Bottom = rTop + StepInterval
   Next
End Sub
Sub ScreenResolution(x As Single, y As Single) 'Change the Screen Resolution
   Dim u As Boolean
   Dim i&
   i = 0
   Do
       u = EnumDisplaySettings(0&, i&, DevChg)
       i = i + 1
   Loop Until (u = False)

   Dim V&
   DevChg.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
   DevChg.dmPelsWidth = x
   DevChg.dmPelsHeight = y
   V = ChangeDisplaySettings(DevChg, 0)
End Sub



Sub Drag(ByVal uObject As Object)  'Drag an object
    ReleaseCapture
    SendMessage uObject.hWnd, WM_NCLBUTTONDOWN, LP_HT_CAPTION, 0
End Sub



Sub EllipseObject(ByVal Width As Long, ByVal Height As Long, ByVal hWnd As Long)  'Creates an ellptic object.
    Dim Right As Long
    Dim Bottom As Long
    Dim Rgn As Long
    Right = Width / Screen.TwipsPerPixelX
    Bottom = Height / Screen.TwipsPerPixelY
    Rgn = CreateEllipticRgn(0, 0, Right, Bottom)
    SetWindowRgn hWnd, Rgn, True
End Sub
Sub RoundRectObject(ByVal u As Object, ByVal V As Long, ByVal W As Long)  'Creates a rounded rectangular object.
    Dim lRight As Long
    Dim lBottom As Long
    Dim hRgn As Long
With u
lRight = .Width / Screen.TwipsPerPixelX
lBottom = .Height / Screen.TwipsPerPixelY
hRgn = CreateRoundRectRgn(0, 0, lRight, lBottom, V, W)
SetWindowRgn .hWnd, hRgn, True
End With
End Sub
Sub HideStartMenu()  'Hide the start menu button.
    tWnd = FindWindow("Shell_traywnd", vbNullString)
    bWnd = GetWindow(tWnd, GW_CHILD)
    Do
        GetClassName bWnd, sSave, 250
        If LCase(Left$(sSave, 6)) = "button" Then Exit Do
        bWnd = GetWindow(bWnd, GW_HWNDNEXT)
    Loop
    SetWindowPos bWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW
End Sub
Sub ShowStartMenu()  'Shows the original start menu.
    SetWindowPos bWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW
End Sub
Function ShowColor(ByVal uObject As Object)  'Creates the color selection dialog.
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    cc.lStructSize = Len(cc)
    cc.hwndOwner = uObject.hWnd
    cc.hInstance = App.hInstance
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 0
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function
Function ShowOpen(ByVal uObject As Object, ByVal uCaption As String) As String  'Creates the open dialog
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = uObject.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = uCaption
    OFName.flags = 0

    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function
Function ShowFont(ByVal uObject As Object) As String  'Shows the font dialog.
    Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim fontname As String, Retval As Long
    lfont.lfHeight = 0
    lfont.lfWidth = 0
    lfont.lfEscapement = 0
    lfont.lfOrientation = 0
    lfont.lfWeight = FW_NORMAL
    lfont.lfCharSet = DEFAULT_CHARSET
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    lfont.lfQuality = DEFAULT_QUALITY
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    lfont.lfFaceName = "Times New Roman" & vbNullChar
  
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, lfont, Len(lfont)

    cf.lStructSize = Len(cf)
    cf.hwndOwner = uObject.hWnd
    cf.HDC = Printer.HDC
    cf.lpLogFont = pMem
    cf.iPointSize = 120
    cf.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    cf.rgbColors = RGB(0, 0, 0)
    cf.nFontType = REGULAR_FONTTYPE
    cf.nSizeMin = 10
    cf.nSizeMax = 72

    Retval = CHOOSEFONT(cf)
    If Retval <> 0 Then
        CopyMemory lfont, ByVal pMem, Len(lfont)
        ShowFont = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        Debug.Print
    End If
    Retval = GlobalUnlock(hMem)
    Retval = GlobalFree(hMem)
    cf.lpTemplateName = uObject.Font
End Function
Function SaveDialog(ByVal uObject As Object, ByVal uCaption As String) As String  'Creates the save as Dialog.
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = uObject.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = uCaption
    OFName.flags = 0
    If GetSaveFileName(OFName) Then
        SaveDialog = Trim$(OFName.lpstrFile)
    Else
        SaveDialog = ""
    End If
End Function
Function ShowPageSetupDlg(ByVal uObject As Object) As String  'Page Setup Dialog
    Dim m_PSD As PAGESETUPDLG
    m_PSD.lStructSize = Len(m_PSD)
    m_PSD.hwndOwner = uObject.hWnd
    m_PSD.hInstance = App.hInstance
    m_PSD.flags = 0
    If PAGESETUPDLG(m_PSD) Then
        ShowPageSetupDlg = 0
    Else
        ShowPageSetupDlg = -1
    End If
End Function
Sub ShowPrinter(frmOwner As Form, Optional PrintFlags As Long)  'Print Dialog
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DEVMODE As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String
    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = frmOwner.hWnd

    PrintDlg.flags = PrintFlags
    On Error Resume Next
    DEVMODE.dmDeviceName = Printer.DeviceName
    DEVMODE.dmSize = Len(DEVMODE)
    DEVMODE.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DEVMODE.dmPaperWidth = Printer.Width
    DEVMODE.dmOrientation = Printer.Orientation
    DEVMODE.dmPaperSize = Printer.PaperSize
    DEVMODE.dmDuplex = Printer.Duplex
    On Error GoTo 0
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DEVMODE))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DEVMODE, Len(DEVMODE)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    If PrintDialog(PrintDlg) <> 0 Then

        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames
        
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DEVMODE, ByVal lpDevMode, Len(DEVMODE)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DEVMODE.dmDeviceName, InStr(DEVMODE.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                End If
            Next
        End If

        On Error Resume Next
        Printer.Copies = DEVMODE.dmCopies
        Printer.Duplex = DEVMODE.dmDuplex
        Printer.Orientation = DEVMODE.dmOrientation
        Printer.PaperSize = DEVMODE.dmPaperSize
        Printer.PrintQuality = DEVMODE.dmPrintQuality
        Printer.ColorMode = DEVMODE.dmColor
        Printer.PaperBin = DEVMODE.dmDefaultSource
        On Error GoTo 0
    End If
End Sub
Sub FontSelect(ByVal uObject As Object, ByVal vObject As Object)  'Create the font dialog to select the font of an object.
    On Error Resume Next
    uObject.Font = ShowFont(vObject)
End Sub
Sub FileOpen(ByVal uObject As Object, ByVal uTitle As String)
    Dim sFile As String
    sFile = ShowOpen(uObject, uTitle)
    If sFile <> "" Then
        MsgBox "You chose this file: " + sFile  'RichTextBox1.LoadFile OFName.lpstrFile, 1  is used to Open File. Do not erase. This is the example how to open files"
    Else
        MsgBox "No file has been selected"
    End If
End Sub
Sub FileSaveAs(ByVal uObject As Object, ByVal uTitle As String)  'Save As Dialog
    Dim sFile As String
    sFile = SaveDialog(uObject, uTitle)
    If sFile <> "" Then
        MsgBox "Selected File:" + sFile  'RichTextBox1.SaveFile OFName.lpstrFile, 1  is used to Save File. Do not erase."
    Else
        MsgBox "No color has been entered."
    End If
End Sub
Sub ColorSelect(ByVal uObject As Object)  'Color selection dialog
    Dim NewColor As Long
    NewColor = ShowColor(uObject)
    If NewColor <> -1 Then
        uObject.BackColor = NewColor
    Else
        MsgBox "No color has been selected.", vbInformation
    End If
End Sub
Sub FlatScrollBarV(ByVal uObject As Object)  'Vertical ScrollBar
    Dim SI As SCROLLINFO
    InitializeFlatSB uObject.hWnd
    FlatSB_SetScrollProp uObject.hWnd, WSB_PROP_VSTYLE, FSB_ENCARTA_MODE, False
    FlatSB_EnableScrollBar uObject.hWnd, SB_VERT, ESB_DISABLE_UP
    FlatSB_SetScrollRange uObject.hWnd, SB_VERT, 20, 80, False
    FlatSB_SetScrollPos uObject.hWnd, SB_VERT, 60, False
    FlatSB_ShowScrollBar uObject.hWnd, SB_HORZ, False
    SI.cbSize = Len(SI)
    SI.fMask = SIF_ALL
    FlatSB_GetScrollInfo uObject.hWnd, SB_VERT, SI
    SI.nPos = SI.nPos - 10
    FlatSB_SetScrollInfo uObject.hWnd, SB_VERT, SI, True
    Dim RetMin As Long, RetMax As Long
    FlatSB_GetScrollRange uObject.hWnd, SB_VERT, RetMin, RetMax
    uObject.AutoRedraw = True
End Sub
Sub FlatScrollBarH(ByVal uObject As Object)  'Horizontal ScrollBar
    Dim SI As SCROLLINFO
    InitializeFlatSB uObject.hWnd
    FlatSB_SetScrollProp uObject.hWnd, WSB_PROP_VSTYLE, FSB_ENCARTA_MODE, False
    FlatSB_EnableScrollBar uObject.hWnd, SB_HORZ, ESB_DISABLE_UP
    FlatSB_SetScrollRange uObject.hWnd, SB_HORZ, 20, 80, False
    FlatSB_SetScrollPos uObject.hWnd, SB_HORZ, 60, False
    FlatSB_ShowScrollBar uObject.hWnd, SB_HORZ, False
    SI.cbSize = Len(SI)
    SI.fMask = SIF_ALL
    FlatSB_GetScrollInfo uObject.hWnd, SB_HORZ, SI
    SI.nPos = SI.nPos - 10
    FlatSB_SetScrollInfo uObject.hWnd, SB_HORZ, SI, True
    Dim RetMin As Long, RetMax As Long
    FlatSB_GetScrollRange uObject.hWnd, SB_HORZ, RetMin, RetMax
    uObject.AutoRedraw = True
End Sub
Sub RemoveScrollBar(ByVal uObject As Object)  'Destroy the scrollbar
    UninitializeFlatSB uObject.hWnd
End Sub
Function NewStartMenu(ByVal uTitle As String) As String  'Creates a new start button. Under 5 characters.
    Dim r As RECT
    Dim aWnd As Long, bWnd As Long, ncWnd As Long
    aWnd = FindWindow("Shell_TrayWnd", vbNullString)
    bWnd = FindWindowEx(tWnd, ByVal 0&, "BUTTON", vbNullString)
    GetWindowRect bWnd, r
    ncWnd = CreateWindowEx(ByVal 0&, "BUTTON", uTitle, WS_CHILD, 0, 0, r.Right - r.Left, r.Bottom - r.Top, aWnd, ByVal 0&, App.hInstance, ByVal 0&)
    ShowWindow ncWnd, SW_NORMAL
    ShowWindow bWnd, SW_HIDE
End Function
Sub EmptyRecycleBin(ByVal uObject As Object)  'Empty the recycle bin
    Dim RBinInfo As SHQUERYRBINFO, msg As VbMsgBoxResult
    RBinInfo.cbSize = Len(RBinInfo)
    SHQueryRecycleBin vbNullString, RBinInfo
    If (RBinInfo.i64Size.lowpart And &H80000000) = &H80000000 Or RBinInfo.i64Size.highpart > 0 Then
        msg = MsgBox("Your Recycle Bin consumes over 2 GB right now!" + vbCrLf + "Do you want to empty it?", vbYesNo + vbQuestion)
    Else
        msg = MsgBox("Your Recycle Bin consumes" + str$((RBinInfo.i64Size.lowpart) / 1024) + " KB right now." + vbCrLf + "Do you want to empty it?", vbYesNo + vbQuestion)
    End If
    If msg = vbYes Then
        SHEmptyRecycleBin uObject.hWnd, vbNullString, 0
        SHUpdateRecycleBinIcon
    End If
End Sub
Sub RotateText(ByVal uForm As Form, ByVal uDegrees As Long, ByVal uString As String)  'Rotate the text of an object
    Dim RotateMe As LOGFONT
    Dim uSize As Long
    uForm.AutoRedraw = True
    uSize = 16
    RotateMe.lfEscapement = uDegrees * 10
    RotateMe.lfHeight = (uSize * -20) / Screen.TwipsPerPixelY
    rFont = CreateFontIndirect(RotateMe)
    Curent = SelectObject(uForm.HDC, rFont)
    uForm.CurrentX = 500
    uForm.CurrentY = 200
    uForm.Print uString
End Sub
Function FileProperties(ByVal fileName As String, ByVal hWnd As Long)  'Show the file properties of a sepcified file
    Dim SEI As SHELLEXECUTEINFO
    Dim r As Long
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = hWnd
        .lpVerb = "Properties"
        .lpFile = fileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    r = ShellExecuteEx(SEI)
End Function
Function RunDialog(ByVal uCaption As String, ByVal uPrompt As String, ByVal hWnd As Long)  'Shows the run Dialog
    Dim sTitle As String, sPrompt As String
    sTitle = uCaption
    sPrompt = uPrompt
    If IsWinNT Then
        SHRunDialog hWnd, 0, 0, StrConv(sTitle, vbUnicode), StrConv(sPrompt, vbUnicode), &H2
    Else
        SHRunDialog hWnd, 0, 0, sTitle, sPrompt, &H2
    End If
End Function
Function IsWinNT() As Boolean
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    Ret& = GetVersionEx(OSInfo)
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function
Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function

Function HideProgOnList()  'Hide program on CTRL ALT DEL List.
    Dim pid As Long, reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Function
Function ShowProgOnList()  ' Show program on CTRL ALT DEL List.
    Dim pid As Long, reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
End Function
Function ReverseString(ByVal RevStr As String) As String
    'Example:
    'strResult = ReverseString("String")
    'MsgBox strResult
    Dim DoReverse As Long
    ReverseString = ""
    For DoReverse = Len(RevStr) To 1 Step -1
    ReverseString = ReverseString & Mid$(RevStr, DoReverse, 1)
    Next
End Function
Function Bin(ByVal MyNum As Variant) As String
Dim LoopCounter As Integer
    If MyNum >= 2 ^ 31 Then
        Bin = "Value too Large"
        Exit Function
    End If
    Do
    If (MyNum And 2 ^ LoopCounter) = 2 ^ LoopCounter Then
        Bin = "1" & Bin
        Else
        Bin = "0" & Bin
    End If
    LoopCounter = LoopCounter + 1
    Loop Until 2 ^ LoopCounter > MyNum
End Function
Function TransparentObject(ByVal hWnd As Long)
'Example:
'Private Sub Form_Load() To make a flat-edged textbox...
'TransparentObject Text1.Hwnd
'End Sub
SetWindowLong hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME
End Function
Function Round(Value As Double, Digits As Integer) As Double
    Round = Int(Value * (10 ^ Digits) + 0.5) / (10 ^ Digits)
End Function
Function TerminateForms()
    Dim Form As Form
    For Each Form In Forms
    Unload Form
    Set Form = Nothing
    Next Form
End Function
Function ExitWindows(ByVal uFlags As Long)
   Call ExitWindowsEx(uFlags, 0)
End Function
Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   i = j
   End If
   GetActiveWindowTitle = GetWindowTitle(i)
End Function
Function HideTaskBar()
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd <> 0 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    End If
End Function
Function ShowTaskBar()
    If TaskBarhWnd <> 0 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End If
End Function
Function GetActiveWindow(ByVal ReturnParent As Boolean) As Long
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   End If
   GetActiveWindow = i
End Function
Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hWnd)
   S = Space(l + 1)
   GetWindowText hWnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function
Function TopMostForm(F As Form, Top As Boolean)
   If Top Then
      SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
   Else
      SetWindowPos F.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
   End If
End Function
Function Pause(ByVal Seconds As Single)
   Call Sleep(Int(Seconds * 1000#))
End Function
Function AboutBox(Frm As Form, Optional Copyright As Variant)
   If VarType(Copyright) = vbString Then
      Call ShellAbout(Frm.hWnd, App.ProductName, Copyright, Frm.Icon)
   Else
      Call ShellAbout(Frm.hWnd, App.ProductName, "", Frm.Icon)
   End If
End Function
Function KillObject(ByVal hWnd As Long)
    Dim temp
    temp = SendMessage(hWnd, WM_DESTROY, 0, 0)
End Function
Function Disable_CRTL_ALT_DEL()
    Dim Ret As Integer
    Dim pOld As Boolean
    Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Function
Function Enable_CRTL_ALT_DEL()
    Dim Ret As Integer
    Dim pOld As Boolean
    Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Function
Function DisableXButton(Frm As Form)
    Dim hSysMenu As Long
    Dim nCnt As Long
    Frm.Show
    hSysMenu = GetSystemMenu(Frm.hWnd, False)
    If hSysMenu Then
        nCnt = GetMenuItemCount(hSysMenu)
    If nCnt Then
        RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
        RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
        DrawMenuBar Frm.hWnd
    End If
End If
End Function
Function OpenCDDoor()
    Dim returnstring, retvalue
    retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Function
Function CD_Close()
    Dim CDValue As String
    CDValue$ = mciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Function
Function Hide_Clock()
    Dim shelltraywnd As Long
    Dim ShelltryWnd As Long, traynotifywnd As Long, trayclockwclass As Long
    shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
    traynotifywnd& = FindWindowEx(shelltraywnd&, 0&, "TrayNotifyWnd", vbNullString)
    trayclockwclass& = FindWindowEx(traynotifywnd&, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(trayclockwclass&, SW_HIDE)
End Function
Function Hide_Desktop()
    Dim Progman As Long, SHELLDLLDefView As Long, SysListView As Long
    Progman& = FindWindow("Progman", vbNullString)
    SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
    SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
    Call ShowWindow(SysListView&, SW_HIDE)
End Function
Function Hide_TaskBar()
    Dim shelltraywnd As Long
    shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(shelltraywnd&, SW_HIDE)
End Function
Function Hide_TrayIcons()
    Dim shelltraywnd As Long, traynotifywnd As Long
    shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
    traynotifywnd& = FindWindowEx(shelltraywnd&, 0&, "TrayNotifyWnd", vbNullString)
    Call ShowWindow(traynotifywnd&, SW_HIDE)
End Function
Function Show_Clock()
    Dim shelltraywnd As Long
    Dim ShelltryWnd As Long, traynotifywnd As Long, trayclockwclass As Long
    shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
    traynotifywnd& = FindWindowEx(shelltraywnd&, 0&, "TrayNotifyWnd", vbNullString)
    trayclockwclass& = FindWindowEx(traynotifywnd&, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(trayclockwclass&, SW_SHOW)
End Function
Function Show_Desktop()
    Dim Progman As Long, SHELLDLLDefView As Long, SysListView As Long
    Progman& = FindWindow("Progman", vbNullString)
    SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
    SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
    Call ShowWindow(SysListView&, SW_SHOW)
End Function
Function Show_TaskBar()
    Dim shelltraywnd As Long
    shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(shelltraywnd&, SW_SHOW)
End Function
Function Show_TrayIcons()
    Dim shelltraywnd As Long, traynotifywnd As Long
    shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
    traynotifywnd& = FindWindowEx(shelltraywnd&, 0&, "TrayNotifyWnd", vbNullString)
    Call ShowWindow(traynotifywnd&, SW_SHOW)
End Function
Public Function File_Copy(currentFilename As String, newFilename As String)
    Dim A%, Buffer%, temp$, fRead&, fSize&, B%
    A = FreeFile
    Buffer = 4048
    Open currentFilename For Binary Access Read As A
    B = FreeFile
    Open newFilename For Binary Access Write As B
    fSize = FileLen(currentFilename)
    While fRead < fSize
    DoEvents
    If Buffer > (fSize - fRead) Then Buffer = (fSize - fRead)
    temp = Space(Buffer)
    Get A, , temp
    Put B, , temp
    fRead = fRead + Buffer
    Wend
    Close B
    Close A
    File_Copy = 1
    Exit Function
End Function
Function StartScreenSaver(ByValfrm As Form)
    SendMessage Frm.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0
End Function
Function Reboot()
    ExitWindowsEx EWX_REBOOT, 0
End Function
Function ForceShutdown()
    ExitWindowsEx EWX_FORCE, 0
End Function
Function Shutdown()
    ExitWindowsEx EWX_SHUTDOWN, 0
End Function
Function SaveText(File As String, vText As TextBox)
    Open File For Output As #1
    Print #1, vText
    Close #1
End Function
Function MouseHide()
   ShowCursor False
End Function
Sub MouseShow()
    ShowCursor True
End Sub
Function SetKeyState(ByVal intKey As Integer, ByVal fTurnOn As Boolean)
    Dim abytBuffer(0 To 255) As Byte
    GetKeyboardState abytBuffer(0)
    abytBuffer(intKey) = CByte(Abs(fTurnOn))
    SetKeyboardState abytBuffer(0)
End Function
Function CapsLockON()
    SetKeyState vbKeyCapital, True
End Function
Function CapsLockOFF()
    SetKeyState vbKeyCapital, False
End Function
Function NumLockON()
    SetKeyState vbKeyNumlock, True
End Function
Function NumLockOFF()
    SetKeyState vbKeyNumlock, False
End Function
Function ScrollLockON()
    SetKeyState vbKeyScrollLock, True
End Function
Function ScrollLockOFF()
    SetKeyState vbKeyScrollLock, False
End Function
Function File_Open(ByVal File As String, ByVal vText As TextBox)
    Open File For Input As #1
    vText = Input(LOF(1), 1)
    Close #1
End Function
Function CheckError() As String
    CheckError = Space$(255)
    mciGetErrorString error, CheckError, Len(CheckError)
End Function
Function InitCD()
    error = mciSendString("open cdaudio alias cd wait shareable", 0, 0, 0)
End Function
Function PlayCD()
    error = mciSendString("play cd", 0, 0, 0)
End Function
Function StopCD()
    error = mciSendString("stop cd", 0, 0, 0)
End Function
Function CloseAll()
    error = mciSendString("close all", 0, 0, 0)
End Function
Function GetPositionMSF() As String
    error = mciSendString("set cd time format msf", 0, 0, 0)
    error = mciSendString("status cd position", TheStat, 128, 0)
    GetPositionMSF = TheStat
End Function
Function GetPositionMS() As String
    error = mciSendString("set cd time format ms", 0, 0, 0)
    error = mciSendString("status cd position", TheStat, 128, 0)
    GetPositionMS = TheStat
End Function
Function GetPositionTMSF() As String
    error = mciSendString("set cd time format tmsf", 0, 0, 0)
    error = mciSendString("status cd position", TheStat, 128, 0)
    GetPositionTMSF = TheStat
End Function
Function GetTimeFormat() As String
    error = mciSendString("status cd time format", TheStat, 128, 0)
    GetTimeFormat = TheStat
End Function
Function GetNumberTracks() As Long
    On Error Resume Next
    error = mciSendString("status cd number of tracks", TheStat, 128, 0)
    GetNumberTracks = CLng(Trim$(TheStat))
End Function
Function SetTrack(track As Long)
    error = mciSendString("play cd from " & track, 0, 0, 0)
End Function
Function PauseCD()
    error = mciSendString("pause cd", 0, 0, 0)
End Function
Function ResumeCD()
    error = mciSendString("play cd", 0, 0, 0)
End Function
Function FastForward(Seconds As Long)
    Dim pos As String * 128
    error = mciSendString("set cd time format ms", 0, 0, 0)
    error = mciSendString("status cd position", pos, 128, 0)
    pos = CLng(Val(pos))
    If ISPlaying = True Then
        error = mciSendString("play cd from " & pos + Seconds * 1000, 0, 0, 0)
    Else
        error = mciSendString("seek cd to " & pos + Seconds * 1000, 0, 0, 0)
    End If
End Function
Function FastRewind(Seconds As Long)
    Dim pos As String * 128
    error = mciSendString("set cd time format ms", 0, 0, 0)
    error = mciSendString("status cd position", pos, 128, 0)
    pos = CLng(Val(pos))
    If ISPlaying = True Then
        error = mciSendString("play cd from " & pos - Seconds * 1000, 0, 0, 0)
    Else
        error = mciSendString("seek cd to " & pos - Seconds * 1000, 0, 0, 0)
    End If
End Function
Function ISPlaying() As Boolean
    error = mciSendString("status cd mode", TheStat, 128, 0)
    If Left(TheStat, 7) = "playing" Then
        ISPlaying = True
    Else
        ISPlaying = False
    End If
End Function
Function TrackLengthMS(track As Long) As Long
    error = mciSendString("set cd time format ms", 0, 0, 0)
    error = mciSendString("status cd length track " & track, TheStat, 128, 0)
    TrackLengthMS = Val(TheStat)
End Function
Function CDLengthMS() As Long
    error = mciSendString("set cd time format ms", 0, 0, 0)
    error = mciSendString("status cd length", TheStat, 128, 0)
    CDLengthMS = Val(TheStat)
End Function
Function CDLengthMSF() As String
    error = mciSendString("set cd time format msf", 0, 0, 0)
    error = mciSendString("status cd length", TheStat, 128, 0)
    CDLengthMSF = TheStat
End Function
Function TrackLengthMSF(track As Long) As String
    error = mciSendString("set cd time format msf", 0, 0, 0)
    error = mciSendString("status cd length track " & track, TheStat, 128, 0)
    TrackLengthMSF = TheStat
End Function
Function SetAudioOff()
    error = mciSendString("set cd audio all off", 0, 0, 0)
End Function
Function SetAudioOn()
    error = mciSendString("set cd audio all on", 0, 0, 0)
End Function
Function PictureInvert(ByVal PictureObject As Object)
  On Error Resume Next
  PictureObject.PaintPicture PictureObject.Picture, 0, 0, , , , , , , vbDstInvert
End Function
Function PictureRotate(ByVal PictureSource As Object, ByRef PictureDestination As Object, ByVal RotateAngle As Single)
On Error Resume Next
  
  Dim A   As Single
  Dim n   As Long
  Dim r   As Long
  Dim c0  As Long
  Dim c1  As Long
  Dim c2  As Long
  Dim c3  As Long
  Dim c1x As Long
  Dim c1y As Long
  Dim c2x As Long
  Dim c2y As Long
  Dim p1x As Long
  Dim p1y As Long
  Dim p2x As Long
  Dim p2y As Long
  Dim pic1hDC As Long
  Dim pic2hDC As Long
  Dim Theta As Double
  Dim TheNumber As Single
    
 If RotateAngle > 360 Then
    MsgBox "Rotation angle can not be greater than 360", vbOKOnly + vbExclamation, "  Invalid Rotation Angle"
    Exit Function
    ElseIf RotateAngle < 0 Then
    MsgBox "Rotation angle can not be less than 0", vbOKOnly + vbExclamation, "  Invalid Rotation Angle"
    Exit Function
  End If
  TheNumber = 180 / RotateAngle
  Theta = pi / TheNumber
  Screen.MousePointer = vbHourglass
  PictureDestination.MousePointer = vbHourglass
  PictureSource.MousePointer = vbHourglass
  PictureDestination.Cls
  c1x = PictureSource.ScaleWidth \ 2
  c1y = PictureSource.ScaleHeight \ 2
  c2x = PictureDestination.ScaleWidth \ 2
  c2y = PictureDestination.ScaleHeight \ 2
  If c2x < c2y Then
    n = c2y
  Else
    n = c2x
  End If
  n = n - 1
  pic1hDC = PictureSource.HDC
  pic2hDC = PictureDestination.HDC
    For p2x = 0 To n
    For p2y = 0 To n
      If p2x = 0 Then
        A = pi / 2
      Else
        A = Atn(p2y / p2x)
      End If
      r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
      p1x = r * Cos(A + Theta)
      p1y = r * Sin(A + Theta)
      c0 = GetPixel(pic1hDC, c1x + p1x, c1y + p1y)
      c1 = GetPixel(pic1hDC, c1x - p1x, c1y - p1y)
      c2 = GetPixel(pic1hDC, c1x + p1y, c1y - p1x)
      c3 = GetPixel(pic1hDC, c1x - p1y, c1y + p1x)
      If c0 <> -1 Then
        SetPixel pic2hDC, c2x + p2x, c2y + p2y, c0&
      End If
      If c1 <> -1 Then
        SetPixel pic2hDC, c2x - p2x, c2y - p2y, c1&
      End If
      If c2 <> -1 Then
        SetPixel pic2hDC, c2x + p2y, c2y - p2x, c2&
      End If
      If c3 <> -1 Then
        SetPixel pic2hDC, c2x - p2y, c2y + p2x, c3&
      End If
    Next
  Next
  Screen.MousePointer = vbDefault
  PictureDestination.MousePointer = vbDefault
  PictureDestination.Refresh
End Function
Function AutoComplete(ByRef CboBox As ComboBox)
On Error Resume Next
  
  Dim ReturnValue As Long
  Dim lngPosition As Long

  With CboBox
    lngPosition = Len(.Text)
    If lngPosition <> 0 Then
      ReturnValue = SendMessage(.hWnd, &H14C, -1&, ByVal .Text)
      .ListIndex = ReturnValue
      .SelStart = lngPosition
      .SelLength = Len(.Text) - lngPosition
    End If
  End With

End Function
Function CenterWindow(Optional ByVal Handle As Long, Optional ByVal Title As String = "", Optional ByVal Classname As String = "") As Boolean
  
  Dim TheHandle As Long
  Dim ReturnValue As Long
  Dim W_Top As Integer
  Dim W_Left As Integer
  Dim W_Width As Integer
  Dim W_Height As Integer
  Dim rct As RECT
  
  If Classname = "" And Title = "" And Handle = 0 Then
    Exit Function
  ElseIf Handle <> 0 Then
    TheHandle = Handle
  ElseIf Classname = "" Then
    TheHandle = FindWindow(vbNullString, Title)
  Else
    TheHandle = FindWindow(Classname, vbNullString)
  End If
  
  If TheHandle = 0 Then
    Exit Function
  End If

  ReturnValue = GetWindowRect(TheHandle, rct)
  If ReturnValue = 0 Then
    Exit Function
  End If
  
  W_Top = rct.Top
  W_Left = rct.Left
  W_Height = rct.Bottom - rct.Top
  W_Width = rct.Right - rct.Left
  W_Top = ((Screen.Height / Screen.TwipsPerPixelY) - W_Height) / 2
  W_Left = ((Screen.Width / Screen.TwipsPerPixelX) - W_Width) / 2
  ShowWindow TheHandle, SW_SHOW
  ShowWindow TheHandle, SW_RESTORE
  MoveWindow TheHandle, W_Left, W_Top, W_Width, W_Height, True
  CenterWindow = True
  Exit Function

End Function
Function ConvertLong2Short(ByVal strFullPath As String) As String
On Error Resume Next

  strFullPath = Trim(strFullPath)
  If strFullPath = "" Then Exit Function
  strFullPath = strFullPath & Chr(0)

  If Dir(strFullPath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function

  ConvertLong2Short = String(MAX_PATH, Chr(0))
  If GetShortPathName(strFullPath, ConvertLong2Short, MAX_PATH) <> 0 Then
    ConvertLong2Short = UCase(Left(ConvertLong2Short, InStr(ConvertLong2Short, Chr(0)) - 1))
  Else
    ConvertLong2Short = ""
  End If
  
End Function
Function EndTheProgram()
On Error Resume Next
  
  Dim Form As Form
  
  ExitingProgram = True
  
  For Each Form In Forms
    Unload Form
    Set Form = Nothing
  Next

  
End Function
Function FileDelete(ByVal FilePath As String, Optional ByVal blnPromptUser As Boolean = False, Optional ByRef Return_ErrNum As Long, Optional ByRef Return_ErrDesc As String) As Boolean
  
  Dim lngErrNum  As Long
  Dim strErrDesc As String

  Return_ErrNum = 0
  Return_ErrDesc = ""

  FilePath = Trim(FilePath)
  If FilePath = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "No file specified to delete"
    Exit Function
  ElseIf FileExists(FilePath) = False Then
    FileDelete = True
    Exit Function
  End If

  If blnPromptUser = True Then
    If MsgBox(FilePath & Chr(13) & Chr(13) & "Are you sure you want to delete this file?", vbYesNo + vbExclamation, "  Confirm File Delete") <> vbYes Then
      FileDelete = True
      Exit Function
    End If
  End If

  SetAttr FilePath, vbNormal
  Kill FilePath

  FileDelete = True
  Exit Function
  
End Function
Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function
Function FileExists(ByVal strFilePath As String) As Boolean
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function
Function FormPosition(ByVal FormHandle As Long, ByVal MakeTopMost As Boolean) As Boolean
  
  Dim ReturnValue As Long
  Dim Flag As Long
  
  If MakeTopMost = True Then
    Flag = HWND_TOPMOST
  Else
    Flag = HWND_NOTOPMOST
  End If
  
  ReturnValue = SetWindowPos(FormHandle, Flag, 0, 0, 0, 0, SWP_FLAGS)
  If ReturnValue <> 0 Then
    FormPosition = True
  End If
  
  Exit Function
  
End Function
Function GetAppVersion() As String
On Error Resume Next
  
  GetAppVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
  
End Function
Function GetWinPath() As String
On Error Resume Next

  Dim strFolder As String
  Dim lngResult As Long
  
  strFolder = String(MAX_PATH, 0)
  lngResult = GetWindowsDirectory(strFolder, MAX_PATH)
  If lngResult <> 0 Then
    GetWinPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
  Else
    GetWinPath = ""
  End If
  
End Function
Function GetWinSysPath() As String
On Error Resume Next

  Dim strFolder As String
  Dim lngResult As Long
  
  strFolder = String(MAX_PATH, 0)
  lngResult = GetSystemDirectory(strFolder, MAX_PATH)
  
  If lngResult <> 0 Then
    GetWinSysPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
  Else
    GetWinSysPath = ""
  End If
  
End Function
Function GetWinTemp() As String
On Error Resume Next
  
  Dim strFolder As String
  Dim lngResult As Long
  
  strFolder = String(MAX_PATH, 0)
  lngResult = GetTempPath(MAX_PATH, strFolder)
  
  If lngResult <> 0 Then
    GetWinTemp = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
  Else
    GetWinTemp = ""
  End If
  
End Function
Function GetX() As Long
On Error Resume Next
  
  Dim MyPoint As POINTAPI
  Dim ReturnValue As Long
  
  ReturnValue = GetCursorPos(MyPoint)
  If ReturnValue <> 0 Then
    GetX = MyPoint.x
  Else
    GetX = -1
  End If
  
End Function
Function GetY() As Long
On Error Resume Next
  
  Dim MyPoint As POINTAPI
  Dim ReturnValue As Long
  
  ReturnValue = GetCursorPos(MyPoint)
  If ReturnValue <> 0 Then
    GetY = MyPoint.y
  Else
    GetY = -1
  End If
  
End Function
Function INI_Read(ByVal SectionName As String, ByVal KeyName As String, ByVal INIPath As String, Optional ByVal DefaultValue As String = "") As String
On Error Resume Next
  
  Dim lngLength As Long
  
  INI_Read = String(MAX_PATH, Chr(0))
  lngLength = GetPrivateProfileString(SectionName & Chr(0), KeyName & Chr(0), DefaultValue & Chr(0), INI_Read, Len(INI_Read), INIPath & Chr(0))
  INI_Read = Left(INI_Read, lngLength)
  
End Function
Function INI_Write(ByVal SectionName As String, ByVal KeyName As String, ByVal Value As String, ByVal INIPath As String, Optional ByVal blnDeleteKeyIfBlank As Boolean = False) As Boolean
On Error Resume Next
  
  If blnDeleteKeyIfBlank = True Then
    If SectionName = "" Then SectionName = vbNullString Else SectionName = SectionName & Chr(0)
    If KeyName = "" Then KeyName = vbNullString Else KeyName = KeyName & Chr(0)
    If Value = "" Then Value = vbNullString Else Value = Value & Chr(0)
  Else
    SectionName = SectionName & Chr(0)
    KeyName = KeyName & Chr(0)
    Value = Value & Chr(0)
  End If
  
  If WritePrivateProfileString(SectionName, KeyName, Value, INIPath & Chr(0)) <> 0 Then INI_Write = True
  
End Function
Function RenameFile(ByVal strOldFilePath As String, ByVal strNewFileName As String, Optional ByRef Return_ErrNum As Long, Optional ByRef Return_ErrDesc As String) As Boolean
On Error Resume Next
  
  Dim strPath     As String
  Dim strFileName As String
  Dim strFileExt  As String
  
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  strOldFilePath = Trim(strOldFilePath)
  strNewFileName = Trim(strNewFileName)
  If strOldFilePath = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "No file specified to rename"
    Exit Function
  ElseIf strNewFileName = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "No file specified to rename the file to"
    Exit Function
  ElseIf FileExists(strOldFilePath) = False Then
    Return_ErrNum = -1
    Return_ErrDesc = "The file specified to rename does not exist"
    Exit Function
  ElseIf FileExists(strNewFileName) = True Then
    Return_ErrNum = -1
    Return_ErrDesc = "The file name specified to rename to already exists"
    Exit Function
  ElseIf InStr(strNewFileName, "\") <> 0 Or _
         InStr(strNewFileName, "/") <> 0 Or _
         InStr(strNewFileName, ":") <> 0 Or _
         InStr(strNewFileName, "*") <> 0 Or _
         InStr(strNewFileName, "?") <> 0 Or _
         InStr(strNewFileName, Chr(34)) <> 0 Or _
         InStr(strNewFileName, "<") <> 0 Or _
         InStr(strNewFileName, ">") <> 0 Or _
         InStr(strNewFileName, "|") <> 0 Then
    Return_ErrNum = -1
    Return_ErrDesc = "New file name contains one or more of the followin invalid characters:  \ / : * ? " & Chr(34) & " < > |"
    Exit Function
  ElseIf FileInUse(strOldFilePath) = True Then
    Return_ErrNum = -1
    Return_ErrDesc = "The specified file to rename is currently in use by another process so it can not be renamed"
    Exit Function
  End If
  
  strOldFilePath = ConvertLong2Short(strOldFilePath)
  If strOldFilePath = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "An error occured while changing the path to a 16-bit path."
    Exit Function
  End If

  If GetFileNameAndExt(strOldFilePath, strFileName, strFileExt) = False Then
    Return_ErrNum = -1
    Return_ErrDesc = "An error occured while trying to get the specified file's path"
    Exit Function
  Else
    strPath = Left(strOldFilePath, Len(strOldFilePath) - Len(strFileName) - Len(".") - Len(strFileExt))
  End If
  
  If MoveFile(strOldFilePath & Chr(0), strPath & strNewFileName & Chr(0)) = 0 Then
    Return_ErrNum = Err.LastDllError
    Return_ErrDesc = "An error occured while trying to rename the specified file"
  Else
    RenameFile = True
  End If
  
End Function
Function GetFileNameAndExt(ByVal strFilePath As String, ByRef Return_FileName As String, ByRef Return_FileExtention As String) As Boolean
  
  Dim StringSoFar As String
  Dim CharLeft    As String
  Dim CharRight   As String
  Dim lngCounter  As Long
  Dim blnFoundExt As Boolean

  Return_FileName = ""
  Return_FileExtention = ""
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If InStr(strFilePath, "\") = 0 And InStr(strFilePath, "/") = 0 And InStr(strFilePath, ".") = 0 Then
    Return_FileName = strFilePath
    Exit Function
  End If

  For lngCounter = 1 To Len(strFilePath)
    CharRight = Right(strFilePath, lngCounter)
    CharLeft = Left(CharRight, 1)
    If CharLeft = "." And blnFoundExt = False Then
      blnFoundExt = True
      Return_FileExtention = StringSoFar
      StringSoFar = ""
    ElseIf CharLeft = "\" Or CharLeft = "/" Then
      Return_FileName = StringSoFar
      GetFileNameAndExt = True
      Exit Function
    Else
      StringSoFar = CharLeft & StringSoFar
    End If
  Next
  
  Return_FileName = StringSoFar
  GetFileNameAndExt = True
  
End Function
Function ProgressBar(ByRef ThePictureBox As PictureBox, ByVal Min As Long, ByVal Max As Long, ByVal Value As Long, Optional ByVal ShowProgressCaption As Boolean = False, Optional ByVal ForeColor As Long = 16777215, Optional ByVal BackColor As Long = 16711680, Optional ByVal FillColor As Long = vbButtonFace, Optional ByVal Alignment As AlignmentConstants = vbCenter, Optional ByVal ByPassChecks As Boolean = False)
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

Function OpenProgram(ByVal FullProgramPath As String, Optional ByVal ProgramPriority As Priorities = p_Normal, Optional ByVal WindowState As WindowStates = SW_NORMAL, Optional ByVal DOS_FullScreen As Boolean = True, Optional ByRef Return_ProcessHandle As Long, Optional ByRef Return_ThreadHandle As Long, Optional ByRef Return_ProcessID As Long, Optional ByVal Return_ThreadID As Long) As Boolean

  Dim ReturnValue As Long
  Dim flags       As Long
  Dim StartInfo   As STARTUPINFO
  Dim ProcInfo    As PROCESS_INFORMATION
  
  If DOS_FullScreen = True Then
    flags = STARTF_RUNFULLSCREEN
  End If
  flags = flags Or STARTF_USESHOWWINDOW
  
  With StartInfo
    .cb = Len(StartInfo)
    .dwFlags = flags
    .wShowWindow = WindowState
  End With
  
  ReturnValue = CreateProcess(vbNullString, FullProgramPath, 0, 0, 0, ProgramPriority, ByVal 0&, StripPath(FullProgramPath), StartInfo, ProcInfo)
  If ReturnValue = 0 Then
    OpenProgram = False
  Else
    OpenProgram = True
    With ProcInfo
      Return_ProcessHandle = .hProcess
      Return_ProcessID = .dwProcessId
      Return_ThreadHandle = .hThread
      Return_ThreadID = .dwThreadId
    End With
  End If
  
  Exit Function
  
End Function
Function CloseProgram(Optional ByVal ProcessHandle As Long, Optional ByVal ThreadHandle As Long, Optional CloseThread As Boolean = True) As Boolean
  
  Dim ReturnValue As Long
  Dim ExitCode    As Long

  If CloseThread = True Then
    ReturnValue = GetExitCodeThread(ThreadHandle, ExitCode)
    If ReturnValue <> 0 Then
  '    TerminateProcess ProcessHandle, ExitCode
       TerminateThread ThreadHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
    
  Else
    ReturnValue = GetExitCodeProcess(ProcessHandle, ExitCode)
    If ReturnValue <> 0 Then
   '   TerminateThread ThreadHandle, ExitCode
       TerminateProcess ProcessHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
  End If
  Exit Function
  
End Function
Function StripPath(ByVal FullPath As String) As String
  If InStr(FullPath, "\") = 0 Then
  StripPath = FullPath
  Exit Function
  End If
  StripPath = Left(FullPath, InStrRev(FullPath, "\"))
End Function
