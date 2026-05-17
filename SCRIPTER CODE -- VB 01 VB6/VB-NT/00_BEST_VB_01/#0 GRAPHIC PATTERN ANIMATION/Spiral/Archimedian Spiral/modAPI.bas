Attribute VB_Name = "modAPI"
''''  Windows32  Platform Declares
''''  Contains Declares for Methods, Types and Constants from the Windows API
''''  Contains Specific Declarations for NT, 98, 2000 and Millenium

Option Explicit

'' ****** CONSTANTS ******


''' Point structures for scale widths and heights

Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_CLASSDC = &H40
Public Const CS_DBLCLKS = &H8
Public Const CS_HREDRAW = &H2
Public Const CS_INSERTCHAR = &H2000
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_NOCLOSE = &H200
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOMOVECARET = &H4000
Public Const CS_OWNDC = &H20
Public Const CS_PARENTDC = &H80
Public Const CS_PUBLICCLASS = &H4000
Public Const CS_SAVEBITS = &H800
Public Const CS_VREDRAW = &H1
Public Const CT_CTYPE1 = &H1
Public Const CT_CTYPE2 = &H2
Public Const CT_CTYPE3 = &H4
Public Const CTLCOLOR_BTN = 3
Public Const CTLCOLOR_DLG = 4
Public Const CTLCOLOR_EDIT = 1
Public Const CTLCOLOR_LISTBOX = 2
Public Const CTLCOLOR_MAX = 8
Public Const CTLCOLOR_MSGBOX = 0&
Public Const CTLCOLOR_SCROLLBAR = 5
Public Const CTLCOLOR_STATIC = 6

''' Window Creation

'' Window Styles 1

Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000

Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000

Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME

'' Window Styles 2

Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED + _
                 WS_CAPTION + _
                 WS_SYSMENU + _
                 WS_THICKFRAME + _
                 WS_MINIMIZEBOX + _
                 WS_MAXIMIZEBOX)

Public Const WS_POPUPWINDOW = (WS_POPUP + _
                 WS_BORDER + _
                 WS_SYSMENU)

Public Const WS_CHILDWINDOW = (WS_CHILD)

Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
 '
 ' Extended Window Styles
 '
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&

Public Const WS_EX_MDICHILD = &H40&
Public Const WS_EX_TOOLWINDOW = &H80&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_CONTEXTHELP = &H400&

Public Const WS_EX_RIGHT = &H1000&
Public Const WS_EX_LEFT = &H0&
Public Const WS_EX_RTLREADING = &H2000&
Public Const WS_EX_LTRREADING = &H0&
Public Const WS_EX_LEFTSCROLLBAR = &H4000&
Public Const WS_EX_RIGHTSCROLLBAR = &H0&

Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000

Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE + WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE + WS_EX_TOOLWINDOW + WS_EX_TOPMOST)


' Windows 5.0& (2000/Millenium)

Public Const WS_EX_LAYERED = &H80000

Public Const WS_EX_NOINHERITLAYOUT = &H100000     ' Disable inheritence of mirroring by children
Public Const WS_EX_LAYOUTRTL = &H400000           ' Right to left mirroring

' Windows NT 5.0& (Windows 2000) only

Public Const WS_EX_NOACTIVATE = &H8000000


' Windows Virtual Key Code Constants

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4           ' NOT contiguous with L & RBUTTON ''

Public Const VK_BACK = &H8
Public Const VK_TAB = &H9

Public Const VK_SEMICOLON = &H27

Public Const VK_QUOTE = &H28
Public Const VK_BACK_QUOTE = &H29

Public Const VK_SLASH = &H35
Public Const VK_BACK_SLASH = &H2B

Public Const VK_PERIOD = &H34
Public Const VK_COMMA = &H33

Public Const VK_LEFT_BRACE = &H1A
Public Const VK_RIGHT_BRACE = &H1B

Public Const VK_DASH = &HC
Public Const VK_EQUAL = &HD

Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD

Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14

Public Const VK_KANA = &H15
Public Const VK_HANGEUL = &H15        ' old name - should be here for compatibility ''
Public Const VK_HANGUL = &H15
Public Const VK_JUNJA = &H17
Public Const VK_FINAL = &H18
Public Const VK_HANJA = &H19
Public Const VK_KANJI = &H19

Public Const VK_ESCAPE = &H1B

Public Const VK_CONVERT = &H1C
Public Const VK_NONCONVERT = &H1D
Public Const VK_ACCEPT = &H1E
Public Const VK_MODECHANGE = &H1F

Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F

' VK_0& thru VK_9 are the same as ASCII '0&' thru '9' (=&H30 - =&H39) ''
' VK_A thru VK_Z are the same as ASCII 'A' thru 'Z' (=&H41 - =&H5A) ''

Public Const VK_LWIN = &H5B
Public Const VK_RWIN = &H5C
Public Const VK_APPS = &H5D

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87

Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91

'
' * VK_L* & VK_R* - left and right Alt, Ctrl and Shift virtual keys.
' * Used only as parameters to GetAsyncKeyState() and GetKeyState().
' * No other API or message will distinguish left and right keys in this way.
 ''
Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5

'' if (WINVER >= =&H0400)
Public Const VK_PROCESSKEY = &HE5
'' end if ' WINVER >= =&H0400 ''

Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE


'' System Command Constants

Public Const SC_ARRANGE = &HF110&
Public Const SC_CLOSE = &HF060&
Public Const SC_HOTKEY = &HF150&
Public Const SC_HSCROLL = &HF080&
Public Const SC_KEYMENU = &HF100&
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_MINIMIZE = &HF020&
Public Const SC_MOUSEMENU = &HF090&
Public Const SC_MOVE = &HF010&
Public Const SC_NEXTWINDOW = &HF040&
Public Const SC_PREVWINDOW = &HF050&
Public Const SC_RESTORE = &HF120&
Public Const SC_SCREENSAVE = &HF140&
Public Const SC_SIZEAPI = &HF000&
Public Const SC_TASKLIST = &HF130&
Public Const SC_VSCROLL = &HF070&

Public Const SC_ICON = SC_MINIMIZE
Public Const SC_ZOOM = SC_MAXIMIZE


'' Global Memory Constants

Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GMEM_LOCKCOUNT = &HFF
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40

Public Const GMEM_LOWER = GMEM_NOT_BANKED

'' Platform Version Constants

Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32s = 0&

''' Drawing and Device Context

'' Owner-drawn actions

Public Const ODA_DRAWENTIRE = &H1
Public Const ODA_FOCUS = &H4
Public Const ODA_SELECT = &H2

'' Owner-drawn states

Public Const ODS_CHECKED = &H8
Public Const ODS_DISABLED = &H4
Public Const ODS_FOCUS = &H10
Public Const ODS_GRAYED = &H2
Public Const ODS_SELECTED = &H1
Public Const ODS_DEFAULT = &H20
Public Const ODS_HOTLIGHT = &H40
Public Const ODS_NOACCEL = &H100

'' Owner-drawn types

Public Const ODT_BUTTON = 4
Public Const ODT_COMBOBOX = 3
Public Const ODT_LISTBOX = 2
Public Const ODT_MENU = 1

'' DrawText Constants

Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_CHARSTREAM = 4
Public Const DT_DISPFILE = 6
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0&
Public Const DT_RASCAMERA = 3
Public Const DT_RASDISPLAY = 1
Public Const DT_RASPRINTER = 2
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

'' 3D Border Styles

Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2

'' More 3D Border Styles

Public Const BF_ADJUST = &H2000
Public Const BF_BOTTOM = &H8
Public Const BF_DIAGONAL = &H10
Public Const BF_FLAT = &H4000
Public Const BF_LEFT = &H1
Public Const BF_MIDDLE = &H800
Public Const BF_MONO = &H8000
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000
Public Const BF_TOP = &H2
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)

Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

'' System Colors


Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_GRADIENTACTIVECAPTION = 27
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_GRADIENTINACTIVECAPTION = 28
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8

Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_INFOTEXT = 23
Public Const COLOR_INFOBK = 24

' Windows 98/2000

Public Const COLOR_HOTLIGHT = 26

'' Draw Modes

Public Const OPAQUE = 2
Public Const TRANSPARENT = 1

'' Text Clipboard Formats

Public Const CF_OEMTEXT = 7
Public Const CF_UNICODETEXT = 13
Public Const CF_TEXT = 1

'' Text Alignment

Public Const TA_BASELINE = 24
Public Const TA_BOTTOM = 8
Public Const TA_CENTER = 6
Public Const TA_LEFT = 0&
Public Const TA_NOUPDATECP = 0&
Public Const TA_RIGHT = 2
Public Const TA_TOP = 0&
Public Const TA_UPDATECP = 1

Public Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)

'' Check states.

Public Const DRWCHK_NORMAL = 0&
Public Const DRWCHK_SELECTED = 1
Public Const DRWCHK_DISABLED = 2

'' Draw Types

Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

'' Draw states

Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10         ' Gray string appearance '
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Public Const DSS_RIGHT = &H8000

Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_RIGHTALIGN = &H8&

Public Const TPM_TOPALIGN = &H0&
Public Const TPM_VCENTERALIGN = &H10&
Public Const TPM_BOTTOMALIGN = &H20&

Public Const TPM_HORIZONTAL = &H0&            '' Horz alignment matters more ''
Public Const TPM_VERTICAL = &H40&             '' Vert alignment matters more ''
Public Const TPM_NONOTIFY = &H80&             '' Don't send any notification msgs ''
Public Const TPM_RETURNCMD = &H100&

Public Const TPM_RECURSE = &H1&

'' LOGFONT face size

Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Public Const LF_FACESIZEW = 64
Public Const LF_FULLFACESIZEW = 128

'' GetObject object constants

Public Const OBJ_BITMAP = 7
Public Const OBJ_BRUSH = 2
Public Const OBJ_DC = 3
Public Const OBJ_ENHMETADC = 12
Public Const OBJ_ENHMETAFILE = 13
Public Const OBJ_EXTPEN = 11
Public Const OBJ_FONT = 6
Public Const OBJ_MEMDC = 10
Public Const OBJ_METADC = 4
Public Const OBJ_METAFILE = 9
Public Const OBJ_PAL = 5
Public Const OBJ_PEN = 1
Public Const OBJ_REGION = 8

' Gradient Fill flags ' Windows 2000

Public Const GRADIENT_FILL_RECT_H = &H0&
Public Const GRADIENT_FILL_RECT_V = &H1&
Public Const GRADIENT_FILL_TRIANGLE = &H2&
Public Const GRADIENT_FILL_OP_FLAG = &HFF&

' Brush Styles

Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERN8X8 = 8
Public Const BS_DIBPATTERNPT = 6
Public Const BS_HATCHED = 2
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_PATTERN = 3
Public Const BS_PATTERN8X8 = 7
Public Const BS_SOLID = 0&

'' Hatch brush constants

Public Const HS_BDIAGONAL = 3
Public Const HS_BDIAGONAL1 = 7
Public Const HS_CROSS = 4
Public Const HS_DIAGCROSS = 5
Public Const HS_DITHEREDBKCLR = 24
Public Const HS_DITHEREDCLR = 20
Public Const HS_DITHEREDTEXTCLR = 22
Public Const HS_FDIAGONAL = 2
Public Const HS_FDIAGONAL1 = 6
Public Const HS_HALFTONE = 18
Public Const HS_HORIZONTAL = 0&
Public Const HS_NOSHADE = 17
Public Const HS_SOLID = 8
Public Const HS_SOLIDBKCLR = 23
Public Const HS_SOLIDCLR = 19
Public Const HS_SOLIDTEXTCLR = 21
Public Const HS_VERTICAL = 1

' Pen Styles

Public Const PS_ALTERNATE = 8
Public Const PS_COSMETIC = &H0
Public Const PS_DASH = 1
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_DOT = 2
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_ENDCAP_MASK = &HF00
Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_GEOMETRIC = &H10000
Public Const PS_INSIDEFRAME = 6
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MASK = &HF000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_ROUND = &H0
Public Const PS_NULL = 5
Public Const PS_SOLID = 0&
Public Const PS_STYLE_MASK = &HF
Public Const PS_ptcMASK = &HF0000
Public Const PS_USERSTYLE = 7

Public Const IMC_SETSTATUSWINDOWPOS = &H10
Public Const IMC_SETCOMPOSITIONWINDOW = &HC
Public Const IMC_SETCOMPOSITIONFONT = &HA
Public Const IMC_SETCANDIDATEPOS = &H8
Public Const IMC_OPENSTATUSWINDOW = &H22
Public Const IMC_GETSTATUSWINDOWPOS = &HF
Public Const IMC_GETCOMPOSITIONWINDOW = &HB
Public Const IMC_GETCOMPOSITIONFONT = &H9
Public Const IMC_GETCANDIDATEPOS = &H7
Public Const IMC_CLOSESTATUSWINDOW = &H21

' GetDeviceCaps parameters

Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

' System Parameters info constants (to retrieve the current non-client font)

Public Const SPI_GETACCESSTIMEOUT = 60
Public Const SPI_GETANIMATION = 72
Public Const SPI_GETBEEP = 1
Public Const SPI_GETBORDER = 5
Public Const SPI_GETDEFAULTINPUTLANG = 89
Public Const SPI_GETDRAGFULLWINDOWS = 38
Public Const SPI_GETFASTTASKSWITCH = 35
Public Const SPI_GETFILTERKEYS = 50
Public Const SPI_GETFONTSMOOTHING = 74
Public Const SPI_GETGRIDGRANULARITY = 18
Public Const SPI_GETHIGHCONTRAST = 66
Public Const SPI_GETICONMETRICS = 45
Public Const SPI_GETICONTITLELOGFONT = 31
Public Const SPI_GETICONTITLEWRAP = 25
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_GETKEYBOARDPREF = 68
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_GETLOWPOWERACTIVE = 83
Public Const SPI_GETLOWPOWERTIMEOUT = 79
Public Const SPI_GETMENUDROPALIGNMENT = 27
Public Const SPI_GETMINIMIZEDMETRICS = 43
Public Const SPI_GETMOUSE = 3
Public Const SPI_GETMOUSEKEYS = 54
Public Const SPI_GETMOUSETRAILS = 94
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Const SPI_GETPOWEROFFACTIVE = 84
Public Const SPI_GETPOWEROFFTIMEOUT = 80
Public Const SPI_GETSCREENREADER = 70
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SPI_GETSCREENSAVETIMEOUT = 14
Public Const SPI_GETSERIALKEYS = 62
Public Const SPI_GETSHOWSOUNDS = 56
Public Const SPI_GETSOUNDSENTRY = 64
Public Const SPI_GETSTICKYKEYS = 58
Public Const SPI_GETTOGGLEKEYS = 52
Public Const SPI_GETWINDOWSEXTENSION = 92
Public Const SPI_GETWORKAREA = 48

Public Const SPI_ICONHORIZONTALSPACING = 13
Public Const SPI_ICONVERTICALSPACING = 24

Public Const SPI_LANGDRIVER = 12

Public Const SPI_SCREENSAVERRUNNING = 97

Public Const SPI_SETACCESSTIMEOUT = 61
Public Const SPI_SETANIMATION = 73
Public Const SPI_SETBEEP = 2
Public Const SPI_SETBORDER = 6
Public Const SPI_SETCURSORS = 87
Public Const SPI_SETDEFAULTINPUTLANG = 90
Public Const SPI_SETDESKPATTERN = 21
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDOUBLECLICKTIME = 32
Public Const SPI_SETDOUBLECLKHEIGHT = 30
Public Const SPI_SETDOUBLECLKWIDTH = 29
Public Const SPI_SETDRAGFULLWINDOWS = 37
Public Const SPI_SETDRAGHEIGHT = 77
Public Const SPI_SETDRAGWIDTH = 76
Public Const SPI_SETFASTTASKSWITCH = 36
Public Const SPI_SETFILTERKEYS = 51
Public Const SPI_SETFONTSMOOTHING = 75
Public Const SPI_SETGRIDGRANULARITY = 19
Public Const SPI_SETHANDHELD = 78
Public Const SPI_SETHIGHCONTRAST = 67
Public Const SPI_SETICONMETRICS = 46
Public Const SPI_SETICONS = 88
Public Const SPI_SETICONTITLELOGFONT = 34
Public Const SPI_SETICONTITLEWRAP = 26
Public Const SPI_SETKEYBOARDDELAY = 23
Public Const SPI_SETKEYBOARDPREF = 69
Public Const SPI_SETKEYBOARDSPEED = 11
Public Const SPI_SETLANGTOGGLE = 91
Public Const SPI_SETLOWPOWERACTIVE = 85
Public Const SPI_SETLOWPOWERTIMEOUT = 81
Public Const SPI_SETMENUDROPALIGNMENT = 28
Public Const SPI_SETMINIMIZEDMETRICS = 44
Public Const SPI_SETMOUSE = 4
Public Const SPI_SETMOUSEBUTTONSWAP = 33
Public Const SPI_SETMOUSEKEYS = 55
Public Const SPI_SETMOUSETRAILS = 93
Public Const SPI_SETNONCLIENTMETRICS = 42
Public Const SPI_SETPENWINDOWS = 49
Public Const SPI_SETPOWEROFFACTIVE = 86
Public Const SPI_SETPOWEROFFTIMEOUT = 82
Public Const SPI_SETSCREENREADER = 71
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_SETSCREENSAVETIMEOUT = 15
Public Const SPI_SETSERIALKEYS = 63
Public Const SPI_SETSHOWSOUNDS = 57
Public Const SPI_SETSOUNDSENTRY = 65
Public Const SPI_SETSTICKYKEYS = 59
Public Const SPI_SETTOGGLEKEYS = 53
Public Const SPI_SETWORKAREA = 47

Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0&
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0&
Public Const HWND_TOPMOST = -1

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Const EM_CANUNDO = &HC6
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const EM_FMTLINES = &HC8
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETMODIFY = &HB8
Public Const EM_GETPASSWORDCHAR = &HD2
Public Const EM_GETRECT = &HB2
Public Const EM_GETSEL = &HB0
Public Const EM_GETTHUMB = &HBE
Public Const EM_GETWORDBREAKPROC = &HD1
Public Const EM_LIMITTEXT = &HC5
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINESCROLL = &HB6
Public Const EM_REPLACESEL = &HC2
Public Const EM_SCROLL = &HB5
Public Const EM_SCROLLCARET = &HB7
Public Const EM_SETHANDLE = &HBC
Public Const EM_SETMODIFY = &HB9
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_SETREADONLY = &HCF
Public Const EM_SETRECT = &HB3
Public Const EM_SETRECTNP = &HB4
Public Const EM_SETSEL = &HB1
Public Const EM_SETTABSTOPS = &HCB
Public Const EM_SETWORDBREAKPROC = &HD0
Public Const EM_UNDO = &HC7

Public Const EN_CHANGE = &H300
Public Const EN_ERRSPACE = &H500
Public Const EN_HSCROLL = &H601
Public Const EN_KILLFOCUS = &H200
Public Const EN_MAXTEXT = &H501
Public Const EN_SETFOCUS = &H100
Public Const EN_UPDATE = &H400
Public Const EN_VSCROLL = &H602

' BitBlt flags

Public Const MERGEPAINT = &HBB0226
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020

''' Window movement / Window attributes

' Combobox Window Messages (including WM_MEASUREITEM and WM_DRAWITEM)

Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_DIR = &H145
Public Const CB_ERR = (-1)
Public Const CB_ERRSPACE = (-2)
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETDROPPEDCONTROLRECT = &H152
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_GETEDITSEL = &H140
Public Const CB_GETEXTENDEDUI = &H156
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETLOCALE = &H15A
Public Const CB_INSERTSTRING = &H14A
Public Const CB_LIMITTEXT = &H141
Public Const CB_MSGMAX = &H15B
Public Const CB_OKAY = 0&
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SETEDITSEL = &H142
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_SETITEMDATA = &H151
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_SETLOCALE = &H159
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_GETTOPINDEX = &H15B
Public Const CB_SETTOPINDEX = &H15C

' Combo Box Notify Constants

Public Const CBN_CLOSEUP = 8
Public Const CBN_DBLCLK = 2
Public Const CBN_DROPDOWN = 7
Public Const CBN_EDITCHANGE = 5
Public Const CBN_EDITUPDATE = 6
Public Const CBN_ERRSPACE = (-1)
Public Const CBN_KILLFOCUS = 4
Public Const CBN_SELCHANGE = 1
Public Const CBN_SELENDCANCEL = 10
Public Const CBN_SELENDOK = 9
Public Const CBN_SETFOCUS = 3

' Combo Box Style Constants

Public Const CBS_AUTOHSCROLL = &H40&
Public Const CBS_DISABLENOSCROLL = &H800&
Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&
Public Const CBS_HASSTRINGS = &H200&
Public Const CBS_NOINTEGRALHEIGHT = &H400&
Public Const CBS_OEMCONVERT = &H80&
Public Const CBS_OWNERDRAWFIXED = &H10&
Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CBS_SIMPLE = &H1&
Public Const CBS_SORT = &H100&

' List Box Commands / Constants

Public Const LB_ADDSTRING = &H180
Public Const LB_CTLCODE = 0&
Public Const LB_DELETESTRING = &H182
Public Const LB_DIR = &H18D
Public Const LB_ERR = (-1)
Public Const LB_ERRSPACE = (-2)
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETANCHORINDEX = &H19D
Public Const LB_GETCARETINDEX = &H19F
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETCURSEL = &H188
Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const LB_GETITEMRECT = &H198
Public Const LB_GETLOCALE = &H1A6
Public Const LB_GETSEL = &H187
Public Const LB_GETSELCOUNT = &H190
Public Const LB_GETSELITEMS = &H191
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETTOPINDEX = &H18E
Public Const LB_INSERTSTRING = &H181
Public Const LB_MSGMAX = &H1A8
Public Const LB_OKAY = 0&
Public Const LB_RESETCONTENT = &H184
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SELITEMRANGE = &H19B
Public Const LB_SELITEMRANGEEX = &H183
Public Const LB_SETANCHORINDEX = &H19C
Public Const LB_SETCARETINDEX = &H19E
Public Const LB_SETCOLUMNWIDTH = &H195
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETITEMDATA = &H19A
Public Const LB_SETITEMHEIGHT = &H1A0
Public Const LB_SETLOCALE = &H1A5
Public Const LB_SETSEL = &H185
Public Const LB_SETTABSTOPS = &H192
Public Const LB_SETTOPINDEX = &H197

' List Box Messages

Public Const LBN_DBLCLK = 2
Public Const LBN_ERRSPACE = (-2)
Public Const LBN_KILLFOCUS = 5
Public Const LBN_SELCANCEL = 3
Public Const LBN_SELCHANGE = 1
Public Const LBN_SETFOCUS = 4

' List Box Styles

Public Const LBS_DISABLENOSCROLL = &H1000&
Public Const LBS_EXTENDEDSEL = &H800&
Public Const LBS_HASSTRINGS = &H40&
Public Const LBS_MULTICOLUMN = &H200&
Public Const LBS_MULTIPLESEL = &H8&
Public Const LBS_NODATA = &H2000&
Public Const LBS_NOINTEGRALHEIGHT = &H100&
Public Const LBS_NOREDRAW = &H4&
Public Const LBS_NOTIFY = &H1&
Public Const LBS_OWNERDRAWFIXED = &H10&
Public Const LBS_OWNERDRAWVARIABLE = &H20&
Public Const LBS_SORT = &H2&
Public Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)
Public Const LBS_USETABSTOPS = &H80&
Public Const LBS_WANTKEYBOARDINPUT = &H400&

'' Edit Styles

Public Const ES_AUTOHSCROLL = &H80&
Public Const ES_AUTOVSCROLL = &H40&
Public Const ES_CENTER = &H1&
Public Const ES_LEFT = &H0&
Public Const ES_LOWERCASE = &H10&
Public Const ES_MULTILINE = &H4&
Public Const ES_NOHIDESEL = &H100&
Public Const ES_OEMCONVERT = &H400&
Public Const ES_PASSWORD = &H20&
Public Const ES_READONLY = &H800&
Public Const ES_RIGHT = &H2&
Public Const ES_UPPERCASE = &H8&
Public Const ES_WANTRETURN = &H1000&

Public Const ES_STD = (ES_AUTOHSCROLL + ES_LEFT)
Public Const ES_MULTI_STD = (ES_AUTOHSCROLL + ES_LEFT + ES_MULTILINE + ES_WANTRETURN) + WS_VSCROLL

'' Button Styles

Public Const BS_PUSHBUTTON = &H0
Public Const BS_DEFPUSHBUTTON = &H1
Public Const BS_CHECKBOX = &H2
Public Const BS_AUTOCHECKBOX = &H3
Public Const BS_RADIOBUTTON = &H4
Public Const BS_3STATE = &H5
Public Const BS_AUTO3STATE = &H6
Public Const BS_GROUPBOX = &H7
Public Const BS_USERBUTTON = &H8
Public Const BS_AUTORADIOBUTTON = &H9
Public Const BS_OWNERDRAW = &HB
Public Const BS_LEFTTEXT = &H20
'' (WINVER >= =&H0400)
Public Const BS_TEXT = &H0
Public Const BS_ICON = &H40
Public Const BS_BITMAP = &H80
Public Const BS_LEFT = &H100
Public Const BS_RIGHT = &H200
Public Const BS_CENTER = &H300
Public Const BS_TOP = &H400
Public Const BS_BOTTOM = &H800
Public Const BS_VCENTER = &HC00
Public Const BS_PUSHLIKE = &H1000
Public Const BS_MULTILINE = &H2000
Public Const BS_NOTIFY = &H4000
Public Const BS_FLAT = &H8000
Public Const BS_RIGHTBUTTON = BS_LEFTTEXT
'' (End) '' WINVER >= =&H0400 ''

''
 '' User Button Notification Codes
 ''
Public Const BN_CLICKED = 0
Public Const BN_PAINT = 1
Public Const BN_HILITE = 2
Public Const BN_UNHILITE = 3
Public Const BN_DISABLE = 4
Public Const BN_DOUBLECLICKED = 5
'' (WINVER >= =&H0400)
Public Const BN_PUSHED = BN_HILITE
Public Const BN_UNPUSHED = BN_UNHILITE
Public Const BN_DBLCLK = BN_DOUBLECLICKED
Public Const BN_SETFOCUS = 6
Public Const BN_KILLFOCUS = 7
'' (End) '' WINVER >= =&H0400 ''

''
 '' Button Control Messages
 ''
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const BM_GETSTATE = &HF2
Public Const BM_SETSTATE = &HF3
Public Const BM_SETSTYLE = &HF4
'' (WINVER >= =&H0400)
Public Const BM_CLICK = &HF5
Public Const BM_GETIMAGE = &HF6
Public Const BM_SETIMAGE = &HF7

Public Const BST_UNCHECKED = &H0
Public Const BST_CHECKED = &H1
Public Const BST_INDETERMINATE = &H2
Public Const BST_PUSHED = &H4
Public Const BST_FOCUS = &H8
'' (End) '' WINVER >= =&H0400 ''

''
 '' Static Control Constants
 ''
Public Const SS_LEFT = &H0
Public Const SS_CENTER = &H1
Public Const SS_RIGHT = &H2
Public Const SS_ICON = &H3
Public Const SS_BLACKRECT = &H4
Public Const SS_GRAYRECT = &H5
Public Const SS_WHITERECT = &H6
Public Const SS_BLACKFRAME = &H7
Public Const SS_GRAYFRAME = &H8
Public Const SS_WHITEFRAME = &H9
Public Const SS_USERITEM = &HA
Public Const SS_SIMPLE = &HB
Public Const SS_LEFTNOWORDWRAP = &HC
'' (WINVER >= =&H0400)
Public Const SS_OWNERDRAW = &HD
Public Const SS_BITMAP = &HE
Public Const SS_ENHMETAFILE = &HF
Public Const SS_ETCHEDHORZ = &H10
Public Const SS_ETCHEDVERT = &H11
Public Const SS_ETCHEDFRAME = &H12
Public Const SS_TYPEMASK = &H1F
'' (End) '' WINVER >= =&H0400 ''
Public Const SS_NOPREFIX = &H80              '' Don't do "&" character translation ''
'' (WINVER >= =&H0400)
Public Const SS_NOTIFY = &H100
Public Const SS_CENTERIMAGE = &H200
Public Const SS_RIGHTJUST = &H400
Public Const SS_REALSIZEIMAGE = &H800
Public Const SS_SUNKEN = &H1000
Public Const SS_ENDELLIPSIS = &H4000
Public Const SS_PATHELLIPSIS = &H8000
Public Const SS_WORDELLIPSIS = &HC000
Public Const SS_ELLIPSISMASK = &HC000
'' (End) '' WINVER >= =&H0400 ''



'' ndef NOWINMESSAGES
''
 '' Static Control Mesages
 ''
Public Const STM_SETICON = &H170
Public Const STM_GETICON = &H171
'' (WINVER >= =&H0400)
Public Const STM_SETIMAGE = &H172
Public Const STM_GETIMAGE = &H173
Public Const STN_CLICKED = 0
Public Const STN_DBLCLK = 1
Public Const STN_ENABLE = 2
Public Const STN_DISABLE = 3
'' (End) '' WINVER >= =&H0400 ''
Public Const STM_MSGMAX = &H174



'' Frame Control

Public Const DFC_BUTTON = 4
Public Const DFC_CAPTION = 1
Public Const DFC_MENU = 2
Public Const DFC_SCROLL = 3

Public Const DFCS_ADJUSTRECT = &H2000
Public Const DFCS_BUTTON3STATE = &H8
Public Const DFCS_BUTTONCHECK = &H0
Public Const DFCS_BUTTONPUSH = &H10
Public Const DFCS_BUTTONRADIO = &H4
Public Const DFCS_BUTTONRADIOIMAGE = &H1
Public Const DFCS_BUTTONRADIOMASK = &H2
Public Const DFCS_CAPTIONCLOSE = &H0
Public Const DFCS_CAPTIONHELP = &H4
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONRESTORE = &H3
Public Const DFCS_CHECKED = &H400
Public Const DFCS_FLAT = &H4000
Public Const DFCS_INACTIVE = &H100
Public Const DFCS_MENUARROW = &H0
Public Const DFCS_MENUARROWRIGHT = &H4
Public Const DFCS_MENUBULLET = &H2
Public Const DFCS_MENUCHECK = &H1
Public Const DFCS_MONO = &H8000
Public Const DFCS_PUSHED = &H200
Public Const DFCS_SCROLLCOMBOBOX = &H5
Public Const DFCS_SCROLLDOWN = &H1
Public Const DFCS_SCROLLLEFT = &H2
Public Const DFCS_SCROLLRIGHT = &H3
Public Const DFCS_SCROLLSIZEGRIP = &H8
Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10
Public Const DFCS_SCROLLUP = &H0

Public Const SW_ERASE = &H4
Public Const SW_HIDE = 0&
Public Const SW_INVALIDATE = &H2
Public Const SW_MAX = 10
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_OTHERUNZOOM = 4
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTCLOSING = 1
Public Const SW_PARENTOPENING = 3
Public Const SW_RESTORE = 9
Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

'' System Cursors

Public Const IDC_ARROW = (32512&)
Public Const IDC_IBEAM = (32513&)
Public Const IDC_WAIT = (32514&)
Public Const IDC_CROSS = (32515&)
Public Const IDC_UPARROW = (32516&)
Public Const IDC_SIZEAPI = (32640&)            '' OBSOLETE: use IDC_SIZEALL ''
Public Const IDC_ICON = (32641&)            '' OBSOLETE: use IDC_ARROW ''
Public Const IDC_SIZENWSE = (32642&)
Public Const IDC_SIZENESW = (32643&)
Public Const IDC_SIZEWE = (32644&)
Public Const IDC_SIZENS = (32645&)
Public Const IDC_SIZEALL = (32646&)
Public Const IDC_NO = (32648&)             ''not in win3.1 ''
Public Const IDC_HAND = (32649&)
Public Const IDC_APPSTARTING = (32650&)    ''not in win3.1 ''
Public Const IDC_HELP = (32651&)


' Window Messages

Public Const HTBORDER = 18
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTCAPTION = 2
Public Const HTCLIENT = 1
Public Const HTERROR = (-2)
Public Const HTGROWBOX = 4
Public Const HTHSCROLL = 6
Public Const HTLEFT = 10
Public Const HTMAXBUTTON = 9
Public Const HTMENU = 5
Public Const HTMINBUTTON = 8
Public Const HTNOWHERE = 0&

Public Const WM_ERASEBKGND = &H14
Public Const WM_DISPLAYCHANGE = &H7E

Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_COMPAREITEM = &H39
Public Const WM_COMMAND = &H111
Public Const WM_MENUCHAR = &H120
Public Const WM_MENUSELECT = &H11F

Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117

Public Const WM_MENURBUTTONUP = &H122
Public Const WM_MENUDRAG = &H123
Public Const WM_MENUGETOBJECT = &H124
Public Const WM_UNINITMENUPOPUP = &H125
Public Const WM_MENUCOMMAND = &H126

Public Const WM_USER = &H400
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_TIMER = &H113
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCOMMAND = &H112
Public Const WM_KEYUP = &H101
Public Const WM_KEYLAST = &H108
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYFIRST = &H100
Public Const WM_ACTIVATE = &H6
Public Const WM_SETTEXT = &HC
Public Const WM_MOVE = &H3
Public Const WM_PAINT = &HF
Public Const WM_SETREDRAW = &HB
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETFONT = &H30
Public Const WM_DESTROY = &H2
Public Const WM_SETFOCUS = &H7
Public Const WM_SETCURSOR = &H20
Public Const WM_ENABLE = &HA
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SIZE = &H5
Public Const WM_CHAR = &H102
Public Const WM_KILLFOCUS = &H8
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114

Public Const WM_NEXTMENU = &H213

Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_MBUTTON = &H10
Public Const MK_SHIFT = &H4
Public Const MK_CONTROL = &H8

Public Const WM_COPY = &H301
Public Const WM_COPYDATA = &H4A
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302

' Multiple Document Interface Window Messages

Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICASCADE = &H227
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDINEXT = &H224
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDISETMENU = &H230
Public Const WM_MDITILE = &H226

' Non-client area painting

Public Const WM_NCPAINT = &H85
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCHITTEST = &H84
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCREATE = &H81
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_IME_KEYUP = &H291
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_STARTCOMPOSITION = &H10D

Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Public Const IMAGE_BITMAP = 0&
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const IMAGE_ENHMETAFILE = 3

Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_ERASENOW = &H200
Public Const RDW_FRAME = &H400
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_INVALIDATE = &H1
Public Const RDW_NOCHILDREN = &H40
Public Const RDW_NOERASE = &H20
Public Const RDW_NOFRAME = &H800
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_UPDATENOW = &H100
Public Const RDW_VALIDATE = &H8

' GetCharacterPlacement Constant Values

Public Const GCP_CLASSIN = &H80000
Public Const GCP_DBCS = &H1&
Public Const GCP_DIACRITIC = &H100&
Public Const GCP_DISPLAYZWG = &H400000
Public Const GCP_ERROR = &H8000&
Public Const GCP_GLYPHSHAPE = &H10&
Public Const GCP_JUSTIFY = &H10000
Public Const GCP_JUSTIFYIN = &H200000
Public Const GCP_KASHIDA = &H400&
Public Const GCP_LIGATE = &H20&
Public Const GCP_MAXEXTENT = &H100000
Public Const GCP_NEUTRALOVERRIDE = &H2000000
Public Const GCP_NODIACRITICS = &H20000
Public Const GCP_NUMERICOVERRIDE = &H1000000
Public Const GCP_NUMERICSLATIN = &H4000000
Public Const GCP_NUMERICSLOCAL = &H8000000
Public Const GCP_REORDER = &H2&
Public Const GCP_SYMSWAPOFF = &H800000
Public Const GCP_USEKERNING = &H8&

' GetCharacterPlacement Class Monikers

Public Const GCPCLASS_ARABIC = 2
Public Const GCPCLASS_HEBREW = 2
Public Const GCPCLASS_LATIN = 1
Public Const GCPCLASS_LATINNUMBER = 5
Public Const GCPCLASS_LATINNUMERICSEPARATOR = 7
Public Const GCPCLASS_LATINNUMERICTERMINATOR = 6
Public Const GCPCLASS_LOCALNUMBER = 4
Public Const GCPCLASS_NEUTRAL = 3
Public Const GCPCLASS_NUMERICSEPARATOR = 8
Public Const GCPCLASS_PREBOUNDLTR = &H40
Public Const GCPCLASS_PREBOUNDRTL = &H80


Public Const ETO_OPAQUE = &H2
Public Const ETO_CLIPPED = &H4

' If (WINVER >= =&H0400)
Public Const ETO_GLYPH_INDEX = &H10
Public Const ETO_RTLREADING = &H80
Public Const ETO_NUMERICSLOCAL = &H400
Public Const ETO_NUMERICSLATIN = &H800
Public Const ETO_IGNORELANGUAGE = &H1000
' ' WINVER >= =&H0400

' If  (_WIN32_WINNT >= =&H0500) (Windows 2000)
Public Const ETO_PDY = &H2000
' ' (_WIN32_WINNT >= =&H0500)

Public Const FLI_GLYPHS = &H40000
Public Const FLI_MASK = &H103B

'' Scroll Bar Types

Public Const SB_BOTH = 3
Public Const SB_BOTTOM = 7
Public Const SB_CTL = 2
Public Const SB_ENDSCROLL = 8
Public Const SB_HORZ = 0&
Public Const SB_LEFT = 6
Public Const SB_LINEDOWN = 1
Public Const SB_LINELEFT = 0&
Public Const SB_LINERIGHT = 1
Public Const SB_LINEUP = 0&
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGELEFT = 2
Public Const SB_PAGERIGHT = 3
Public Const SB_PAGEUP = 2
Public Const SB_RIGHT = 7
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6
Public Const SB_VERT = 1

'' Scroll Bar Messages

Public Const SBM_ENABLE_ARROWS = &HE4
Public Const SBM_GETPOS = &HE1
Public Const SBM_GETRANGE = &HE3
Public Const SBM_SETPOS = &HE0
Public Const SBM_SETRANGE = &HE2
Public Const SBM_SETRANGEREDRAW = &HE6

'' Scroll Bar Window Styles

Public Const SBS_BOTTOMALIGN = &H4&
Public Const SBS_HORZ = &H0&
Public Const SBS_LEFTALIGN = &H2&
Public Const SBS_RIGHTALIGN = &H4&
Public Const SBS_SIZEBOX = &H8&
Public Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Public Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Public Const SBS_TOPALIGN = &H2&
Public Const SBS_VERT = &H1&

'' EnableScrollBar() flags

Public Const ESB_DISABLE_BOTH = &H3
Public Const ESB_DISABLE_DOWN = &H2
Public Const ESB_DISABLE_LEFT = &H1
Public Const ESB_DISABLE_RIGHT = &H2
Public Const ESB_DISABLE_UP = &H1

Public Const ESB_DISABLE_RTDN = ESB_DISABLE_RIGHT
Public Const ESB_DISABLE_LTUP = ESB_DISABLE_LEFT

Public Const ESB_ENABLE_BOTH = &H0

Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Public Const MNC_IGNORE = 0&
Public Const MNC_CLOSE = (1 * &H10000)
Public Const MNC_EXECUTE = (2 * &H10000)
Public Const MNC_SELECT = (3 * &H10000)

Public Const FALT = &H10
Public Const FCONTROL = &H8
Public Const FNOINVERT = &H2
Public Const FSHIFT = &H4

Public Const NIF_CUSTOM_MSG = (WM_USER + &H108)

Public Const STRETCH_HALFTONE = 4
Public Const STRETCH_DELETESCANS = 3
Public Const STRETCH_ORSCANS = 2
Public Const STRETCHBLTC = 2048

' fMask flags

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20

'' New for Windows 98/2000

Public Const MIIM_STRING = &H40
Public Const MIIM_BITMAP = &H80
Public Const MIIM_FTYPE = &H100

'' End fMask flags

' Menu Flags

Public Const MF_INSERT = &H0&
Public Const MF_CHANGE = &H80&
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_REMOVE = &H1000&

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

Public Const MF_SEPARATOR = &H800&

Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&

Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_USECHECKBITMAPS = &H200&

Public Const MF_STRING = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_OWNERDRAW = &H100&

Public Const MF_POPUP = &H10&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&

Public Const MF_UNHILITE = &H0&
Public Const MF_HILITE = &H80&

Public Const MF_DEFAULT = &H1000&
Public Const MF_SYSMENU = &H2000&
Public Const MF_HELP = &H4000&
Public Const MF_RIGHTJUSTIFY = &H4000&

Public Const MF_MOUSESELECT = &H8000&
Public Const MF_END = &H80&                    ' Obsolete -- only used by old RES files


Public Const MFT_STRING = MF_STRING
Public Const MFT_BITMAP = MF_BITMAP
Public Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Public Const MFT_MENUBREAK = MF_MENUBREAK
Public Const MFT_OWNERDRAW = MF_OWNERDRAW
Public Const MFT_RADIOGROUP = &H200&
Public Const MFT_SEPARATOR = MF_SEPARATOR
Public Const MFT_RIGHTORDER = &H2000&
Public Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY

Public Const MFS_GRAYED = &H3&
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = MF_CHECKED
Public Const MFS_HILITE = MF_HILITE
Public Const MFS_ENABLED = MF_ENABLED
Public Const MFS_UNCHECKED = MF_UNCHECKED
Public Const MFS_UNHILITE = MF_UNHILITE
Public Const MFS_DEFAULT = MF_DEFAULT

' New for Windows 2000/98

Public Const MFS_MASK = &H108B&
Public Const MFS_HOTTRACKDRAWN = &H10000000
Public Const MFS_CACHEDBMP = &H20000000
Public Const MFS_BOTTOMGAPDROP = &H40000000
Public Const MFS_TOPGAPDROP = &H80000000
Public Const MFS_GAPDROP = &HC0000000

' for the SetMenuInfo API function

Public Const MNS_NOCHECK = &H80000000
Public Const MNS_MODELESS = &H40000000
Public Const MNS_DRAGDROP = &H20000000
Public Const MNS_AUTODISMISS = &H10000000
Public Const MNS_NOTIFYBYPOS = &H8000000
Public Const MNS_CHECKORBMP = &H4000000

Public Const MIM_MAXHEIGHT = &H1
Public Const MIM_BACKGROUND = &H2
Public Const MIM_HELPID = &H4
Public Const MIM_MENUDATA = &H8
Public Const MIM_STYLE = &H10
Public Const MIM_APPLYTOSUBMENUS = &H80000000


'' System Object Identifier Constants

Public Const OBJID_WINDOW = &H0&
Public Const OBJID_SYSMENU = &HFFFFFFFF
Public Const OBJID_TITLEBAR = &HFFFFFFFE
Public Const OBJID_MENU = &HFFFFFFFD
Public Const OBJID_CLIENT = &HFFFFFFFC
Public Const OBJID_VSCROLL = &HFFFFFFFB
Public Const OBJID_HSCROLL = &HFFFFFFFA
Public Const OBJID_SIZEGRIP = &HFFFFFFF9
Public Const OBJID_CARET = &HFFFFFFF8
Public Const OBJID_CURSOR = &HFFFFFFF7
Public Const OBJID_ALERT = &HFFFFFFF6
Public Const OBJID_SOUND = &HFFFFFFF5
Public Const OBJID_QUERYCLASSNAMEIDX = &HFFFFFFF4
Public Const OBJID_NATIVEOM = &HFFFFFFF0



'' National Language / IME Support Functions

Public Const CP_ACP = 0&
Public Const CP_NONE = 0&
Public Const CP_OEMCP = 1
Public Const CP_RECTANGLE = 1
Public Const CP_REGION = 2
Public Const CP_WINANSI = 1004
Public Const CP_WINUNICODE = 1200

Public Const ANSI_CHARSET = 0&
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const HANGUL_CHARSET = 129
Public Const GB2312_CHARSET = 134
Public Const CHINESEBIG5_CHARSET = 136
Public Const OEM_CHARSET = 255
Public Const JOHAB_CHARSET = 130
Public Const HEBREW_CHARSET = 177
Public Const ARABIC_CHARSET = 178
Public Const GREEK_CHARSET = 161
Public Const TURKISH_CHARSET = 162
Public Const VIETNAMESE_CHARSET = 163
Public Const THAI_CHARSET = 222
Public Const EASTEUROPE_CHARSET = 238
Public Const RUSSIAN_CHARSET = 204

Public Const MAC_CHARSET = 77
Public Const BALTIC_CHARSET = 186

Public Const FS_LATIN1 = &H1&
Public Const FS_LATIN2 = &H2&
Public Const FS_CYRILLIC = &H4&
Public Const FS_GREEK = &H8&
Public Const FS_TURKISH = &H10&
Public Const FS_HEBREW = &H20&
Public Const FS_ARABIC = &H40&
Public Const FS_BALTIC = &H80&
Public Const FS_VIETNAMESE = &H100&
Public Const FS_THAI = &H10000
Public Const FS_JISJAPAN = &H20000
Public Const FS_CHINESESIMP = &H40000
Public Const FS_WANSUNG = &H80000
Public Const FS_CHINESETRAD = &H100000
Public Const FS_JOHAB = &H200000
Public Const FS_SYMBOL = &H80000000

'' Font Pitch and Family Constants

Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2

Public Const FF_DONTCARE = 0
Public Const FF_ROMAN = 16
Public Const FF_SWISS = 32
Public Const FF_MODERN = 48
Public Const FF_SCRIPT = 64
Public Const FF_DECORATIVE = 80

''' Shell_NotifyIcon Constants

' Shell_NotifyIcon features for Windows NT 5.0& (Windows 2000)

Public Const NIN_SELECT = (WM_USER + 0&)
Public Const NINF_KEY = &H1
Public Const NIN_KEYSELECT = (NIN_SELECT + NINF_KEY)
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4
Public Const NOTIFYICON_VERSION = 3
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10

Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2

' Regular NotifyIcon Features for Windows 95/98/NT

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' Notify Icon Infotip flags

Public Const NIIF_NONE = &H0

' icon flags are mutually exclusive
' and take only the lowest 2 bits

' Notify Icon Balloon Icon constants

Public Const NIIF_INFO = &H1
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3

Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_COLOR = &H2
Public Const LR_COPYRETURNORG = &H4
Public Const LR_COPYDELETEORG = &H8
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_VGACOLOR = &H80
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_SHARED = &H8000


Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE + SIF_PAGE + SIF_POS + SIF_TRACKPOS)


'' ChooseColor constants

Public Const CC_ANYCOLOR = &H100
Public Const CC_CHORD = 4
Public Const CC_CIRCLES = 1
Public Const CC_ELLIPSES = 8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_FULLOPEN = &H2
Public Const CC_INTERIORS = 128
Public Const CC_NONE = 0
Public Const CC_PIE = 2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_RGBINIT = &H1
Public Const CC_ROUNDRECT = 256
Public Const CC_SHOWHELP = &H8
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_STYLED = 32
Public Const CC_WIDE = 16
Public Const CC_WIDESTYLED = 64




'' ****** TYPES ******


Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SIZEAPI
        cx As Long
        cy As Long
End Type

Public Type DIMAPI
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

''' CreateBrushIndirect Brush structure

Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

''' CreatePenIndirect Pen structure

Public Type LOGPEN
        lopnStyle As Long
        lopnWidth As POINTAPI
        lopnColor As Long
End Type

''' CreateFontIndirect

Public Type LOGFONT
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
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Public Type LOGFONTW
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
        lfFaceName(1 To LF_FACESIZE) As Integer
End Type


Public Type CREATESTRUCT
        lpCreateParams As Long
        hInstance As Long
        hMenu As Long
        hWndParent As Long
        cy As Long
        cx As Long
        y As Long
        x As Long
        Style As Long
        lpszName As String
        lpszClass As String
        ExStyle As Long
End Type

Public Type CREATESTRUCT2
        lpCreateParams As Long
        hInstance As Long
        hMenu As Long
        hWndParent As Long
        cy As Long
        cx As Long
        y As Long
        x As Long
        Style As Long
        lpszName As String
        lpszClass As Long
        ExStyle As Long
End Type

Public Type MDICREATESTRUCT
        szClass As String
        szTitle As String
        hOwner As Long
        x As Long
        y As Long
        cx As Long
        cy As Long
        Style As Long
        lParam As Long
End Type

Public Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

''' Structure for the WM_COMPAREITEM message

Public Type COMPAREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        hwndItem As Long
        itemID1 As Long
        itemData1 As Long
        itemID2 As Long
        itemData2 As Long
End Type

''' Structure for the WM_DRAWITEM message

Public Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        ItemId As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hDC As Long
        rcItem As RECT
        ItemData As Long
End Type

''' Structure for the WM_MEASUREITEMSTRUCT

Public Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        ItemId As Long
        itemWidth As Long
        itemHeight As Long
        ItemData As Long
End Type

''' GetCharacterPlacement Results

Public Type GCP_RESULTS
        lStructSize As Long
        lpOutString As Long
        lpOrder As Long
        lpDx As Long
        lpCaretPos As Long
        lpClass As Long
        lpGlyphs As Long
        nGlyphs As Long
        nMaxFit As Long
End Type

''' Measuring and Enumerating

Public Type TEXTMETRIC
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
End Type

Public Type NEWTEXTMETRIC
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


Public Type TEXTMETRICW
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
        tmFirstChar As Integer
        tmLastChar As Integer
        tmDefaultChar As Integer
        tmBreakChar As Integer
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Public Type NEWTEXTMETRICW
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
        tmFirstChar As Integer
        tmLastChar As Integer
        tmDefaultChar As Integer
        tmBreakChar As Integer
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

Public Type FONTSIGNATURE
        fsUsb(4) As Long
        fsCsb(2) As Long
End Type


Public Type ENUMLOGFONTEXW
        elfLogFont As LOGFONTW
        elfFullName(1 To LF_FULLFACESIZE) As Integer
        elfStyle(1 To LF_FACESIZE) As Integer
        elfScript(1 To LF_FACESIZE) As Integer
End Type


Public Type ENUMLOGFONTEX
        elfLogFont As LOGFONT
        elfFullName(1 To LF_FULLFACESIZE) As Byte
        elfStyle(1 To LF_FACESIZE) As Byte
        elfScript(1 To LF_FACESIZE) As Byte
End Type

Public Type NEWTEXTMETRICEXW
        ntmTm As NEWTEXTMETRICW
        ntmFontSig As FONTSIGNATURE
End Type

Public Type NEWTEXTMETRICEX
        ntmTm As NEWTEXTMETRIC
        ntmFontSig As FONTSIGNATURE
End Type

''' Vertex

Public Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

''' Gradient Rectangle

Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type


'' UUID structure

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0& To 7) As Byte
End Type
        
''' Used by OleCreateFontIndirect
        
Public Type FONTDESC

    cbSize As Long
    
    szName As Long
    cySize As Long
    
    sWeight As Integer
    sCharset As Integer
    
    fItalic As Long
    fUnderline As Long
    fStrikethrough As Long
    
End Type


Public Type PICTDESCBMP
    SizeofStruct As Long
    PicType As Long
    hBitmap As Long
    hPalette As Long
End Type

Public Type PICTDESCICON
    SizeofStruct As Long
    PicType As Long
    hIcon As Long
    lReserved As Long
End Type


''' Used by SystemParametersInfo

Public Type NONCLIENTMETRICS
        cbSize As Long
        iBorderWidth As Long
        iScrollWidth As Long
        iScrollHeight As Long
        iCaptionWidth As Long
        iCaptionHeight As Long
        lfCaptionFont As LOGFONT
        iSMCaptionWidth As Long
        iSMCaptionHeight As Long
        lfSMCaptionFont As LOGFONT
        iMenuWidth As Long
        iMenuHeight As Long
        lfMenuFont As LOGFONT
        lfStatusFont As LOGFONT
        lfMessageFont As LOGFONT
End Type

Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Type BITMAPCOREHEADER
        bcSize As Long
        bcWidth As Integer
        bcHeight As Integer
        bcPlanes As Integer
        bcBitCount As Integer
End Type

Public Type BITMAPCOREINFO
        bmciHeader As BITMAPCOREHEADER
        bmciColors As Long  'array of RGB triple
End Type

Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As Long   'array of RGB quad
End Type

Public Type BITMAPV4HEADER
        bV4Size As Long
        bV4Width As Long
        bV4Height As Long
        bV4Planes As Integer
        bV4BitCount As Integer
        bV4V4Compression As Long
        bV4SizeImage As Long
        bV4XPelsPerMeter As Long
        bV4YPelsPerMeter As Long
        bV4ClrUsed As Long
        bV4ClrImportant As Long
        bV4RedMask As Long
        bV4GreenMask As Long
        bV4BlueMask As Long
        bV4AlphaMask As Long
        bV4CSType As Long
        bV4Endpoints As Long
        bV4GammaRed As Long
        bV4GammaGreen As Long
        bV4GammaBlue As Long
End Type

Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type


Public Type ICONINFO
        fIcon As Long
        xHotspot As Long
        yHotspot As Long
        hbmMask As Long
        hbmColor As Long
End Type

Public Type ICONMETRICS
    cbSize As Long
    iHorzSpacing As Long
    iVertSpacing As Long
    iTitleWrap As Long
    lfFont As LOGFONT
End Type

Public Type BITDATA
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Public Type BITDATAINV
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

' Old NotifyIconData

Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

' NotifyIconData for Window NT 5.0& (Windows 2000)

Public Type NOTIFYICONDATA5
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type


'' 2000+ GetVersionEx structure

Public Type OSVERSIONINFOEX
    OSVersionInfoSize As Long
    MajorVersion As Long
    MinorVersion As Long
    BuildNumber As Long
    PlatformId As Long
    CSDVersion As String * 128
    ServicePackMajor As Integer
    ServicePackMinor As Integer
    Reserved(0& To 1) As Integer
End Type

'' Old GetVersionEx structure

Public Type OSVERSIONINFO
        OSVersionInfoSize As Long
        MajorVersion As Long
        MinorVersion As Long
        BuildNumber As Long
        PlatformId As Long
        CSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'' Windows version constants enumerated (100 * Maj) + Min method

Public Enum WindowsVersionConstants
    Windows95 = 400
    WindowsNT = 400
    Windows98 = 410
    WindowsME = 490
    Windows2000 = 500
End Enum



Public Type PAINTSTRUCT
    hDC As Long
    fErase As Boolean
    rcpaint As RECT
    fRestore As Boolean
    fIncUpdate As Boolean
    rgbReserved(0 To 31) As Byte
End Type

' Old NotifyIconData

' NotifyIconData for Window NT 5.0& (Windows 2000)

''' Menu event records

Public Type MENU_EVENT_RECORD
        dwCommandId As Long
End Type

''' Menu item info for 95-2000

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
    
    ' Declared for Windows 2000/98.  Still backward compatible
    
    hbmpItem As Long
End Type

'' With string data.

Public Type MENUITEMINFOSTRING
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
    
    ' Declared for Windows 2000/98.  Still backward compatible
    
    hbmpItem As Long
End Type

''' Item template

Public Type MENUITEMTEMPLATE
        mtOption As Integer
        mtID As Integer
        mtString As Byte
End Type

''' Template header

Public Type MENUITEMTEMPLATEHEADER
        versionNumber As Integer
        Offset As Integer
End Type

''' Template parameters

Public Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type

''' New for Windows 2000/98

Public Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    Back As Long
    ContextHelpID As Long
    MenuData As Long
End Type

Public Type CHOOSE_COLOR
        lStructSize As Long
        hWndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As Long
End Type

Public Type MENUBARINFO
    cbSize As Long
    rcBar As RECT
    hMenu As Long
    hWndMenu As Long
    fBarFocused As Boolean
    fFocused As Boolean
End Type

Public Type ACCEL
        fVirt As Byte
        Key As Integer
        Cmd As Integer
End Type

Public Type MSG
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
End Type

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type


Public Type DIBDATA
    hDC As Long
    hBitmap As Long
    
    Info As BITMAPINFO
    Bytes() As Byte
End Type

'' ****** DECLARES ******

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nx As Long, ByVal nY As Long, lpSize As SIZEAPI) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nx As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

Public Declare Function GetWindowExtEx Lib "gdi32" (ByVal hDC As Long, lpSize As SIZEAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

'' other API functions

Public Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Public Declare Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCaretBlinkTime Lib "user32" () As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DestroyCaret Lib "user32" () As Long

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, ByVal lpBI As Long, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, ByVal lpBI As Long, ByVal wUsage As Long) As Long

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function LoadCursorA Lib "user32" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Public Declare Function LoadCursorW Lib "user32" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, lParam As Any) As Long

''' GDI declares

Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal hBitmap As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long

Public Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hDC As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Public Declare Function EnumFontFamiliesExW Lib "gdi32" (ByVal hDC As Long, lpLogFont As LOGFONTW, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long

'' mapping mode functions

Public Declare Function LPtoDP Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function DPtoLP Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long

''' Draw disabled/enabled icons/bitmaps and text

Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer, ByVal un As Long) As Long
Public Declare Function DrawStateW Lib "user32" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer, ByVal un As Long) As Long
Public Declare Function DrawStateB Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer, ByVal un As Long) As Long
Public Declare Function DrawStateBW Lib "user32" Alias "DrawStateW" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer, ByVal un As Long) As Long

''' Set the painting colors of the device context

Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal ColorRef As Long) As Long

Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long

'' Set background OPAQUE or TRANSPARENT

Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

'' Get colorref from system color index

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'' Device context functions

Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDCOrgEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Declare Function GetObjectW Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONTW) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long

''' Measuring, Coloring and Drawing Text

Public Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Public Declare Function TextOutW Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long

Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByVal lpRect As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal lpDx As Long) As Long
Public Declare Function ExtTextOutW Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByVal lpRect As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal lpDx As Long) As Long

Public Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hDC As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
Public Declare Function GetCharWidth32W Lib "gdi32" (ByVal hDC As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long

Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Public Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As SIZEAPI) As Long

Public Declare Function GetCharacterPlacement Lib "gdi32" Alias " GetCharacterPlacementA" (ByVal hDC As Long, ByVal lpsz As Long, ByVal n1 As Long, ByVal n2 As Long, lpGcpResults As GCP_RESULTS, ByVal dw As Long) As Long
Public Declare Function GetCharacterPlacementW Lib "gdi32" Alias " GetCharacterPlacementW" (ByVal hDC As Long, ByVal lpsz As Long, ByVal n1 As Long, ByVal n2 As Long, lpGcpResults As GCP_RESULTS, ByVal dw As Long) As Long

Public Declare Function GetFontLanguageInfo Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public Declare Function RedrawWindow2 Lib "user32" Alias "RedrawWindow" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function GetTextCharsetInfo Lib "gdi32" (ByVal hDC As Long, lpSig As FONTSIGNATURE, ByVal dwFlags As Long) As Long

'' draw icon

Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

'' draw 3d edge

Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

'' draw caption text

Public Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long

'' draw a frame

Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

'' draw a line

Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'' fill rectangle

Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'' flood fill

Public Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


'' Move painting origin on the canvas of the device

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Declare Function GdiFlush Lib "gdi32" () As Long

'' Ole Functions

Public Declare Function OleCreateFontIndirect Lib "olepro32" (pFontDesc As Any, riid As GUID, pFont As StdFont) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32" (ByVal Pict As Long, riid As GUID, fOwn As Boolean, ppvObj As Object) As Long

'' Muldiv
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Shell_NotifyIconA if OSVersion <= 4

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

' Shell_NotifyIconA if OSVersion >= 5

Public Declare Function Shell_NotifyIcon5 Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA5) As Long

'' Scroll Bar Functions

Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Public Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function ScrollWindow Lib "user32" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Public Declare Function ScrollWindowEx Lib "user32" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT, ByVal fuScroll As Long) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

'' Menu API functions

Public Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
Public Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hWnd As Long, ByVal hAccTable As Long, lpMsg As MSG) As Long
Public Declare Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As Long) As Long

Public Declare Function TrackPopupMenu Lib "user32" (ByVal Handle As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal Handle As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long

Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal Handle As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal Handle As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function InsertMenuItemW Lib "user32" (ByVal Handle As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal Handle As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function AppendMenuW Lib "user32" (ByVal Handle As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal Handle As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function InsertMenuW Lib "user32" (ByVal Handle As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Public Declare Function RemoveMenu Lib "user32" (ByVal Handle As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal Handle As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Declare Function GetMenuItemCount Lib "user32" (ByVal Handle As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal Handle As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal Handle As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function GetMenuContextHelpId Lib "user32" (ByVal Handle As Long) As Long
Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long

Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal Handle As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenuItemInfoW Lib "user32" (ByVal Handle As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function GetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As MENUINFO) As Long
Public Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, ByVal idObject As Long, ByVal idItem As Long, pmbi As MENUBARINFO) As Long

Public Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal Handle As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal Handle As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal Handle As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetMenuStringW Lib "user32" (ByVal Handle As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Public Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As MENUINFO) As Long

Public Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal Handle As Long) As Long
Public Declare Function SetMenuContextHelpId Lib "user32" (ByVal Handle As Long, ByVal dw As Long) As Long
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal Handle As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal Handle As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal Handle As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfoW Lib "user32" (ByVal Handle As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

''' Windows 98/2000

Public Declare Function GradientFillRect Lib "gdi32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Boolean


''' GDI declaresPublic Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, lParam As Any) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

''' OS Version retrieval from kernel

Public Declare Function GetVersion Lib "kernel32" Alias "GetVersionExA" (pData As OSVERSIONINFO) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (pData As OSVERSIONINFOEX) As Long

Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CreateWindowEx2 Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Public Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As Long, lpWndClass As WNDCLASS) As Long
Public Declare Function GetClassInfoW Lib "user32" (ByVal hInstance As Long, ByVal lpClassName As Long, lpWndClass As WNDCLASS) As Long

Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long

Public Declare Function RegisterClassW Lib "user32" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClassW Lib "user32" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long

Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal fErase As Long) As Long
Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long

'' verification functions

Public Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Public Declare Function IsBadHugeReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadHugeWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As String, ByVal ucchMax As Long) As Long
Public Declare Function IsBadWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long

Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long

Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSE_COLOR) As Long

Public Function GetActualColor(ByVal OleValue As OLE_COLOR) As Long

    If (OleValue And &H80000000) Then
        GetActualColor = GetSysColor(OleValue And &H7FF)
    Else
        GetActualColor = OleValue
    End If
End Function

Public Function CompRect(lpRect1 As RECT, lpRect2 As RECT) As Long

    Dim c As RECT
    
    c.Left = Abs(lpRect1.Left - lpRect2.Left)
    c.Top = Abs(lpRect1.Top - lpRect2.Top)
    c.Right = Abs(lpRect1.Right - lpRect2.Right)
    c.Bottom = Abs(lpRect1.Bottom - lpRect2.Bottom)
    
    CompRect = CLng(CDbl(c.Left + c.Top + c.Right + c.Bottom) / 4#)
    
End Function

Public Function WithinRect(rcBounds As RECT, ByVal x As Long, ByVal y As Long) As Boolean
    If ((rcBounds.Left <= x) And (rcBounds.Right >= x)) And _
        ((rcBounds.Top <= y) And (rcBounds.Bottom >= y)) Then
        
        WithinRect = True
    Else
        WithinRect = False
    End If
End Function

Public Function WithinRectPt(rcBounds As RECT, pt As POINTAPI) As Boolean
    With pt
        If ((rcBounds.Left <= .x) And (rcBounds.Right >= .x)) And _
            ((rcBounds.Top <= .y) And (rcBounds.Bottom >= .y)) Then
            
            WithinRectPt = True
        Else
            WithinRectPt = False
        End If
    End With
End Function

Public Function PointFromLong(ByVal lParam As Long, lpPoint As POINTAPI)
    Dim pt(0 To 1) As Integer
    
    CopyMemory pt(0), lParam, 4
    lpPoint.x = pt(0)
    lpPoint.y = pt(1)

End Function


Public Function PointToLong(lpPoint As POINTAPI, lParam As Long)
    Dim pt(0 To 1) As Integer
    
    pt(0) = CInt(lpPoint.x)
    pt(1) = CInt(lpPoint.y)
    
    CopyMemory lParam, pt(0), 4

End Function

Public Function ExpandRect(lpRect As RECT, ByVal Value As Long)
    lpRect.Left = lpRect.Left - Value
    lpRect.Top = lpRect.Top - Value
    lpRect.Right = lpRect.Right + Value
    lpRect.Bottom = lpRect.Bottom + Value
End Function


Public Function ShiftRect(lpRect As RECT, ByVal cx As Long, ByVal cy As Long)
    lpRect.Left = lpRect.Left + cx
    lpRect.Top = lpRect.Top + cy
    lpRect.Right = lpRect.Right + cx
    lpRect.Bottom = lpRect.Bottom + cy
End Function



Public Function BoundryX(lpRect As RECT, ByVal x As Long) As Long

    If Abs(lpRect.Left - x) < Abs(lpRect.Right - x) Then
        BoundryX = Abs(lpRect.Left - x)
    Else
        BoundryX = Abs(lpRect.Right - x)
    End If

End Function

Public Function BoundryY(lpRect As RECT, ByVal y As Long) As Long

    If Abs(lpRect.Top - y) < Abs(lpRect.Bottom - y) Then
        BoundryY = Abs(lpRect.Top - y)
    Else
        BoundryY = Abs(lpRect.Bottom - y)
    End If

End Function

Public Function NearBoundry(lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

    Dim ax As Long, _
        ay As Long
        
    ax = BoundryX(lpRect, x)
    ay = BoundryY(lpRect, y)
    
    If (ax < ay) Then NearBoundry = ax Else NearBoundry = ay

End Function

Public Function Pex(Value) As String
    Dim v As String, _
        i As Integer
        
    Dim c As Integer
        
    i = Len(Hex(Value))
    
    Select Case VarType(Value)
        Case vbByte:
            c = 2
        Case vbInteger:
            c = 4
        Case vbLong:
            c = 8
    End Select

    v = String(c - i, "0")
    v = v + Hex(Value)
    Pex = v
End Function




