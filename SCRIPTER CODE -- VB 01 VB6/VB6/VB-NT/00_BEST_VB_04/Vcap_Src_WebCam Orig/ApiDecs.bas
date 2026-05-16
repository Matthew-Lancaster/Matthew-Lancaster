Attribute VB_Name = "mApiDecs"
Option Explicit
'misc WinApi declares and constants are here
'remove this module if you use a WinAPI typelibrary

Public Const WS_CHILD As Long = &H40000000
Public Const WS_VISIBLE As Long = &H10000000
Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_NOMOVE As Long = &H2&
Public Const SWP_NOZORDER As Long = &H4&
Public Const SWP_NOSENDCHANGING As Long = &H400&   ' /* Don't send WM_WINDOWPOSCHANGING */
Public Const HWND_BOTTOM As Long = 1&
'/*
' * GetSystemMetrics() codes
' */
'#define SM_CXSCREEN             0
'#define SM_CYSCREEN             1
'#define SM_CXVSCROLL            2
'#define SM_CYHSCROLL            3
Public Const SM_CYCAPTION As Long = 4
Public Const SM_CXBORDER As Long = 5
Public Const SM_CYBORDER   As Long = 6
'#define SM_CXDLGFRAME           7
'#define SM_CYDLGFRAME           8
'#define SM_CYVTHUMB             9
'#define SM_CXHTHUMB             10
'#define SM_CXICON               11
'#define SM_CYICON               12
'#define SM_CXCURSOR             13
'#define SM_CYCURSOR             14
Public Const SM_CYMENU   As Long = 15
'#define SM_CXFULLSCREEN         16
'#define SM_CYFULLSCREEN         17
'#define SM_CYKANJIWINDOW        18
'#define SM_MOUSEPRESENT         19
'#define SM_CYVSCROLL            20
'#define SM_CXHSCROLL            21
'#define SM_DEBUG                22
'#define SM_SWAPBUTTON           23
'#define SM_RESERVED1            24
'#define SM_RESERVED2            25
'#define SM_RESERVED3            26
'#define SM_RESERVED4            27
'#define SM_CXMIN                28
'#define SM_CYMIN                29
'#define SM_CXSIZE               30
'#define SM_CYSIZE               31
'#define SM_CXFRAME              32
'#define SM_CYFRAME              33
'#define SM_CXMINTRACK           34
'#define SM_CYMINTRACK           35
'#define SM_CXDOUBLECLK          36
'#define SM_CYDOUBLECLK          37
'#define SM_CXICONSPACING        38
'#define SM_CYICONSPACING        39
'#define SM_MENUDROPALIGNMENT    40
'#define SM_PENWINDOWS           41
'#define SM_DBCSENABLED          42
'#define SM_CMOUSEBUTTONS        43
'#define SM_CXFIXEDFRAME           SM_CXDLGFRAME  /* ;win40 name change */
'#define SM_CYFIXEDFRAME           SM_CYDLGFRAME  /* ;win40 name change */
'#define SM_CXSIZEFRAME            SM_CXFRAME     /* ;win40 name change */
'#define SM_CYSIZEFRAME            SM_CYFRAME     /* ;win40 name change */
'#define SM_SECURE               44
Public Const SM_CXEDGE     As Long = 45
Public Const SM_CYEDGE    As Long = 46
'#define SM_CXMINSPACING         47
'#define SM_CYMINSPACING         48
'#define SM_CXSMICON             49
'#define SM_CYSMICON             50
'#define SM_CYSMCAPTION          51
'#define SM_CXSMSIZE             52
'#define SM_CYSMSIZE             53
'#define SM_CXMENUSIZE           54
'#define SM_CYMENUSIZE           55
'#define SM_ARRANGE              56
'#define SM_CXMINIMIZED          57
'#define SM_CYMINIMIZED          58
'#define SM_CXMAXTRACK           59
'#define SM_CYMAXTRACK           60
'#define SM_CXMAXIMIZED          61
'#define SM_CYMAXIMIZED          62
'#define SM_NETWORK              63
'#define SM_CLEANBOOT            67
'#define SM_CXDRAG               68
'#define SM_CYDRAG               69
'#define SM_SHOWSOUNDS           70
'#define SM_CXMENUCHECK          71   /* Use instead of GetMenuCheckMarkDimensions()! */
'#define SM_CYMENUCHECK          72
'#define SM_SLOWMACHINE          73
'#define SM_MIDEASTENABLED       74
'#define SM_MOUSEWHEELPRESENT    75
'#define SM_CMETRICS             76
Declare Function ShellAbout Lib "shell32" Alias "ShellAboutA" _
                            (ByVal hWnd As Long, _
                            ByVal szApp As String, _
                            ByVal szOtherStuff As String, _
                            ByVal hIcon As Long) As Long
Declare Function SetWindowTextAsLong Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal LPCSTR As Long) As Long ' C BOOL
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long 'C BOOL
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
                            (ByVal lpRootPathName As String, _
                            lpSectorsPerCluster As Long, _
                            lpBytesPerSector As Long, _
                            lpNumberOfFreeClusters As Long, _
                            lpTtoalNumberOfClusters As Long) As Long 'C BOOL

'== Global Memory Functions ==================================================
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpStringDest As Long, ByVal lpStringSrc As Long) As Long
Declare Sub CopyPTRtoANY Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Dest As Any, ByVal PtrSrc As Long, ByVal length As Long)
Declare Sub CopyPTRtoLONG Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef LONGDest As Long, ByVal PtrSrc As Long, ByVal length As Long)
'Declare Sub CopyPTRtoSTRUCT Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef STRUCTDest As Any, ByVal PtrSrc As Long, ByVal length As Long)
'Declare Sub CopyBYTEStoPTR Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByVal PtrDest As Long, ByRef ByteSrc As Byte, ByVal length As Long)

Public Const GMEM_MOVEABLE = &H2&
Public Const GMEM_SHARE = &H2000&
Public Const GMEM_ZEROINIT = &H40&
