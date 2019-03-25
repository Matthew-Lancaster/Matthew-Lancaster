Attribute VB_Name = "APublics"
Public FR1

Public m_CRC As clsCRC
Public WxHex$


Public FSO

Public TEXT_PATH_1
Public TEXT_PATH_2
Public TEXT_PATH_4

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public FontSizez, SetTrueToLoadLast, LoadFolder

Public A1$, B1$, C1$, OIP$, OIP2$

Public A4$(), B4$(), C4$()

Public FS

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long

Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type


Public cProcesses As New clsCnProc

Public Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

' Public constants for GetWindowLong API declaration
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

' Public constants for window styles
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WM_CLOSE = &H10

' Private constants for extended window styles
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000

' Public constants for SetWindowPos API declaration
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

' Public constants for ShowWindow API declaration
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5

' Public constants for set window on top
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

' Public constants for CreateToolhelp32Snapshot API declaration
Private Const TH32CS_SNAPPROCESS = &H2&

Private m_sPattern As String
Private m_lhFind As Long



Sub SET_UP_PULIC_FSO()

Set FSO = CreateObject("Scripting.FileSystemObject")

End Sub


'Public Function CreateFolderTree(ByVal sPath As String) As Boolean
'    Dim nPos As Integer
'
'    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)
'
'    On Error GoTo CreateFolderTreeError
'
'    nPos = InStr(sPath, "\")
'    While nPos > 0
'        If Not FolderExists(Left$(sPath, nPos - 1)) Then
'            MkDir Left$(sPath, nPos - 1)
'        End If
'        nPos = InStr(nPos + 1, sPath, "\")
'    Wend
'    If Not FolderExists(sPath) Then MkDir sPath
'
'    CreateFolderTree = True
'    Exit Function
'
'CreateFolderTreeError:
'    Exit Function
'End Function

Public Function FolderExists(sFolder As String) As Boolean
    '##############################################################################################
    'Returns True if the specified folder exists
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function







