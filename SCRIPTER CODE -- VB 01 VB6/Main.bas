Attribute VB_Name = "Main"
Option Explicit

Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent _
   As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId _
   As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
   (ByVal hWnd As Long, ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
   (ByVal hWnd As Long, ByVal lpString As String, _
   ByVal cch As Long) As Long

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
    As Long
    
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As _
    Long, ByVal lParam As Long) As Long

Public Declare Function WindowFromPointXY Lib "user32" _
               Alias "WindowFromPoint" (ByVal xPoint As Long, _
               ByVal yPoint As Long) As Long
               
Private Declare Function SystemParametersInfo Lib "user32" _
        Alias "SystemParametersInfoA" _
        (ByVal uAction As Long, _
        ByVal uParam As Long, _
        lpvParam As Any, _
        ByVal fuWinIni As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (pt As POINTAPI) As Long
Public Declare Function GetWindowInfo Lib "user32" (ByVal hWnd As Long, ByRef pwi As WINDOWINFO) As Boolean

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" Alias "FreeLibraryA" (ByVal hLibrary As Long) As Boolean


Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type WINDOWINFO
    cbSize As Long
    rcWindow As RECT
    rcClient As RECT
    dwStyle As Long
    dwExStyle As Long
    cxWindowBorders As Long
    cyWindowBorders As Long
    atomWindowtype As Long
    wCreatorVersion As Long
End Type

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type MOUSEHOOKSTRUCT
  pt As POINTAPI
  hWnd As Long
  wHitTestCode As Long
  dwExtraInfo As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205

Private Const MK_LBUTTON = &H1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = &H2

Public Const WH_MOUSE = 7
Private Const WHEEL_DELTA = 120

Private Const WM_VSCROLL = &H115
Private Const WM_USER As Long = &H400
Private Const WM_SOMETHING = WM_USER + 3139

Public Const GWL_WNDPROC = -4
Public Const WH_MOUSE_LL = 14

Public Const SB_LINEUP = 0
Public Const SB_LINELEFT = 0
Public Const SB_LINEDOWN = 1
Public Const SB_LINERIGHT = 1
Public Const SB_ENDSCROLL = 8
Public Const WS_VISIBLE = &H10000000
Public Const SBS_VERT = 1
Public Const SBS_HORZ = 0
Public Const WM_HSCROLL = &H114
Public Const SPI_GETWHEELSCROLLLINES = 104

Public Enum mButtons
  LBUTTON = &H1
  MBUTTON = &H10
  RBUTTON = &H2
End Enum

   Public Const REG_SZ As Long = 1
   Public Const REG_DWORD As Long = 4

   Public Const HKEY_CLASSES_ROOT = &H80000000
   Public Const HKEY_CURRENT_USER = &H80000001
   Public Const HKEY_LOCAL_MACHINE = &H80000002
   Public Const HKEY_USERS = &H80000003

   Public Const ERROR_NONE = 0
   Public Const ERROR_BADDB = 1
   Public Const ERROR_BADKEY = 2
   Public Const ERROR_CANTOPEN = 3
   Public Const ERROR_CANTREAD = 4
   Public Const ERROR_CANTWRITE = 5
   Public Const ERROR_OUTOFMEMORY = 6
   Public Const ERROR_ARENA_TRASHED = 7
   Public Const ERROR_ACCESS_DENIED = 8
   Public Const ERROR_INVALID_PARAMETERS = 87
   Public Const ERROR_NO_MORE_ITEMS = 259

   Public Const KEY_QUERY_VALUE = &H1
   Public Const KEY_SET_VALUE = &H2
   Public Const KEY_ALL_ACCESS = &H3F

   Public Const REG_OPTION_NON_VOLATILE = 0

   Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
   
   Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
        "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
        As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
        As Long, phkResult As Long, lpdwDisposition As Long) As Long
   
   Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
        "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
        Long) As Long
   
   Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
        As String, lpcbData As Long) As Long
            
   Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, lpData As _
        Long, lpcbData As Long) As Long
   
   Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
        As Long, lpcbData As Long) As Long
   
   Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
        String, ByVal cbData As Long) As Long
    
   Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
       "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
       ByVal cbData As Long) As Long

Dim nKeys As Long, Delta As Long, XPos As Long, YPos As Long
Dim OriginalWindowProc As Long
Dim pthWnd As Long
Dim lLineNumbers As Long
Dim MainWindowHwnd As Long  ' Main IDE window handle
Dim bHook As Boolean
Dim sLib As String
Dim hLib As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) _
                           As Long
    Select Case uMsg
      Case WM_MOUSEWHEEL
        nKeys = wParam And 65535
        Delta = wParam / 65536 / WHEEL_DELTA

        XPos = LowWord(lParam)
        YPos = HighWord(lParam)
        
        pthWnd = WindowFromPointXY(XPos, YPos)
                
        ' Get the scroll bar for this window and send the vscroll to it
        Dim lret As Long
        lret = EnumChildWindows(pthWnd, AddressOf EnumChildProc, lParam)
               
    End Select

    If OriginalWindowProc <> 0 Then
        WindowProc = CallWindowProc(OriginalWindowProc, hWnd, uMsg, wParam, lParam)
    End If
End Function

Public Sub UnHook()

    'Ensures that you don't try to unsubclass the window when
    'it is not subclassed.
    If OriginalWindowProc = 0 Then Exit Sub
    

    'Reset the window's function back to the original address.
    Dim hr As Long
    hr = SetWindowLong(MainWindowHwnd, GWL_WNDPROC, OriginalWindowProc)
    If hr <> 0 Then
        OriginalWindowProc = 0
        bHook = False
    Else
        Debug.Print "Unable to unhook:  SetWindowLong returns " & vbCrLf & hr & vbCrLf & Err.LastDllError
    End If
    
End Sub

Public Sub Hook()
    On Error GoTo Error
    
    ' GetLine Numbers
    SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, lLineNumbers, 0
    
    ' Adjust just in case, otherwise we'll never get the scroll notification.
    If lLineNumbers = 0 Then
        lLineNumbers = 1
    End If
    
    OriginalWindowProc = SetWindowLong(MainWindowHwnd, GWL_WNDPROC, AddressOf WindowProc)
    
    ' Set a flag indicating that we are hooking
    bHook = True
    
    ' Find out where we live on the filesystem
    Dim lRetVal As Long
    Dim sKeyName As String
    Dim sValue As String
    sKeyName = "CLSID\{B84F8C6E-BDDE-4384-9946-82EEE7F81D48}\InprocServer32"
    sValue = QueryValue(sKeyName, "")
    
    ' If we found where we live let's increase our ref count so we can do our own cleanup later
    If Len(sValue) > 0 Then
        sLib = sValue
        hLib = LoadLibrary(sLib)
    End If
        
    Exit Sub
    
Error:
    Debug.Print "Unable to set hook:  " & vbCrLf & Err.Description & vbCrLf & Err.LastDllError
End Sub

Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) _
   As Long
   Dim RetVal As Long
   Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
   Dim WinClass As String, WinTitle As String
   Dim WinRect As RECT
   Dim WinWidth As Long, WinHeight As Long

   RetVal = GetClassName(pthWnd, WinClassBuf, 255)
   WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
   RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
   WinTitle = StripNulls(WinTitleBuf)
   
   ' see the Windows Class and Title for each Child Window enumerated
   'Debug.Print "   hWnd = " & Hex(lhWnd) & " Child Class = "; WinClass; ", Title = "; WinTitle
   ' You can find any type of Window by searching for its WinClass
   Dim lret As Long
   Dim i As Long
   
   ' Since we can have split windows we need to figure out which scroll bar to move.
   ' We can do this by comparing the Y position of the cursor against the vertical scrollbars
   ' that are children of the current window
   Dim wi As WINDOWINFO
   wi.cbSize = Len(wi)
   If GetWindowInfo(lhWnd, wi) And WinClass <> "MDIClient" Then
        If IsVerticalScrollBar(lhWnd) = True And wi.rcWindow.Top < YPos And wi.rcWindow.Bottom > YPos Then    ' TextBox Window
          
             If Delta > 0 Then                       ' Scroll Up
                  Do While i < Delta * lLineNumbers
                     lret = PostMessage(pthWnd, WM_VSCROLL, SB_LINEUP, lhWnd)
                     i = i + 1
                  Loop
              Else                                   ' Scroll Down
                  Do While i > Delta * lLineNumbers
                     lret = PostMessage(pthWnd, WM_VSCROLL, SB_LINEDOWN, lhWnd)
                     i = i - 1
                  Loop
              End If
        ElseIf IsHorizontalScrollBar(lhWnd) = True Then
             If Delta > 0 Then                       ' Scroll Left
                 Do While i < Delta * lLineNumbers
                     lret = PostMessage(pthWnd, WM_HSCROLL, SB_LINELEFT, lhWnd)
                     i = i + 1
                 Loop
              Else                                   ' Scroll Right
                 Do While i > Delta * lLineNumbers
                     lret = PostMessage(pthWnd, WM_HSCROLL, SB_LINERIGHT, lhWnd)
                     i = i - 1
                 Loop
              End If
        End If
   End If
   
   EnumChildProc = bHook                              ' Continue enumerating the windows based on whether we are hooking or not
   
   ' It's possible that the addin has already been requested to unload and in such a case we will call free library on ourselves
   ' to reduce our ref count since we incremented it on our own so we can do a clean shutdown
   If Not bHook Then
        If Not FreeLibrary(hLib) Then
             Debug.Print "Unable to FreeLibrary: " & Err.Number & vbCrLf & Err.LastDllError
        End If
   End If
   
End Function

Function EnumThreadProc(ByVal lhWnd As Long, ByVal lParam As Long) _
   As Long
   Dim RetVal As Long
   Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
   Dim WinClass As String, WinTitle As String

On Error GoTo Error

   RetVal = GetClassName(lhWnd, WinClassBuf, 255)
   WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
   RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
   WinTitle = StripNulls(WinTitleBuf)

   ' see the Windows Class and Title for top level Window
   Debug.Print "Thread Window Class = "; WinClass; ", Title = "; _
   WinTitle
   EnumThreadProc = True
   
   If InStr(1, WinTitle, "Microsoft Visual Basic") <> 0 _
    And WinClass = "wndclass_desked_gsk" _
    And MainWindowHwnd = 0 Then
    
    MainWindowHwnd = lhWnd
    ' Setup the windows Hook
    Hook
   
   End If
   
   Exit Function
Error:
    MsgBox Err.Description
   
End Function

Public Function StripNulls(OriginalStr As String) As String
   ' This removes the extra Nulls so String comparisons will work
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function

Public Function IsVerticalScrollBar(hWnd As Long) As Boolean

    ' Check the style of the window specified by hWnd to see if it's a vertical scrollbar

    Dim wi As WINDOWINFO
    wi.cbSize = Len(wi)
    
    If GetWindowInfo(hWnd, wi) Then
        If (wi.dwStyle And WS_VISIBLE) > 0 And (wi.dwStyle And SBS_VERT) > 0 Then
            IsVerticalScrollBar = True
            Exit Function
        End If
    End If
  
    IsVerticalScrollBar = False

End Function

Public Function IsHorizontalScrollBar(hWnd As Long) As Boolean

    ' Check the style of the window specified by hWnd to see if it's a horizontal scrollbar

    Dim wi As WINDOWINFO
    wi.cbSize = Len(wi)
    
    If GetWindowInfo(hWnd, wi) Then
        If (wi.dwStyle And WS_VISIBLE) > 0 And (wi.dwStyle And SBS_HORZ) > 0 Then
            IsHorizontalScrollBar = True
            Exit Function
        End If
    End If
  
    IsHorizontalScrollBar = False

End Function


Private Function QueryValue(sKeyName As String, sValueName As String) As Variant
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value

    lRetVal = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    RegCloseKey (hKey)
    
    QueryValue = vValue
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
       Dim lValue As Long
       Dim sValue As String
       Select Case lType
           Case REG_SZ
               sValue = vValue & Chr$(0)
               SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
           Case REG_DWORD
               lValue = vValue
               SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
           End Select
End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim sValue As String

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       If lrc <> ERROR_NONE Then Error 5

       Select Case lType
           ' For strings
           Case REG_SZ:
               sValue = String(cch, 0)

   lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
   sValue, cch)
               If lrc = ERROR_NONE Then
                   vValue = Left$(sValue, cch - 1)
               Else
                   vValue = Empty
               End If
           ' For DWORDS
           Case REG_DWORD:
   lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
   lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
           Case Else
               'all other data types not supported
               lrc = -1
       End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function

QueryValueExError:
       Resume QueryValueExExit
End Function

Private Function LowWord(ByVal inDWord As Long) As Integer
    LowWord = inDWord And &H7FFF&
    If (inDWord And &H8000&) Then LowWord = LowWord Or &H8000
End Function

Private Function HighWord(ByVal inDWord As Long) As Integer
    HighWord = LowWord(((inDWord And &HFFFF0000) \ &H10000) And &HFFFF&)
End Function

