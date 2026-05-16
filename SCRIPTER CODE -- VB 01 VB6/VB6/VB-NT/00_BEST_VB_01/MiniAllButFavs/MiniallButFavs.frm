VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mini All But Favs"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "MiniallButFavs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetParent _
        Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function SetParent _
        Lib "user32" _
        (ByVal hWndChild As Long, _
         ByVal hWndNewParent As Long) As Long

Dim TT()

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3

Enum BOOL
   FALSE_ = 0
   TRUE_ = 1
End Enum

Const GW_CHILD = 5
'Const GW_HWNDNEXT = 2

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As BOOL
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal strTempFilePath As String, ByVal strFilePrefix As String, ByVal lngUniqueValue As Long, ByVal ReturnBuffer As String) As Long
'Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal lngAccessType As Long, ByVal blnInheritHandle As BOOL, ByVal lngProcessID As Long) As Long

Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Sub Form_Load()

If App.PrevInstance = True Then End

a = FindWinPartFront

End

End Sub

Function FindWinPartFront() As Long
FindWinPartFront = False
ReDim TT(100)

Dim ash$
Dim Test_Hwnd As Long
    'Test_pid As Long, _
    'Test_thread_id As Long

Dim cText As String, FileSpec As String
Dim Rect8 As RECT
Dim WState
Dim VState
Dim ParentHwnd As Long

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
Test_Hwnd = FindWindow2(ByVal 0&, ByVal 0&)
br$ = ""

        sx = 0
        sx = sx + 1: TT(sx) = "\VB6.EXE"
        sx = sx + 1: TT(sx) = "\Cid-RunMe.exe"
        sx = sx + 1: TT(sx) = "\VolumeBar.exe"
        sx = sx + 1: TT(sx) = "\WMICPU.exe"
        sx = sx + 1: TT(sx) = "\Tidal.exe"
        sx = sx + 1: TT(sx) = "\BBCWeather.exe"
        sx = sx + 1: TT(sx) = "\Winamp.exe"
        sx = sx + 1: TT(sx) = "\RecordGo.exe"
        sx = sx + 1: TT(sx) = "MiniAllButFavs.exe"
'        sx = sx + 1: TT(sx) = "\Explorer.EXE"
        sx = sx + 1: TT(sx) = "SONICS"
        sx = sx + 1: TT(sx) = "URL Logger.exe"
        sx = sx + 1: TT(sx) = "Drive_Detach.exe"
        sx = sx + 1: TT(sx) = "ClipBoard Logger.exe"
        sx = sx + 1: TT(sx) = ""
        sx = sx + 1: TT(sx) = ""
        sx = sx + 1: TT(sx) = ""
        sx = sx + 1: TT(sx) = ""
'        sx = sx + 1: TT(sx) = "\Explorer.EXE"
        
        
        For r = 1 To 50
            If TT(r) = "" Then
                ReDim Preserve TT(r - 1): Exit For
            End If
        Next

Do While Test_Hwnd <> 0
    
    hwnd9 = GetWindowRect(Test_Hwnd, Rect8)
    gws = GetWindowState(Test_Hwnd)
    If gws = -1 Then gws2$ = "Window State Normal"
    If gws = vbMinimized Then gws2$ = "Window State Minimized"
    If gws = vbMaximized Then gws2$ = "Window State Maximized"
 
    If gws = -1 Or gws = vbMaximized Then WState = True Else WState = False
    
    'WState = True
    'VState = True
    If (Rect8.Top > 0 Or Rect8.Left > 0) And WState = True Then
        ash$ = GetWindowTitle(Test_Hwnd)
        ash2$ = GetParentTitle(Test_Hwnd)
        ParentHwnd = GetParentHwnd(Test_Hwnd)
        xzag = 0
        If ash$ <> "" Or ash2$ <> "" Then xzag = 1
        If Test_Hwnd <> ParentHwnd Then
            xzag = 1
        End If
        VState = False
        If Rect8.Top > 0 And xzag = 1 Then VState = WindowVisible.WindowVisible(Test_Hwnd)
        If VState = False Then xzag = 0
        
        FileSpec = GetFileFromHwnd(Test_Hwnd)
        
        If InStr(FileSpec, TT(1)) > 0 And _
        Test_Hwnd = Me.hwnd Or ParentHwnd = Me.hwnd Then: 'xzag = 0
        
        'If InStr(FileSpec, TT(1)) > 0 Then Stop
        'If InStr(LCase(FileSpec), LCase("Explorer.EXE")) > 0 And xzag = 1 Then Stop
        
        If InStr(LCase(FileSpec), LCase("WinDowse")) > 0 Then
            'Stop
        End If
        If InStr(LCase(FileSpec), LCase("ViceVersa Pro 2")) > 0 And xzag = 1 Then
            'Stop
            'hh = GetLastHwnd(Test_Hwnd)
        End If
        
        
        'xzag1 = xzag
        xzag2 = 0
        For r = 2 To UBound(TT)
            If InStr(LCase(FileSpec), LCase(TT(r))) > 0 Then
                xzag2 = 1: Exit For
            End If
            'yy$ = Mid$(FileSpec, InStrRev(FileSpec, "\"))
            'Clipboard.Clear
            'Clipboard.SetText (yy$)
        Next
        
        If xzag = 1 And xzag2 = 0 Then
            'yy$ = Mid$(FileSpec, InStrRev(FileSpec, "\"))
            'Clipboard.Clear
            'Clipboard.SetText (yy$)
            
            If ash$ = "Recording Window" Then
                WindowVisible.WindowVisible(Test_Hwnd) = False: xzag = 0
            End If
            
            'If InStr(LCase(FileSpec), LCase("ViceVersa Pro 2")) > 0 Then
            '    If Test_Hwnd = ParentHwnd Then
        '    '        xzag = 0
             '   End If
            'End If
            
            If xzag = 1 Then
                'ParentHwnd = GetParentHwnd(Test_Hwnd)
                'If InStr(LCase(FileSpec), LCase("ClipBoard Logger.exe")) Then Stop
                ShowWindow ParentHwnd, SW_MINIMIZE
                If WindowVisible.WindowVisible(ParentHwnd) = True Then
                    ShowWindow ParentHwnd, SW_FORCEMINIMIZE = 11
                End If
                If WindowVisible.WindowVisible(Test_Hwnd) = True Then
                    ShowWindow Test_Hwnd, SW_MINIMIZE
                End If
                If WindowVisible.WindowVisible(Test_Hwnd) = True Then
                    ShowWindow Test_Hwnd, SW_FORCEMINIMIZE = 11
                End If
            End If
        End If
    End If
        
    'retrieve the next window
    Test_Hwnd = GetWindow(Test_Hwnd, GW_HWNDNEXT)

Loop
'MsgBox Str(huge) + " Windows Brought To Front"
FindWinPartFront = True

End Function


Function GetParentTitle(TestHwnd As Long) As String
   Dim i As Long
   Dim j As Long
   i = ReturnParent
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   i = j
   End If
   GetParentTitle = GetWindowTitle(i)
End Function


Function GetParentHwnd(ReturnParent As Long) As String
   Dim i As Long
   Dim j As Long
   i = ReturnParent
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   i = j
   End If
   GetParentHwnd = i
End Function

Function GetLastHwnd(ReturnParent As Long) As String
   Dim i As Long
   Dim j As Long
   i = ReturnParent
   If ReturnParent Then
      Do While i <> 0
         j = i
         'i = GetParent(i)
         i = GetWindow(i, GW_HWNDNEXT)
        '3475466
      Loop
   i = j
   End If
   GetLastHwnd = i
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

Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function

Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
    If IsIconic(lngHwnd) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lngHwnd) <> 0 Then
        GetWindowState = vbMaximized
    End If
    End If
End Function


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function



