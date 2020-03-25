Attribute VB_Name = "Cid_Run_Me_M"
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction&, ByVal uParam&, ByRef lpvParam As Any, ByVal fuWinIni&) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Declare Function ShowOwnedPopups Lib "user32" (ByVal hWnd As Long, ByVal fShow As Long) As Long

Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function closewindow Lib "user32" Alias "CloseWindow" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long



Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
'Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, hModule As Long, ByVal cb As Long, cbNeeded As Long) As Long




Public Declare Function GetModuleFileNameA Lib "kernel32" _
         (ByVal hModule As Long, ByVal lpFileName As String, _
         ByVal nSize As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long




Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


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






Public StopMe As Long
Public Snatch

Public cProcesses As New clsCnProc

Public NoMusic As Long
Public NoMonOff As Long
Public NoMonOff2 As Long
'Public NoMmusicEx As long
'Public NoMonoffEx As long
Public Ebuyer As Long
Public Ebuy As Long
Public Otss As Long
Public Cmdv As Long
Public Tagad As Long
Public CurProcHwnd As Long

Public ProcessId22 As Long
Public ProcessId23 As Long
Public ProcessId24 As Long
Public ProcessId25Ssam As Long

Public VirusActive As Long

Public Ebuyer3() As Long
Public Ebuyer4() As Long

Public GetWinLen() As Long
Public HwndArray() As Long
Public HwndArra2() As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type


Public globalhot$


'Const PROCESS_VM_READ = &H10
'Const PROCESS_QUERY_INFORMATION = &H400

'mattstimer2
Public MajorOnTime As Date
Public MinorOnTime As Date

Public MattsTimer2 As Date






Public Monitor_Timer As Date

Public Cheese As Long



Public sArray() As String
Public iCtr As Long

Public Peet3$
Public Peet33$

Public FlingGrater1$
Public FlingGrater2$

Public CheckT1$

Public OffSetGoogle
Public OldRect23


Public Function GetFileFromHwnd(lngHwnd) As String

'MsgBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function



Public Sub Perfect(Ew1$, LockStop)

Dim dickweed
Dim R
Dim ash$
Dim totss As Boolean
Dim OutputData As Long
Dim totss2 As Long

Snatch = 0

ewd$ = Ew1$
Ew1$ = LCase$(Ew1$)

If InStr(Ew1$, LCase("FS2.exe")) Then
    Snatch = 1
End If

If InStr(Ew1$, LCase("nero.exe")) Then
    NoMonOff = 1
    NoMusic = 1
    Snatch = 1
End If


If InStr(Ew1$, LCase("LeechGet.exe")) Then
    NoMonOff2 = 1
    NoMusic = 1
    Snatch = 1
End If

'''''outlook Express
If InStr(Ew1$, LCase("msimn.exe")) Then
    NoMonOff = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("Del_empty_folders.exe")) Then
    NoMonOff = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("IceView.exe")) Then
    NoMonOff = 1
    Snatch = 1
End If

If InStr(Ew1$, "realplay.exe") Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("CTRec.exe")) Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("PowerDVD.exe")) Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("PowerVCD.exe")) Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If


If InStr(Ew1$, LCase("SeriousSam.exe")) Then
    NoMonOff = 1
    NoMusic = 1
    ProcessId25Ssam = Otss
    Snatch = 1
End If

If InStr(Ew1$, "p Video\winamp.exe") Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, "p beavis\winamp.exe") Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, "p midi\winampexe") Then
    NoMonOff = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("FRONTPG.EXE")) Then
    NoMonOff = 1
    frontpagepid = Otss
    Snatch = 1
End If

If InStr(Ew1$, "cmd.exe") Then
    NoMonOff = 1
    Cmdv = 1
    Snatch = 1
End If

If InStr(Ew1$, "iexplore.exe") Then
    NoMonOff = 1: NoMusic = 1
    Ebuy = 1
    Snatch = 1
End If

If InStr(LCase(Ew1$), "amp caller\winamp.exe") Then
    ProcessId22 = Otss
    Snatch = 1
End If
                
If InStr(LCase(Ew1$), "amp gold\winamp.exe") Then
    ProcessId24 = Otss
    Snatch = 1
End If

If Snatch = 1 And LockStop = 0 Then
    'dickweed = 0
   
    
    
    If InStr(Peet33$, str(Otss)) = 0 Then
        'totss = cProcesses.Convert(CurProcHwnd, Otss, cnFromhWnd Or cnToProcessID)
        totss = cProcesses.Convert(Otss, totss2, cnTohWnd Or cnFromProcessID)
        CID_Run_Me.List2.AddItem Format$(totss2, "000000000") + " " + Format$(Otss, "0000000") + " "
        Peet33$ = Peet33$ + str(Otss)
        Peet3$ = Peet3$ + str(totss2)
    End If

    
    
    
    'totss = cProcesses.Convert(Otss, totss2, cnTohWnd Or cnFromProcessID)
                            
    'dickweed = 0
    
    'For r = 0 To Tagad
    '    If totss2 = HwndArray(r) Then
    '        dickweed = 1
    '        Exit For
    '    End If
    'Next

    'If dickweed = 0 Then
    '    For r = 0 To Tagad
    '        If HwndArray(r) < 0 Then
    '            HwndArray(r) = totss2
    '            HwndArra2(r) = Otss
    '            Tagad = Tagad + 1
    '            Exit For
    '
    '        End If
    '    Next
    'End If
End If

Ew1$ = ewd$
End Sub



