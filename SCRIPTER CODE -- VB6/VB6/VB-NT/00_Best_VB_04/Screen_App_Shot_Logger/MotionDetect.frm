VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Screen-App-Shot Logger"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   100
      Top             =   100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filexxx$, CurProcHwnd, TTT
Dim Rect1 As RECT

Private Declare Function GetParent _
        Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint _
                 Lib "user32" (ByVal xPoint As Long, _
                               ByVal yPoint As Long) As Long


Dim TS1 As String, TS2 As String, TS3 As String, QQ, Txw$, ii, FF$, XX, YY
Private Declare Function GetForegroundWindow Lib "user32" () As Long

'Public cProcesses As New clsCnProc
    '#### Functions/Consts used for GetHWnd() (part of Convert())
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long  'MODULE 1141
'Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Dim TJax, GJax, ET, TBak, TY, Tx, HH

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054

'declarations
Dim P() As Long
Dim POn() As Boolean

Dim inten As Integer

Dim i As Integer, j As Integer

Dim Ri As Long, Wo As Long
Dim RealRi As Long

Dim C As Long, c2 As Long

Dim R As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

Dim Tppx As Single, Tppy As Single
Dim Tolerance As Integer

Dim RealMov As Integer

Dim Counter As Integer

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim LastTime As Long, TT

'Option Explicit

Private Sub Form_Load()

If App.PrevInstance = True Then
    End
End If

Set FS = CreateObject("Scripting.FileSystemObject")

If FS.DRIVEEXISTS("m") = False Then End
Dim d
Set d = FS.GetDrive("M:\")
If d.ISREADY = False Then End

On Error Resume Next

ChDir App.Path

'set up the visual stuff
'Picture1.Width = 640 * Screen.TwipsPerPixelX
'Picture1.Height = 480 * Screen.TwipsPerPixelY
'Picture1.Left = 0
'Picture1.Top = 0

' = cProcesses.Convert(Otss, totss2, cnTohWnd Or cnFromProcessID)


CurProcHwnd = GetParentHwnd(GetForegroundWindow)
Filexxx$ = GetFileFromHwnd(CurProcHwnd)

Hwnd9 = GetWindowRect(CurProcHwnd, Rect1)

TTT = Mid$(Format$(Now, "HH:MM"), 1, 4)
pri = pri + TTT
pri = pri + Str(CurProcHwnd)
pri = pri + GetWindowTitle(CurProcHwnd)
pri = pri + GetWindowClass(CurProcHwnd)
pri = pri + Filexxx$
pri = pri + Str(Rect1.Bottom) + Str(Rect1.Left) + Str(Rect1.Right) + Str(Rect1.Top)
Dim Mwnd As Long
FindCursor Nx, Ny
Mwnd = WindowFromPoint(Nx, Ny)
hallo = GetWindowTitle(Mwnd)

hallo = GetParentWindowTitle(Mwnd)
    If InStr(hallo, "JPEG Encoder") > 0 Then
    End
End If

hallo = GetWindowTitle(CurProcHwnd)
    If InStr(hallo, "JPEG Encoder") > 0 Then
    End
End If

TS3 = App.Path + "\VBDataNoArchive\Screen-App-Shot ArrayPic.txt"
Open TS3 For Input As #1
rr = Input(LOF(1), 1)
Close #1

If rr = pri Then End


'XX = Screen.Width / Screen.TwipsPerPixelX
'YY = Screen.Height / Screen.TwipsPerPixelY

R = Rect1.Right
L = Rect1.Left
B = Rect1.Bottom
T = Rect1.Top
If L < 0 Then L = 0
If T < 0 Then T = 0
Form1.Height = ((B - T) * Screen.TwipsPerPixelX) + 420
Form1.Width = ((R - L) * Screen.TwipsPerPixelY) + 60

'Form1.Height = (Rect1.Bottom - Rect1.Top)
'Form1.Width = (Rect1.Right - Rect1.Left)
Refresh
DoEvents

XX = Form1.Width / Screen.TwipsPerPixelX
YY = Form1.Height / Screen.TwipsPerPixelY

'Inten is the measure of how many pixels are going to be recognized. I highly dont recommend
'setting it lower than this, i have a 3.0 GHz PC and it starts to lag a little. On this setting,
'every 15th pixel is checked
inten = 2
'The tolerance of recognizing the pixel change
Tolerance = 1

Tppx = Screen.TwipsPerPixelX
Tppy = Screen.TwipsPerPixelY

ReDim POn(XX / inten, YY / inten)
ReDim P(XX / inten, YY / inten)


If IsIDE = True Then Form1.Show

If Err.Number > 0 Then End

Timer1.Enabled = True

End Sub


Private Sub Timer1_Timer()

'Exit Sub

Dim Ip
On Error Resume Next
'Set Ip = Image
'Set Ip = LoadPicture(App.Path + "\VBDataNoArchive\Screen-App-Shot ArrayPic.jpg")
'Set Form1.Picture = Ip
'Refresh
'DoEvents
'On Error GoTo 0
'Call ComParePic(XX, YY)
'Form1.Picture = Nothing

TT = PrintScreenOntoFormApp(Form1, Rect1.Bottom, Rect1.Left, Rect1.Right, Rect1.Top)
Form1.Refresh
Refresh
DoEvents

'Call ComParePic(XX, YY)


FF$ = "ScreenAppCapture_" + Format$(Now, "YYYY-MM-DD-DDD")
On Error Resume Next
MkDir "M\0 00 Art Loggers\Screen-App-Shots\" + FF$
If Err.Number = 76 Then End
On Error GoTo 0
FileInUseDelay App.Path + "\VBDataNoArchive\Screen-App-Shot Pic Always.jpg"
SavePicture Form1.Picture, App.Path + "\VBDataNoArchive\Screen-App-Shot Pic Always.jpg"
FileInUseDelay App.Path + "\VBDataNoArchive\Screen-App-Shot % Pic Data.txt"
Open App.Path + "\VBDataNoArchive\Screen-App-Shot % Pic Data.txt" For Append As #1
'Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + " - " + Format$(Wo / (Ri + Wo) * 100, "0") & "%"
Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
Close #1
'If Wo / (Ri + Wo) * 100 >= 1 Then
    FileInUseDelay App.Path + "\VBDataNoArchive\Screen-App-Shot ArrayPic.jpg"
    SavePicture Form1.Picture, App.Path + "\VBDataNoArchive\Screen-App-Shot ArrayPic.jpg"
    'TS1 = "M\0 00 Art Loggers\Screen-App-Shots\" + FF$ + "\ScreenAppCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + " - " + Format$(Wo / (Ri + Wo) * 100, "0") & "%.jpg"
    TS1 = "M\0 00 Art Loggers\Screen-App-Shots\" + FF$ + "\ScreenAppCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
    FileInUseDelay TS1
    SavePicture Form1.Picture, TS1
    'TS2 = "M\0 00 Art Loggers\Screen-App-Shots\" + FF$ + "\ScreenAppCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + " - " + Format$(Wo / (Ri + Wo) * 100, "0") & "%.txt"
    TS2 = "M\0 00 Art Loggers\Screen-App-Shots\" + FF$ + "\ScreenAppCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".txt"
    FileInUseDelay TS2
    Open TS2 For Output As #1
    Print #1, "---------------------"
    Print #1, "Screen Application Shot FileName ="
    Print #1, TS1
    Print #1, "---------------------"
    Print #1, "Current Handle Hwnd ="; CurProcHwnd
    Print #1, "---------------------"
    Print #1, "Application Title = "; GetWindowTitle(CurProcHwnd)
    Print #1, "---------------------"
    Print #1, "Application Class Title = "; GetWindowClass(CurProcHwnd)
    Print #1, "---------------------"
    Print #1, "Application Exe = "; Filexxx$
    Print #1, "---------------------"
    Print #1, "Resolution -----"
    Print #1, "---------------------"
    Print #1, "Top ="; Rect1.Top
    Print #1, "Bottom ="; Rect1.Bottom
    Print #1, "Left ="; Rect1.Left
    Print #1, "Right ="; Rect1.Right
    Print #1, "---------------------"
    Print #1, "Width ="; Rect1.Right - Rect1.Left
    Print #1, "Height ="; Rect1.Bottom - Rect1.Top
    Print #1, "---------------------"
    Close #1
    
    Hwnd9 = GetWindowRect(CurProcHwnd, Rect1)
    pri = ""
    pri = pri + TTT
    pri = pri + Str(CurProcHwnd)
    pri = pri + GetWindowTitle(CurProcHwnd)
    pri = pri + GetWindowClass(CurProcHwnd)
    pri = pri + Filexxx$
    pri = pri + Str(Rect1.Bottom) + Str(Rect1.Left) + Str(Rect1.Right) + Str(Rect1.Top)
    
    TS3 = App.Path + "\VBDataNoArchive\Screen-App-Shot ArrayPic.txt"
    FileInUseDelay TS3
    Open TS3 For Output As #1
    Print #1, pri;
    Close #1



'End If

Timer1.Enabled = False

If IsIDE = False Then Unload Form1

'Unload Form1

'Exit Sub

Exit Sub
Errcode2:
DoEvents
Resume Next

End Sub


Sub ComParePic(XX, YY)
On Error Resume Next
Ri = 0 'right
Wo = 0 'wrong

'LastTime = GetTickCount

For i = 0 To XX / inten - 1
    For j = 0 To YY / inten - 1
    'get a point
    C = Form1.POINT(i * inten * Tppx, j * inten * Tppy)
    'analyze it, Red, Green, Blue
        R = C Mod 256
        G = (C \ 256) Mod 256
        B = (C \ 256 \ 256) Mod 256
        
    'recall what the point was one step before this
    c2 = P(i, j)
        'analyze it
        R2 = c2 Mod 256
        G2 = (c2 \ 256) Mod 256
        B2 = (c2 \ 256 \ 256) Mod 256
        
    'main comparison part... if each R, G and B are somewhat same, then it pixel is same still
    'in a perfect camera and software tolerance should theoretically be 1 but this isnt true...
    If Abs(R - R2) < Tolerance And Abs(G - G2) < Tolerance And Abs(B - B2) < Tolerance Then
    'pixel remained same
    Ri = Ri + 1
    'Pon stores a boolean if the pixel changed or didnt, to be used to detect REAL movement
    POn(i, j) = True
    
    Else
    'Pixel changed
    Wo = Wo + 1
    'make a red dot
    P(i, j) = Form1.POINT(i * inten * Tppx, j * inten * Tppy)
    Form1.PSet (i * inten * Tppx, j * inten * Tppy), vbRed
    POn(i, j) = False
    End If
    
    Next j
    
Next i

RealRi = 0

For i = 1 To XX / inten - 2
    For j = 1 To YY / inten - 2
    If POn(i, j) = False Then
        'Real movement is simply occuring when all 4 pixels around one pixel changed
        'Simply put, If this pixel is changed and all around it changed too, then this is a real
        'movement
        If POn(i, j + 1) = False Then
            If POn(i, j - 1) = False Then
                If POn(i + 1, j) = False Then
                    If POn(i - 1, j) = False Then
                    RealRi = RealRi + 1
                    Form1.PSet (i * inten * Tppx, j * inten * Tppy), vbGreen
                    End If
                End If
            End If
        End If
        
    End If
        
        
    Next j
Next i

End Sub


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************


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

Sub FileInUseDelay(Txw$)
        
    QQ = Now + TimeSerial(0, 1, 30)
    Do
'        DoEvents
        ii = FileInUse(Txw$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or QQ < Now
    
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Txw$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub




Public Function GetFileFromHwnd(lngHwnd) As String

'MsgBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameExA(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function


Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hwnd)
   S = Space(L + 1)
   GetWindowText hwnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function


Function GetWindowClass(ByVal hwnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hwnd, sText, 255)
    sText = Left$(sText, Ret)
    GetWindowClass = sText
End Function

Public Sub FindCursor(x, Y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
Y = P.Y '/ GetSystemMetrics(1) * Screen.Height

End Sub

Public Function GetParentWindowTitle(Test_Hwnd5 As Long) As String
   Dim i As Long
   Dim j As Long
   i = Test_Hwnd5
   If i Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   i = j
   End If
    GetParentWindowTitle = GetWindowTitle(i)
    Test_Hwnd5 = i
End Function

Public Function GetParentHwnd(Test_Hwnd5 As Long) As String
   Dim i As Long
   Dim j As Long
   i = Test_Hwnd5
   If i Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
    i = j
    End If
    GetParentHwnd = i
End Function


