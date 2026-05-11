VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Screen-Shot Logger"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   14085
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3180
      Top             =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FOLDERNAME
Dim TS As String, QQ, Txw$, ii, FF$, XX, YY, FS

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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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

Option Explicit

Private Sub Form_Load()
If App.PrevInstance = True Then
    End
End If

Set FS = CreateObject("Scripting.FileSystemObject")


'If FS.DRIVEEXISTS("m") = False Then End

'Dim d
'Set d = FS.GetDrive("D:\")
'If d.ISREADY = False Then End


ChDir App.Path

'set up the visual stuff

XX = Screen.Width / Screen.TwipsPerPixelX
YY = Screen.Height / Screen.TwipsPerPixelY
'XX = Screen.Width
'YY = Screen.Height

Form1.Width = Screen.Width
Form1.Height = Screen.Height
'Picture1.Width = Form1.Width
'Picture1.Height = Form1.Height
'Picture1.Left = 0
'Picture1.Top = 0

'Inten is the measure of how many pixels are going to be recognized. I highly dont recommend
'setting it lower than this, i have a 3.0 GHz PC and it starts to lag a little. On this setting,
'every 15th pixel is checked
inten = 10
'The tolerance of recognizing the pixel change
Tolerance = 10

Tppx = Screen.TwipsPerPixelX
Tppy = Screen.TwipsPerPixelY

ReDim POn(XX / inten, YY / inten)
ReDim P(XX / inten, YY / inten)


If IsIDE = True Then Form1.Show

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

'Exit Sub




'Dim Ip
'On Error Resume Next
'Set Ip = Image
'Set Ip = LoadPicture(App.Path + "\VBDataNoArchive\Screen-Shot ArrayPic.jpg")
'Set Form1.Picture = Ip
'
'On Error GoTo 0
'Call ComParePic
'Set Ip = Nothing

DoEvents
TT = PrintScreenOntoForm(Form1)
DoEvents

'FileInUseDelay App.Path + "\VBDataNoArchive\Screen-Shot ArrayPic2.jpg"
'SavePicture Form1.Picture, App.Path + "\VBDataNoArchive\Screen-Shot ArrayPic2.jpg"

'Set Ip = Image
'Set Ip = LoadPicture(App.Path + "\VBDataNoArchive\Screen-Shot ArrayPic2.jpg")
'Set Form1.Picture = Ip

'Call ComParePic

FF$ = "ScreenCapture_" + Format$(Now, "YYYY-MM-DD-DDD")
On Error Resume Next
FOLDERNAME = "D:\0 00 ART LOGGERS\# APP AND SCREEN -- SHOT\Screen-Shots\" + FF$

If FS.folDerexists("D:\0 00 ART LOGGERS") = False Then
    MkDir "D:\0 00 ART LOGGERS"
End If
If FS.folDerexists("D:\0 00 ART LOGGERS\# APP AND SCREEN -- SHOT") = False Then
    MkDir "D:\0 00 ART LOGGERS\# APP AND SCREEN -- SHOT"
End If
If FS.folDerexists("D:\0 00 ART LOGGERS\# APP AND SCREEN -- SHOT\Screen-Shots") = False Then
    MkDir "D:\0 00 ART LOGGERS\# APP AND SCREEN -- SHOT\Screen-Shots"
End If
If FS.folDerexists(FOLDERNAME) = False Then
    MkDir FOLDERNAME
End If

If Err.Number > 0 Then MsgBox "PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME '= 76 Then End
'On Error GoTo 0

'FileInUseDelay App.Path + "\VBDataNoArchive\Screen-Shot Pic Always.jpg"
'SavePicture Form1.Picture, App.Path + "\VBDataNoArchive\Screen-Shot Pic Always.jpg"

'FileInUseDelay App.Path + "\VBDataNoArchive\Screen-Shot % Pic Data.txt"
'Open App.Path + "\VBDataNoArchive\Screen-Shot % Pic Data.txt" For Append As #1
''Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + " - " + Format$(Wo / (Ri + Wo) * 100, "0") & "%"
'Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
'Close #1

'FileInUseDelay App.Path + "\VBDataNoArchive\WebCam % Pic Data2.txt"
'Open App.Path + "\VBDataNoArchive\WebCam % Pic Data2.txt" For Output As #1
'Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
'Close #1
'If Wo / (Ri + Wo) * 100 >= 2 Then
    
    'FileInUseDelay App.Path + "\VBDataNoArchive\Screen-Shot ArrayPic.jpg"
    'SavePicture Form1.Picture, App.Path + "\VBDataNoArchive\Screen-Shot ArrayPic.jpg"
    
    'TS = "M\0 00 Art Loggers\Screen-Shots\" + FF$ + "\ScreenCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + " - " + Format$(Wo / (Ri + Wo) * 100, "0") & "%.jpg"
TS = FOLDERNAME + "\ScreenCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
    'FileInUseDelay TS
SavePicture Form1.Picture, TS
'End If

Timer1.Enabled = False


'If IsIDE = False Then Unload Form1

'Exit Sub

Exit Sub
Errcode2:
DoEvents
Resume Next

End Sub


Sub ComParePic()
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




