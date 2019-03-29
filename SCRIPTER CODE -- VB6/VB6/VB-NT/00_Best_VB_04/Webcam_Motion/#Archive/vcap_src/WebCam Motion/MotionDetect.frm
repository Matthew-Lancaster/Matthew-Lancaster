VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "WebCam Motion Capture"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   1350
      Left            =   10545
      ScaleHeight     =   1290
      ScaleWidth      =   1050
      TabIndex        =   4
      Top             =   2190
      Width           =   1110
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      DrawWidth       =   3
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6915
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.ComboBox CmboSource 
         Height          =   315
         Left            =   2130
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   660
         Visible         =   0   'False
         Width           =   1470
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10395
      Top             =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "left click to record, right click to stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   8280
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Firstly id like to say that im no expert with webcams, i only know how to capture and stop it
'The major aim of this code, for me, was to make the motion detection algorithm itself

'FOR WEBCAM DECLARATIONS

Dim TS As String, QQ, Txw$, ii, WhiteFactor2, Cat, WhiteFactor, Oi, TYX

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
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long

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
Dim LastTime As Long

Option Explicit

Private Sub Form_Load()
'Open "Q:\Temp\WebCamFlag.txt" For Output Lock Write As #1
'Close #1

If App.PrevInstance = True Then
    End
End If

If IsIDE = False Then
    'Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + App.Path + "\" + App.EXEName + ".vbp""", vbHide
    'End
End If

If IsIDE = False Then
Form1.Caption = "WebCam Motion Capture"
Else
Form1.Caption = "WebCam Motion Capture In IDE"
End If

ChDir App.Path



'set up the visual stuff
Picture1.Width = 640 * Screen.TwipsPerPixelX
Picture1.Height = 480 * Screen.TwipsPerPixelY
Picture1.Left = 0
Picture1.Top = 0

Form1.Width = Picture1.Width + 110
Form1.Height = Picture1.Height + 420


'Inten is the measure of how many pixels are going to be recognized. I highly dont recommend
'setting it lower than this, i have a 3.0 GHz PC and it starts to lag a little. On this setting,
'every 15th pixel is checked
inten = 20
'The tolerance of recognizing the pixel change
Tolerance = 20

Tppx = Screen.TwipsPerPixelX
Tppy = Screen.TwipsPerPixelY

ReDim POn(640 / inten, 480 / inten)
ReDim P(640 / inten, 480 / inten)

'Exit Sub

If IsIDE = True Then Form1.Show: Me.WindowState = 0

'Form1.Show
'TJax = Now + TimeSerial(0, 0, 10)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'STOPCAM
'End

Exit Sub

Dim DD, ii, iG
On Error Resume Next
Open App.Path + "\VBDataNoArchive\WebCam % Pic Data2.txt" For Input As #1
Line Input #1, DD
Close #1
DD = Mid$(DD, 4)
ii = DateValue(DD) + TimeValue(DD)
iG = DateDiff("n", ii, Now)
If iG > 10 Then Form1.Show: msgbox "No WebCam Pic Has Been Taken for 10 Mins"
On Error GoTo 0


End Sub



Private Sub Timer1_Timer()


On Error Resume Next
Set Picture1 = LoadPicture(App.Path + "\VBDataNoArchive\ArrayPic.jpg")
On Error GoTo 0

Call ComParePic

STARTCAM

'Get the picture from camera.. the main part
SendMessage mCapHwnd, GET_FRAME, 0, 0


On Error Resume Next
TYX = Clipboard.GetFormat(vbCFText)
If TYX = True Then TBak = Clipboard.GetText
Clipboard.Clear

SendMessage mCapHwnd, COPY, 0, 0
'Exit Sub

GJax = Now + TimeSerial(0, 0, 5)
Do
    DoEvents
    TY = Clipboard.GetFormat(vbCFBitmap)
    If Clipboard.GetFormat(vbCFBitmap) = True Then Exit Do
Loop Until TY = True Or GJax < Now
If TY = False Then Form1.Show: DoEvents: STOPCAM: msgbox "Web_Cam Not On"

If TY = True Then
    Picture1.Picture = Clipboard.GetData
    STOPCAM
End If

If TY = False Or Err.Number > 0 Then
If TYX = True Then
    Clipboard.Clear
    Clipboard.SetText (TBak)
End If
Unload Form1: Exit Sub
End If
    
Dim FF$

If TBak <> "" Then
Clipboard.Clear
Clipboard.SetText (TBak)
End If

Call ComParePic

If WhiteFactor > 290 Then
'If WhiteFactor > 190 Then
    If IsIDE = False Or Form1.Visible = False Then Unload Form1
End If





'Set Ip = Picture1

With Picture2
.Height = Picture1.Height
.Left = Picture1.Left
.Top = Picture1.Top
.Width = Picture1.Width
.Visible = False
End With

'state all statistics
With Picture1
Label1.Caption = Int(Wo / (Ri + Wo) * 100) & " % movement" & vbCrLf & "Real Movement: " & RealRi & vbCrLf _
& "Completed in: " & GetTickCount - LastTime
.ForeColor = vbGreen
.CurrentX = 545 * Screen.TwipsPerPixelX
.CurrentY = 445 * Screen.TwipsPerPixelY
Tx = Wo / (Ri + Wo) * 100
If Tx = 100 Then Tx = 99
If Tx < 10 Then HH = " "
Picture1.Print "Motion " + HH;
.ForeColor = vbWhite
Picture1.Print Format$(Wo / (Ri + Wo) * 100, "0") & "%"
.CurrentX = 445 * Screen.TwipsPerPixelX
.CurrentY = 459 * Screen.TwipsPerPixelY
.ForeColor = vbGreen
Picture1.Print Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
End With
DoEvents

'BlendPicture Picture1, Picture2



FF$ = "WebCapture_" + Format$(Now, "YYYY-MM-DD-DDD")
On Error Resume Next
MkDir "M\0 00 Art Loggers\Web_Cam\" + FF$
If Err.Number = 76 Then Unload Form1: Exit Sub
On Error GoTo 0
FileInUseDelay App.Path + "\VBDataNoArchive\Web_Cam Pic Always.jpg"
SavePicture Picture1, App.Path + "\VBDataNoArchive\Web_Cam Pic Always.jpg"
FileInUseDelay App.Path + "\VBDataNoArchive\WebCam % Pic Data.txt"
Open App.Path + "\VBDataNoArchive\WebCam % Pic Data.txt" For Append As #1
Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + " - " + Format$(Wo / (Ri + Wo) * 100, "0") & "%"
Close #1
FileInUseDelay App.Path + "\VBDataNoArchive\WebCam % Pic Data2.txt"
Open App.Path + "\VBDataNoArchive\WebCam % Pic Data2.txt" For Output As #1
Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
Close #1
If Wo / (Ri + Wo) * 100 >= 8 Then
    FileInUseDelay App.Path + "\VBDataNoArchive\ArrayPic.jpg"
    SavePicture Picture1, App.Path + "\VBDataNoArchive\ArrayPic.jpg"
    TS = "M\0 00 Art Loggers\Web_Cam\" + FF$ + "\WebCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD ") + "- " + Format$(Wo / (Ri + Wo) * 100, "0") & "% Compare -- " + Format$(WhiteFactor, "0.00") & "%-WhiteFactor.jpg"
    FileInUseDelay TS
    SavePicture Picture1, TS
End If

Timer1.Enabled = False

'Exit Sub

If IsIDE = False Or Form1.Visible = False Then Unload Form1

'Exit Sub

Exit Sub
Errcode2:
DoEvents
Resume Next

End Sub

Sub ComParePic()
WhiteFactor2 = 0
Ri = 0 'right
Wo = 0 'wrong

'LastTime = GetTickCount
Cat = 0
For i = 0 To 640 / inten - 1
    For j = 0 To 480 / inten - 1
    'get a point
    C = Picture1.POINT(i * inten * Tppx, j * inten * Tppy)
    'analyze it, Red, Green, Blue
        R = C Mod 256
        G = (C \ 256) Mod 256
        B = (C \ 256 \ 256) Mod 256
        
        WhiteFactor = R + G + B
        WhiteFactor2 = WhiteFactor2 + WhiteFactor
        If Cat = 1 Then WhiteFactor2 = WhiteFactor2 / 2
        Cat = 1
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
    'make a red dor
    P(i, j) = Picture1.POINT(i * inten * Tppx, j * inten * Tppy)
    Picture1.PSet (i * inten * Tppx, j * inten * Tppy), vbRed
    POn(i, j) = False
    End If
    
    Next j
    
Next i

RealRi = 0

For i = 1 To 640 / inten - 2
    For j = 1 To 480 / inten - 2
    If POn(i, j) = False Then
        'Real movement is simply occuring when all 4 pixels around one pixel changed
        'Simply put, If this pixel is changed and all around it changed too, then this is a real
        'movement
        If POn(i, j + 1) = False Then
            If POn(i, j - 1) = False Then
                If POn(i + 1, j) = False Then
                    If POn(i - 1, j) = False Then
                    RealRi = RealRi + 1
                    Picture1.PSet (i * inten * Tppx, j * inten * Tppy), vbGreen
                    End If
                End If
            End If
        End If
        
    End If
        
        
    Next j
Next i

WhiteFactor = WhiteFactor2

'BlendPicture Picture2, Picture1

End Sub

Sub STOPCAM()
On Error Resume Next
DoEvents
SendMessage mCapHwnd, DISCONNECT, 0, 0
'Timer1.Enabled = False
End Sub

Sub STARTCAM()

'// Enumerate installed drivers
Dim lpszName As String * 100
Dim lpszVer As String * 100
Dim x As Byte
Dim lResult As Boolean
Dim Caps As CAPDRIVERCAPS

For x = 0 To 19
    lResult = capGetDriverDescriptionA(x, lpszName, 100, lpszVer, 100) '// Retrieves driver info
    'MsgBox lpszName
    If lResult Then
        If InStr(lpszName, "Microsoft WDM Image Capture (Win32)") > 0 Then
            Exit For
        End If
    End If
Next x

If lResult = False And IsIDE = True Then msgbox "WebCam Driver Not Running": End


'// Connect the capture window to the driver
'lResult = capDriverConnect(lwndC, lpszName)
'If lResult = False Then End: MsgBox "WebCam Driver Not Running2"

''// Get the capabilities of the current capture driver
'lResult = capDriverGetCaps(lwndC, VarPtr(Caps), Len(Caps))
'Caps.wDeviceIndex
'If lResult = False Then End: MsgBox "WebCam Driver Not Running2"


mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, Me.hwnd, 0)
If mCapHwnd = 0 Then Form1.Show: DoEvents: msgbox "WebCam Driver Not Running"

'GetStatus


SendMessage mCapHwnd, CONNECT, 0, 0
'GetStatus
'Timer1.Enabled = True
End Sub

Sub GetStatus()
    Dim CAPSTATUS As CAPSTATUS
    Dim lCaptionHeight As Long
    Dim lX_Border As Long
    Dim lY_Border As Long
    Dim A, DumVar
    
    lCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
    lX_Border = GetSystemMetrics(SM_CXFRAME)
    lY_Border = GetSystemMetrics(SM_CYFRAME)
    
    '// Get the capture window attributes .. width and height
    DumVar = capGetStatus(lwndC, VarPtr(CAPSTATUS), Len(CAPSTATUS))
    A = CAPSTATUS.uiImageWidth
    
End Sub

'Originally for saving the picture output... i left it out since it has nothing to do with program
'i commented just so you can see how its done in case you dont know
'Private Sub Timer2_Timer()
'SavePicture Picture1.Image, "C:\pics\img" & Counter & ".bmp"
'Counter = Counter + 1
'End Sub

Sub Select2()
Dim lpszName As String * 100
Dim lpszVer As String * 100
Dim x As Byte
Dim lResult As Boolean
Dim Caps As CAPDRIVERCAPS
'// Enumerate installed drivers
For x = 0 To 9
lResult = capGetDriverDescriptionA(x, lpszName, 100, lpszVer, 100) '// Retrieves driver info
If lResult Then CmboSource.AddItem lpszName
Next x
'// Get the capabilities of the current capture driver
lResult = capDriverGetCaps(lwndC, VarPtr(Caps), Len(Caps))
'// Select the driver that is currently being used
If CmboSource.ListCount = 0 Then
msgbox "Error: No devices found!"
End
Exit Sub

End If
If CmboSource.ListCount = 1 Then
CmboSource.ListIndex = 0
Call cmdSelect_Click
End If
If lResult Then CmboSource.ListIndex = Caps.wDeviceIndex
End Sub

Private Sub cmdSelect_Click()
   
    Dim sTitle As String
    Dim Caps As CAPDRIVERCAPS
    
    If CmboSource.ListIndex <> -1 Then
        
        '// Connect the capture window to the driver
        If capDriverConnect(lwndC, CmboSource.ListIndex) Then
    
            '// Get the capabilities of the capture driver
            capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
            
            '// If the capture driver does not support a dialog, grey it out
            '// in the menu bar.
            'frmMain.mnuSource.Enabled = Caps.fHasDlgVideoSource
            'frmMain.mnuFormat.Enabled = Caps.fHasDlgVideoFormat
            'frmMain.mnuDisplay.Enabled = Caps.fHasDlgVideoDisplay
        
            sTitle = CmboSource.Text
            
            SetWindowText lwndC, sTitle
            ResizeCaptureWindow lwndC
        End If
    
    End If
    
    
    'Unload Me
   
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
        msgbox "Trouble File Stuck In Use " + vbCrLf + Txw$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub




