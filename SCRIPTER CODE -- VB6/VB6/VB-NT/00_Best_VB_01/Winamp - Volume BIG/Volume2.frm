VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Volume2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Volume2"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Volume2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   -180
      TabIndex        =   0
      Top             =   -30
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1665
      Top             =   1530
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1665
      Top             =   945
   End
End
Attribute VB_Name = "Volume2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MeTop, Show2, HH
Public GG2
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Public ForeG

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect _
   As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim Rect2 As RECT


    
    
Public Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function

Public Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function


Private Sub Form_Load()
If App.PrevInstance = True Then End
ProgressBar1.Min = 0
ProgressBar1.Max = 100
'Me.Show
Me.Width = 0
Me.Height = 200

ProgressBar1.ToolTipText = App.Path + " -- " + App.EXEName

'Call PosiXY

'AlwaysOnTop (Volume2.hWnd)
'NotAlwaysOnTop (ATidalX.hWnd)
Call OpenMixer(0)

ProgressBar1.Value = GetVolume(SPEAKER)
Me.Visible = False

End Sub

Sub PosiXY()

'qq = GetForegroundWindow
XT = FindWindow("Winamp v1.x", vbNullString)
If XT = 0 Then End

If XT = 0 And Show2 = True Then Me.Hide: Show2 = False
If XT = 0 Then Exit Sub

re = GetWindowRect(XT, Rect2)
If XT + Rect2.Left + Rect2.Top = ForeG Then Exit Sub
ForeG = XT + Rect2.Left + Rect2.Top

XT = FindWindow("Winamp v1.x", vbNullString)

If XT = 0 Then End

If XT > 0 Then
'    Exit Sub
    'Me.Top = Screen.Height - (Me.Height) - 10
    'Me.Left = 10
    re = GetWindowRect(XT, Rect2)
    Me.Top = (Rect2.Top * Screen.TwipsPerPixelY) - Me.Height
    Me.Left = (Rect2.Left * Screen.TwipsPerPixelX)
    
    Me.Width = (Rect2.Right - Rect2.Left) * Screen.TwipsPerPixelX
    DoEvents
    
    If Show2 = False Then Me.Show: Show2 = True
End If


End Sub

Private Sub Form_Resize()

ProgressBar1.Top = 0
ProgressBar1.Left = 0
ProgressBar1.Height = Me.Height
ProgressBar1.Width = Me.Width

End Sub

Private Sub ProgressBar1_Click()
If GG2 = 0 Then GG2 = Now + TimeSerial(0, 0, 1): HH = 0
If GG2 < Now Then GG2 = 0: HH = 0: Exit Sub
HH = HH + 1
If HH = 2 Then End
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'End

If Button = 2 Then

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End
End If


End Sub

Private Sub Timer1_Timer()


ProgressBar1.Value = GetVolume(SPEAKER)


End Sub

Private Sub Timer2_Timer()
Call PosiXY
End Sub
