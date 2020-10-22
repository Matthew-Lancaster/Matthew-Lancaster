VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00560107&
   BorderStyle     =   0  'None
   Caption         =   "CLOCK"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.PictureBox Picture3 
      Height          =   3375
      Left            =   5250
      Picture         =   "CLOCK #M02 SCR.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.Timer PATTERNTimer 
      Interval        =   5
      Left            =   2730
      Top             =   2280
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   5
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   1170
      HelpContextID   =   100
      Left            =   3450
      ScaleHeight     =   1170
      ScaleMode       =   0  'User
      ScaleWidth      =   1245
      TabIndex        =   6
      Top             =   2940
      Width           =   1245
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   5  'Not Copy Pen
         Height          =   1065
         Left            =   -60
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   585
         Left            =   351
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      Height          =   1215
      Left            =   3450
      ScaleHeight     =   1155
      ScaleWidth      =   1230
      TabIndex        =   1
      Top             =   1425
      Visible         =   0   'False
      Width           =   1290
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "W      E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   375
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   " S "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   450
         TabIndex        =   3
         Top             =   750
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   " N "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   450
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Timer TimerSCROLL 
      Interval        =   100
      Left            =   2775
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2250
      Top             =   1800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   " VB ME "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4575
      TabIndex        =   5
      Top             =   4275
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1950
      TabIndex        =   0
      Top             =   600
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ri, RR3
Dim RR, KO, KOX, OTF, RRCOZ, SWAPFLAG
Dim ONOWDATE As Date, POINTCOUNT, DD, DF
Dim OPOINT, RUN, COUNTH
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const E = 2.7182818284
Const pi = 3.141592648
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const shrdNoMRUString = &H2


Private Type POINTAPI
        x As Long
        y As Long
End Type

'More information: Creating a screen saver in Visual Basic is very simple. You create one form and set the following properties:
'Caption ""
'ControlBox False
'MinButton False
'MaxButton False
'BorderStyle 0 - None
'WindowState 2 - Maximized
'BackColor &H0

'Option Explicit
Dim TT, R, TV, FS
Const T As Double = 57.29577951


Private Sub Form_Click()
If IsIDE Then End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If IsIDE Then End
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End

If IsIDE = True Then
    Set FS = CreateObject("Scripting.FileSystemObject")
    FS.COPYFILE App.Path + "\" + App.EXEName + ".EXE", App.Path + "\" + App.EXEName + ".SCR"
End If

RUN = 20

'HE SAID MIRROR NOT PROPER 3:06A 22 NOV 2010

Me.Show
'DoEvents
Call SUBGO



End Sub


Sub SUBGO()

DoEvents

If KO = True Then Exit Sub
KOX = True

    
On Error Resume Next

'Me.WindowState = 1
Me.WindowState = 2



Me.ScaleHeight = 1000
Me.ScaleWidth = 1000

Me.Width = 1000
Me.Height = 1000

Me.BackColor = &H560107

Picture3.BackColor = &H80FFFF

i = AlwaysOnTop(Me.hwnd)

Picture2.ScaleHeight = 100
Picture2.ScaleWidth = 100

Picture2.Height = 180 '450 'Picture2.ScaleHeight
Picture2.Width = Picture2.Height / 1.4



Picture2.DrawWidth = 10


'Me.CurrentX = 10
'Me.CurrentY = 10

DoEvents

XWidth = Me.ScaleWidth
XHEIGHT = Me.ScaleHeight

'XWidth = (Me.Width / Me.ScaleWidth) * Screen.TwipsPerPixelX
'XHEIGHT = (Me.Height / Me.ScaleHeight) * Screen.TwipsPerPixelY

'XWidth = Me.Width
'XHEIGHT = Me.Height


Label1.Caption = UCase(Format(Now, "DDDD"))
DoEvents
Label1.Top = 10
Label1.Left = XWidth - Label1.Width

Picture3.Top = Label1.Top + Label1.Height
Picture3.Left = XWidth - Picture3.Width

Picture2.Top = Picture3.Top + Picture3.Height + 2
Picture2.Left = XWidth - Picture2.Width

'VB ME
Label5.Top = XHEIGHT - Label5.Height
Label5.Left = XWidth - Label5.Width


If IsIDE Then
    STATE_SHOW
    Me.Show
    Call Form_Resize
    DoEvents
End If

Me.SetFocus

If Err.Number = 0 Then KO = True: KOX = False

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


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then End
End Sub

Private Sub Form_Resize()
On Error Resume Next

If KOX = True Then Exit Sub

Me.Top = 0
Me.Left = 0
Me.ScaleHeight = 1000
Me.ScaleWidth = 1000

Call SUBGO




'If Me.WindowState = 1 Then STATE_HIDE: Exit Sub
'
'If Me.WindowState = 2 Then STATE_SHOW

'NOT PROPER 3.20A
'NAME
'RESPOND COME ON TEACH ME MORE
'NP>



'Me.ScaleHeight = 1000
'Me.ScaleWidth = 1000


End Sub

Private Sub Form_Unload(Cancel As Integer)

If Me.WindowState <> 1 Then Cancel = True: Exit Sub


End
End Sub

Private Sub Label5_Click()
If IsIDE = True Then End
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus

End
End Sub


Private Sub picture3_Click()
'End
End Sub

Private Sub Timer1_Timer()
Dim H As Long, M As Long, S As Long                     'time units
Dim Hd As Double, Md As Double, Sd As Double            'Degrees
Dim Hr As Double, Mr As Double, Sr As Double            'Radians

Me.Cls

H = Hour(Time): M = Minute(Time): S = Second(Time)

If H >= 12 Then H = H - 12

Hd = H * 30
Hd = Hd + M / 2
Md = M * 6
Sd = S * 6

Hd = Hd - 90: Md = Md - 90: Sd = Sd - 90

If Hd < 0 Then Hd = Hd + 360
If Md < 0 Then Md = Md + 360
If Sd < 0 Then Sd = Sd + 360

Hr = Hd / T: Mr = Md / T: Sr = Sd / T

Call CLOCK_NUMBER_HOUR
Call CLOCK_NUMBER_MINUTE


TT = Me.ScaleHeight / 2
Me.DrawWidth = 13
Me.Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.5 * Cos(Hr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.5 * Sin(Hr))), vbWhite
Me.Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.6 * Cos(Mr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.6 * Sin(Mr))), vbBlue
Me.DrawWidth = 13
Me.Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.7 * Cos(Sr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.7 * Sin(Sr))), vbGreen

Me.DrawWidth = 8
Me.Circle (Me.ScaleWidth / 2, Me.ScaleHeight / 2), 5, vbWhite
End Sub


Sub CLOCK_NUMBER_HOUR()

Dim H As Long, M As Long, S As Long                     'time units
Dim Hd As Double, Md As Double, Sd As Double            'Degrees
Dim Hr As Double, Mr As Double, Sr As Double            'Radians

For R = 0 To 12

'Me.Cls
H = R: M = Minute(0): S = Second(0)

If H >= 12 Then H = H - 12

Hd = H * 30
Hd = Hd + M / 2
Md = M * 6
Sd = S * 6

Hd = Hd - 90: Md = Md - 90: Sd = Sd - 90

If Hd < 0 Then Hd = Hd + 360
If Md < 0 Then Md = Md + 360
If Sd < 0 Then Sd = Sd + 360

Hr = Hd / T: Mr = Md / T: Sr = Sd / T




TT = Me.ScaleHeight / 2
TV = 370 'Int((ME.ScaleHeight / 2) / 1.39)
Me.FillColor = vbWhite
Me.FillStyle = vbSolid
Me.DrawWidth = 1
'Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.5 * Cos(Hr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.5 * Sin(Hr))), vbGreen
Me.Circle (TT + (Cos(Hr) * TV), TT + (Sin(Hr) * TV)), 13, vbWhite


Next

End Sub

Sub CLOCK_NUMBER_MINUTE()
'Exit Sub
Dim H As Long, M As Long, S As Long                     'time units
Dim Hd As Double, Md As Double, Sd As Double            'Degrees
Dim Hr As Double, Mr As Double, Sr As Double            'Radians

For R = 0 To 59

'Me.Cls
H = 1: M = R: S = Second(0)

If H >= 12 Then H = H - 12

Hd = H * 30
Hd = Hd + M / 2
Md = M * 6
Sd = S * 6

Hd = Hd - 90: Md = Md - 90: Sd = Sd - 90

If Hd < 0 Then Hd = Hd + 360
If Md < 0 Then Md = Md + 360
If Sd < 0 Then Sd = Sd + 360

Hr = Hd / T: Mr = Md / T: Sr = Sd / T


TT = Me.ScaleHeight / 2
TV = 370

Me.FillColor = vbWhite
Me.FillStyle = vbSolid
Me.DrawWidth = 1
'Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.5 * Cos(Hr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.5 * Sin(Hr))), vbGreen
Me.Circle (TT + (Cos(Mr) * TV), TT + (Sin(Mr) * TV)), 3, vbWhite






Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then End
    
    'If the mouse is moved, unload the Screen Saver
    'Debug.Print X
    'Give enough time for program to run
    COUNTH = COUNTH + 1
    If DF < Now Then COUNTH = 0:    DF = Now + TimeSerial(0, 0, 4)

    If COUNTH > 80 Then
        COUNTH = 0
        Call STATE_HIDE
        'Me.Top = 0
        'Me.Left = 0
        'Me.Width = 1000
        'Me.Height = 1000
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If a key is pressed, unload the Screen Saver
    'If IsIDE Then End
    'Call STATE_HIDE
End
End Sub

Private Sub TimerSCROLL_Timer()
    
    'WORLDS APART CAME --- NP> 3:25A
    If Me.WindowState = 2 Then Exit Sub
    Dim P As POINTAPI

    GetCursorPos P
    If OPOINT <> P.x + P.y Then
        POINTCOUNT = Now + TimeSerial(0, 0, RUN)
        OPOINT = P.x + P.y
        'BOBBING UP AND DOWN
    End If
    
    i = 0
    For i = 1 To 255
        bdf = GetAsyncKeyState(i)
        If bdf < -300 Then
            POINTCOUNT = Now + TimeSerial(0, 0, RUN)
        End If
    Next

    'OLD NP> 3:25A
    'REF SARAH

    'DAD NP> 3:25A

    'Debug.Print POINTCOUNT
    If POINTCOUNT > Now Or POINTCOUNT = 0 Then Exit Sub
    
    POINTCOUNT = 0
    RUN = 20 * 60
    Call STATE_SHOW

End Sub


Sub STATE_SHOW()
    Me.WindowState = 2
    'TimerSCROLL.Enabled = False
    'Timer1.Enabled = True
    DD = Now + TimeSerial(0, 0, 5)
End Sub

Sub STATE_HIDE()
    If DD > Now Then Exit Sub
    
    
    
    POINTCOUNT = Now + TimeSerial(0, 0, RUN)
    Me.WindowState = 1
    'Me.ShowInTaskbar
    
    
    'Timer1.Enabled = False
    'TimerSCROLL.Enabled = True
End Sub

Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Function


Private Sub PATTERNTimer_Timer()
    
On Error GoTo ENDSUB
    
    MODTIMER = 20
    OTFX = Int(Timer) Mod MODTIMER
    
    If OTFX < MODTIMER / 2 And OTFX <> OTF Then
        OTF = Int(Timer) Mod 20
        SWAPFLAG = SWAPFLAG + 1
        If SWAPFLAG > 3 Then SWAPFLAG = 1
    End If

    'SWAPFLAG = 1
    'SWAPFLAG = 2
    'SWAPFLAG = 3

    RRCOZ = RRCOZ + 0.001
    RRVAR = 0.1
    
    If SWAPFLAG <> 2 Then RR = RR + RRVAR
    If SWAPFLAG = 2 Then RR = RR + RRVAR + (Cos(RRCOZ) / 3.5)
    
    OFFSETX = 8
    OFFSETY = 0
    
    TX = (Picture2.Width / 2) + OFFSETX
    TY = (Picture2.Height / 2)
    
    OM = 78 'Int(Picture2.Width / Screen.TwipsPerPixelX)
    OMX = 30
    If SWAPFLAG = 3 Or 8 = 8 Then
        RR3 = RR3 + 0.02
        OMX = (OM / 2.2) * Sin(RR3) + (OM / 2)
    End If
    
    'HALF EXAMPLE
    '------------------------
    '40 OUTER
    '20 SIZE
    'OUTER - SIZE / 2 = 30
    '=30 CIRCUMFERANCE
    
    'SIZE 20 IN 30 CIRCUMFERANCE
    'LOWER CIRCUMFERANCE BY HALF SIZE
    'TEST 30 - 10 = 20      PASS 2 = 20 +20
    
    'LOWER HALF EXAMPLE
    '------------------------
    '40 OUTER
    '10 SIZE
    'OUTER - SIZE / 2 = 15
    
    '1
    '--- MODIFY = CIRCUMFERANCE RAIL TO BE DOUBLE
    
    'RESULT ARRIVED AT
            
    'HIGHER
    'TOP
    'UPPER
            
    'TOP HALF EXAMPLE
    '------------------------
            
    
    
    
    TXD = OMX
    
    REMIANING_DIVEDE_DIAMETER = (OM - OMX) / 2
    Ri = REMIANING_DIVEDE_DIAMETER
    
    If OMX < OM Then
        Ri = REMIANING_DIVEDE_DIAMETER * 2
    Else
        Ri = REMIANING_DIVEDE_DIAMETER / 2
    End If
    REMIANING_DIVEDE_DIAMETER = Ri
    
    
    x = (Ri * Sin(RR) + TX)
    y = (Ri * Cos(RR) + TY)
    
    If SWAPFLAG = 3 Then
        y = (Ri * Cos(RR * 1.2) + TY)
    End If
    'Me.Picture2.Width =
    
    Me.Picture2.Cls
    Me.Picture2.DrawWidth = 3
    Me.Picture2.Circle (TX + OFFSETX, TY + OFFSETY), OM
    Me.Picture2.Circle (x + OFFSETX, y + OFFSETY), OMX

Exit Sub
ENDSUB:
    
    Me.Hide
    Stop
    Resume
    
End Su