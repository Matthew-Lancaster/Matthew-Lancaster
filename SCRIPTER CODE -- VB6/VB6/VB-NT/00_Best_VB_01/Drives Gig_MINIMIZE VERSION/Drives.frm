VERSION 5.00
Begin VB.Form Drives 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Free Gigabyte Memory"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6912
   LinkTopic       =   "Form1"
   ScaleHeight     =   3192
   ScaleWidth      =   6912
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   4050
      Top             =   1710
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3600
      Top             =   1725
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3150
      Top             =   1725
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1095
      TabIndex        =   27
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   1995
      Width           =   1560
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Bin Lad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   26
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   1995
      Width           =   1080
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2685
      TabIndex        =   25
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1095
      TabIndex        =   24
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   285
      Width           =   1560
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   23
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   285
      Width           =   1080
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2670
      TabIndex        =   22
      Top             =   285
      Width           =   345
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2670
      TabIndex        =   21
      Top             =   1710
      Width           =   345
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Mem Use"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   20
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1095
      TabIndex        =   19
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   1710
      Width           =   1560
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1095
      TabIndex        =   18
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   1425
      Width           =   1560
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Tott Tott"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   17
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   1425
      Width           =   1080
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2670
      TabIndex        =   16
      Top             =   1425
      Width           =   345
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1095
      TabIndex        =   15
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   1140
      Width           =   1560
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Tott Free"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   14
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2670
      TabIndex        =   13
      Top             =   1140
      Width           =   345
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2670
      TabIndex        =   12
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2670
      TabIndex        =   11
      Top             =   570
      Width           =   345
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2670
      TabIndex        =   10
      Top             =   855
      Width           =   345
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "S:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   855
      Width           =   1080
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1095
      TabIndex        =   8
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   855
      Width           =   1560
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Mem"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   2685
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   615
      TabIndex        =   6
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   2685
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Back2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2775
      TabIndex        =   5
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   2970
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Back1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2775
      TabIndex        =   4
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   2610
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1095
      TabIndex        =   3
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   570
      Width           =   1560
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "M:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1095
      TabIndex        =   1
      Top             =   0
      Width           =   1560
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   0
      Width           =   1080
   End
End
Attribute VB_Name = "Drives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LockMove, OHwnd1, OHwnd2, Nuke7, NORESIZE
Dim Nuke4, Nuke5, RR

Dim OLLAB5
Dim OLLAB8
Dim OLLAB22
Dim OLLAB10

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

Public MHD1, MHD, MHD3, MHD4
Public TotHd, TuLip
Public Lk1, Lk2, Lk3, Lk4, FreeMem, IIO2, IIO3

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Sub Form_Load()

'END


If App.PrevInstance = True Then End

Set FS = CreateObject("Scripting.FileSystemObject")



'frmMemInfo.Show
OHwnd = 20

Me.BackColor = 0
IIO2 = -1
IIO3 = -1
Lk4 = -1
wc = 0: hc = 0
On Local Error Resume Next
For Each Control In Drives
    If Control.Enabled = True Then
        If Control.Left + Control.Width > wc Then wc = Control.Left + Control.Width
        If Control.Top + Control.Height > hc Then hc = Control.Top + Control.Height
    End If
Next
On Local Error GoTo 0

Drives.Width = wc + 90 + hq
Drives.Height = hc + 440


Load frmMemInfo

'Drives.Show

Me.Show
Me.WindowState = 1




Call Timer1_Timer
Call Timer2_Timer
Call Timer3_Timer






End Sub

Private Sub FORM_RESIZE()

If NORESIZE = True Then Exit Sub
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
If Err.Number > 0 Then Exit Sub
NORESIZE = True



Exit Sub

'    Call Resize_Code
wc = 0: hc = 0
On Local Error Resume Next
For Each Control In Drives
    If Control.Enabled = True Then
        If Control.Left + Control.Width > wc Then wc = Control.Left + Control.Width
        If Control.Top + Control.Height > hc Then hc = Control.Top + Control.Height
    End If
Next
'On Local Error GoTo 0

Drives.Width = wc + 90 + hq
Drives.Height = hc + 380

HWNDW = FindWindow(vbNullString, "Active Idle...")
OHwnd = GetWindowState(HWNDW)

If HWNDW = 0 Then
    If Me.WindowState = 0 Then
        'DONT LOAD IF IDEL ACTIVE RUNNING
        'TEMP REMOVE
        'End
        'Me.WindowState = 1
        'Me.Visible = False
    End If
Exit Sub
End If

If HWNDW <> 0 And OHwnd = 1 Then
    If Me.WindowState = 0 Then
        
        'TEMP REMOVE
        'Me.WindowState = 1
        'Me.Visible = False
    End If

End If




'POSTIONING
Call Timer1_Timer
End Sub

Public Sub Resize_Code()
'Call Timer1_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label10_Click()
TuLip = True

If OLLAB10 > 0 And Val(Label10) = 0 Then Me.WindowState = 0
OLLAB10 = Label10


End Sub

Private Sub Label11_Click()
TuLip = True
End Sub

Private Sub Label15_Click()
TuLip = True
End Sub

Private Sub Label16_Click()
TuLip = True
End Sub

Private Sub Label17_Click()
End
End Sub

Private Sub Label18_Click()
TuLip = True
End Sub

Private Sub Label19_Click()
TuLip = True
End Sub

Private Sub Label20_Click()
TuLip = True
End Sub

Private Sub Label21_Click()
TuLip = True
End Sub

Private Sub Label22_Click()
TuLip = True

If OLLAB22 > 0 And Val(Label22) = 0 Then Me.WindowState = 0
OLLAB22 = Label22



End Sub

Private Sub Label26_Click()

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End

End Sub

Private Sub Label5_Click()
TuLip = True


If OLLAB5 > 0 And Val(Label5) = 0 Then Me.WindowState = 0
OLLAB5 = Label5



End Sub

Private Sub Label6_Click()
TuLip = True
End Sub

Private Sub Label7_Click()
'Call ATidalX.Mnu_Idle_Now_Click
TuLip = True
End Sub

Private Sub Menu_EndDay_Click()
'Shell "notepad " + App.Path + "\00_Text_Data_02\Idle and Active Logger End Day.txt", vbNormalFocus

End Sub

Private Sub Mnu_IdleOver_Click()
'Shell "notepad " + App.Path + "\00_Text_Data_02\Idle and Active Logger Idle Over Time.txt", vbNormalFocus

End Sub

Private Sub Mnu_Now2_Click()
'Shell "notepad " + App.Path + "\00_Text_Data_02\Idle and Active Logger Now2.txt", vbNormalFocus

End Sub

Private Sub Label8_Click()
'Drives.Label8
TuLip = True

If OLLAB8 > 0 And Val(Label8) = 0 Then Me.WindowState = 0
OLLAB8 = Label8


End Sub

Private Sub Label9_Click()
'ATidalX.Mnu_Idel_Click
TuLip = True
End Sub


Private Sub Timer1_Timer()



Timer1.Enabled = False

Exit Sub

Dim Rect4 As RECT
'Exit Sub

HWNDW = FindWindow(vbNullString, "Active Idle...")
OHwnd = GetWindowState(HWNDW)

If HWNDW = 0 Then
    If Me.WindowState = 0 Then
        'TEMP REMOVE
        
        'End
        'Me.WindowState = 1
        ''Me.ShowInTaskbar = True
        'Me.Visible = False
    End If
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

Exit Sub
End If
'If HWNDW <> 0 And OHwnd1 = 0 Then
'If Me.WindowState = 1 Then Me.WindowState = 0
'End If

If HWNDW <> 0 And OHwnd = 1 Then
    If Me.WindowState = 0 Then
        Me.Top = Screen.Height / 2 - Me.Height / 2
        Me.Left = Screen.Width / 2 - Me.Width / 2

        Me.WindowState = 1:        Me.Visible = False: Exit Sub
    End If

End If

If HWNDW <> 0 And OHwnd = -1 Then
    If Me.WindowState = 1 Then
        'TEMP REMOVE
        Me.WindowState = 0:        Me.Visible = True
    End If

End If


OHwnd1 = OHwnd
OHwnd2 = HWNDW

HWNDW = FindWindow(vbNullString, "Active Idle...")

If GetWindowState(HWNDW) <> -1 Then
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2

Exit Sub
End If

If (GetAsyncKeyState(1) And GetForegroundWindow = Me.hwnd) Then
    LockMove = Now + Timer + 0.1
     Exit Sub
End If

   
If LockMove > Now + Timer Then
     Exit Sub
End If



If HWNDW > 0 Then
Hwnd24 = GetWindowRect(HWNDW, Rect4)
On Local Error Resume Next
If Lk1 <> Rect4.Top Then
    Drives.Top = (Rect4.Top * Screen.TwipsPerPixelX)
    Drives.Left = (Rect4.Left * Screen.TwipsPerPixelY - Drives.Width)
Else
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

End If
Lk1 = Rect4.Top + Rect4.Left
pd = 1
End If


HWNDW = FindWindow(vbNullString, "Tidal Information...")

If HWNDW > 0 And pd = 0 Then

Hwnd24 = GetWindowRect(HWNDW, Rect4)
On Local Error Resume Next
If Lk2 <> Rect4.Top And Rect4.Top > 0 Then
    Drives.Top = (Rect4.Top * Screen.TwipsPerPixelX)
    Drives.Left = (Rect4.Left * Screen.TwipsPerPixelY - Drives.Width)
    Lk2 = Rect4.Top + Rect4.Left
Else
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2


End If
'Exit Sub
pd = 2
End If

HWNDW = FindWindow(vbNullString, "BBC 24 Hour Weather")
If HWNDW > 0 And GetWindowState(HWNDW) <> vbMinimized Then
Hwnd24 = GetWindowRect(HWNDW, Rect4)
If Lk3 <> Rect4.Top And Rect4.Top > 0 Then
    'Drives.Top = (Rect4.Bottom * Screen.TwipsPerPixelX)
    'drives.Left = (ATidalX.Left - AI.Width)
    Lk3 = Rect4.Top + Rect4.Left
End If
pd2 = 1
End If


If Lk1 + Lk2 + Lk3 <> Lk4 Then
    'If pd2 = 0 Then Drives.Top = (420)
    'If pd = 0 Then Drives.Left = (Screen.Width - Drives.Width)
    Lk4 = Lk1 + Lk2 + Lk3
End If

End Sub

Sub CDRIVE()
Dim FS
Set FS = CreateObject("Scripting.FileSystemObject")

DRI = "C:\"
Set G = FS.GetDrive(DRI)
If G.ISREADY = False Then Exit Sub


nuke1 = G.freespace
'If nuke1 < MHD1 And nuke1 <> MHD1 Then E$ = "ã" Else E$ = "ä"
If nuke1 < MHD1 Then E$ = "ä"
If nuke1 > MHD1 And MHD1 > 0 Then E$ = "ã"
If nuke1 = MHD1 Then E$ = Label13.Caption
If MHD1 = 0 Then E$ = Label13.Caption
Label10.Caption = Format$(nuke1 / 1024 ^ 3, "0.000000")
Label13.Caption = E$

If MHD1 <> nuke1 And MHD1 > 0 Then
    Label11.BackColor = Label2.BackColor
Else
    Label11.BackColor = Label1.BackColor
End If
MHD1 = nuke1

End Sub


Sub MDRIVE()
Err.Clear
On Error Resume Next

DRI = "M:\"
Set G = FS.GetDrive(DRI)
If G.ISREADY = False Then Exit Sub

If Err.Number > 0 Then Exit Sub
nuke2 = G.freespace
'If nuke2 < MHD2 And nuke2 <> MHD2 Then E$ = "ã" Else E$ = "ä"
If nuke2 < MHD2 Then E$ = "ä"
If nuke2 > MHD2 And MHD2 > 0 Then E$ = "ã"
If nuke2 = MHD2 Then E$ = Label12.Caption
If MHD2 = 0 Then E$ = Label12.Caption
Label8.Caption = Format$(nuke2 / 1024 ^ 3, "0.00000")
Label12.Caption = E$

If MHD2 <> nuke2 And MHD2 > 0 Then
Label9.BackColor = Label2.BackColor
Else
Label9.BackColor = Label1.BackColor
End If
MHD2 = nuke2

End Sub


Sub NDRIVE()
'GoTo jump3
Err.Clear
On Error Resume Next
DRI = "X:\"
Set G = FS.GetDrive(DRI)
If G.ISREADY = False Then Exit Sub

If Err.Number > 0 Then Exit Sub

nuke3 = G.freespace
If nuke3 < MHD3 Then E$ = "ä"
If nuke3 > MHD3 And MHD3 > 0 Then E$ = "ã"
If nuke3 = MHD3 Then E$ = Label7.Caption
If MHD3 = 0 Then E$ = Label7.Caption
Label5.Caption = Format$(nuke3 / 1024 ^ 3, "0.00000")
Label7.Caption = E$

If MHD3 <> nuke3 And MHD3 > 0 Then
Label6.BackColor = Label2.BackColor
Else
Label6.BackColor = Label1.BackColor
End If
MHD3 = nuke3

End Sub


Sub DDRIVE()

Err.Clear
On Error Resume Next
DRI = "D:\"
Set G = FS.GetDrive(DRI)
On Error Resume Next
If G.ISREADY = False Then Exit Sub
If Err.Number > 0 Then Exit Sub
On Error GoTo 0

If Err.Number > 0 Then Exit Sub

Nuke7 = G.freespace
If Nuke7 < MHD4 Then E$ = "ä"
If Nuke7 > MHD4 And MHD4 > 0 Then E$ = "ã"
If Nuke7 = MHD4 Then E$ = Label20.Caption
If MHD4 = 0 Then E$ = Label20.Caption
Label22.Caption = Format$(Nuke7 / 1024 ^ 3, "0.00000")
Label20.Caption = E$

If MHD4 <> Nuke7 And MHD4 > 0 Then
    Label21.BackColor = Label2.BackColor
Else
    Label21.BackColor = Label1.BackColor
End If
MHD4 = Nuke7


End Sub

Private Sub Timer2_Timer()

If TuLip = True Then GoTo jump
iio1 = Minute(Now)
iio1 = Second(Now)
'iio1 = 10
If (iio1 Mod 10 <> 0 Or IIO2 = iio1) And IIO2 <> -1 Then Exit Sub
IIO2 = iio1

jump:
On Error GoTo ErrHand


Call CDRIVE

    

Call MDRIVE

If TuLip = True Then GoTo Jump2


iio1 = Hour(Now)
'iio1 = 10
If (iio1 Mod 2 <> 0 Or IIO3 = iio1) And IIO3 <> -1 Then Exit Sub
IIO3 = iio1
Jump2:
TuLip = False

JUMP4:


Call NDRIVE

jump3:




Call DDRIVE




JUMP5:



'TOAL SIZES
Err.Clear
On Error Resume Next



Call DrivesGB










Exit Sub
ErrHand:
Exit Sub

End Sub

Sub DrivesGB()



Dim Nuke4, Nuke5, RR, icount



Set FS = CreateObject("Scripting.FileSystemObject")
Dim d
On Local Error Resume Next
Set dc = FS.Drives
For Each d In dc
    Err.Clear
    DR = d.DriveLetter
    DRX = d.DriveType

    



    
    IDD = 1
    
    'If DR = "Z" Then IDD = 1
    'If DRX = 1 Or DRX = 4 Or DRX = 3 Then IDD = 1
    
    If DRX = 2 Then IDD = 0
    
    If IDD = 0 Then n = d.VolumeName


    'Select Case d.DriveType
    'Case 0: t = "Unknown"
    'Case 1: t = "Removable"
    'Case 2: t = "Fixed"
    'Case 3: t = "Network"
    'Case 4: t = "CD-ROM"
    'Case 5: t = "RAM Disk"


    'IDD = 0
    If InStr(n, "MatthewLan") > 0 Then IDD = 1
    If InStr(n, "-TRU") > 0 Then IDD = 1
    
    
    If IDD = 0 Then
    
        If InStr(n, "BAK") > 0 And Err.Number = 0 Then
            bakspace = bakspace + d.totalsize / 1024 ^ 3
        End If



        icount = icount + 1
        RR = RR + d.VolumeName + "-"
        Nuke4 = Nuke4 + d.totalsize / 1024 ^ 3
        Nuke5 = Nuke5 + d.freespace / 1024 ^ 3
    End If
Next

If Nuke5 < TotHd Then E$ = "ä"
If Nuke5 > TotHd And TotHd > 0 Then E$ = "ã"
If Nuke5 = TotHd Then E$ = Label14.Caption
If TotHd = 0 Then E$ = Label14.Caption
Label16.Caption = Format$(Nuke5, "0.000000")
Label14.Caption = E$
Label19.Caption = Format$(Nuke4, "0.00000")


If TotHd <> Nuke5 And TotHd > 0 Then
Label15.BackColor = Label2.BackColor
Else
Label15.BackColor = Label1.BackColor
End If
TotHd = Nuke5



'TTz6 = Trim(str(icount)) + " Drive's " + Format(Nuke4 / 1024, "0.00") + " TByte. Minus " + Format(bakspace / 1024, "0.00") + " TB Backup " + Format(Nuke5, "0") + "Mb Free"


End Sub



Private Sub Timer3_Timer()
    
    'Exit Sub
    
    Label28 = QEmptyRecycleBinSize
    
    
    'frmOperatingSystem.ShowForm
    'Exit Sub
    Exit Sub
    If Label3.Enabled = False Then Exit Sub
    'Dim Enumerator As SWbemObjectSet
    'Dim Object As SWbemObject
    'Dim Item As ListItem

    On Error Resume Next
    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_OperatingSystem")
        
      '  freemem2 = Enumerator.FreePhysicalMemory
        
    For Each Object In Enumerator
        freemem2 = Object.FreePhysicalMemory
'        Label3 = Format$(freemem2 / 1024, "0.0000") + "Kb"
        Label3 = Trim(Str(freemem2))
    Next

End Sub


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


