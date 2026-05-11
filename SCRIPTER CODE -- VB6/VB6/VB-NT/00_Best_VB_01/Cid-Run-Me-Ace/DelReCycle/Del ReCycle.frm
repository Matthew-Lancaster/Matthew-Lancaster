VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Del Recycle Bin -- Ask"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   9315
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   5865
      Top             =   1755
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5355
      Top             =   1755
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4785
      Top             =   1755
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Size 0000 MB "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   5235
      TabIndex        =   2
      Top             =   0
      Width           =   3690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Del No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   3030
      TabIndex        =   1
      Top             =   15
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Del Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   405
      TabIndex        =   0
      Top             =   -15
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4

Dim SizeR

Private Type ULARGE_INTEGER
  lowpart As Long
  highpart As Long
End Type

Private Type SHQUERYRBINFO
  cbSize As Long
  i64Size As ULARGE_INTEGER
  i64NumItems As ULARGE_INTEGER
End Type


Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Private Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long


Dim LockMove, ORT

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim Rect8 As RECT

Private Declare Function MoveWindow _
        Lib "user32" _
        (ByVal hWnd As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long

Dim Xxr, Xxr48



Private Sub Form_Load()
Xxr = FindWindow("CabinetWClass", "Recycle Bin")
If Xxr = 0 Then End

If Xxr > 0 And Xxr48 <> Xxr Then
    Xxr48 = Xxr
End If

Label1.Left = 0: Label2.Left = Label1.Width
Label3.Left = Label2.Left + Label2.Width

On Error Resume Next
For Each Control In Controls
    Control.AutoSize = False
    If Control.Width + Control.Left > xx Then xx = Control.Width + Control.Left
    If Control.Height + Control.Top > yy Then yy = Control.Height + Control.Top
Next
On Error GoTo 0

Form1.Width = xx + 100
Form1.Height = yy + 400
i = GetWindowRect(Xxr, Rect8)
Me.Left = Rect8.Left * Screen.TwipsPerPixelX
Me.Top = (Rect8.Top * Screen.TwipsPerPixelY) - Me.Height

Call EmptyRecycleBinQ
If SizeR = 0 Then End
Form1.Show
End Sub




Private Sub Label1_Click()
Xxr = FindWindow("CabinetWClass", "Recycle Bin")
If Xxr = 0 Then End
'Call CloseProgramHwnd(Xxr)

'AppActivate "Recycle Bin", False
'Sleep 100
'SendKeys ("%{F4}"), True
Timer3.Enabled = True

'Call EmptyRecycleBin
'Call EmptyRecycleBin("C:\Recycled", True)
EmptyRecycleBin
'End

End Sub

Private Sub Label2_Click()

Xxr = FindWindow("CabinetWClass", "Recycle Bin")
'Call CloseProgramHwnd(Xxr)
AppActivate "Recycle Bin", False
Sleep 100
SendKeys ("%{F4}"), True
End
End Sub

Private Sub Timer1_Timer()
If Timer3.Enabled = True Then Exit Sub
Xxr = FindWindow("CabinetWClass", "Recycle Bin")
If Xxr = 0 And Timer3.Enabled = False Then
    End
End If

If GetAsyncKeyState(1) Then
    LockMove = 1
End If


i = GetWindowRect(Xxr, Rect8)
If Rect8.Left + Rect8.Top <> ORT Then LockMove = 0
ORT = Rect8.Left + Rect8.Top
If LockMove = 0 Then
Me.Left = Rect8.Left * Screen.TwipsPerPixelX
Me.Top = (Rect8.Top * Screen.TwipsPerPixelY) - Me.Height
End If



End Sub

Sub EmptyRecycleBinQ()  'Empty the recycle bin
    Dim RBinInfo As SHQUERYRBINFO, msg As VbMsgBoxResult
    RBinInfo.cbSize = Len(RBinInfo)
    SHQueryRecycleBin vbNullString, RBinInfo
'    If (RBinInfo.i64Size.lowpart And &H80000000) = &H80000000 Or RBinInfo.i64Size.highpart > 0 Then
'        msg = MsgBox("Your Recycle Bin consumes over 2 GB right now!" + vbCrLf + "Do you want to empty it?", vbYesNo + vbQuestion)
'    Else
'        msg = MsgBox("Your Recycle Bin consumes" + Str$((RBinInfo.i64Size.lowpart) / 1024) + " KB right now." + vbCrLf + "Do you want to empty it?", vbYesNo + vbQuestion)
'    End If
    If (RBinInfo.i64Size.lowpart And &H80000000) = &H80000000 Or RBinInfo.i64Size.highpart > 0 Then
        Label3 = "Over 2 GB"
        SizeR = (1024 ^ 3) * 2
    Else
        Label3 = Format(RBinInfo.i64Size.lowpart / 1024 ^ 2, "00.00") + " MB"
        SizeR = RBinInfo.i64Size.lowpart
    End If

End Sub

Sub EmptyRecycleBin()  'Empty the recycle bin
    Dim RBinInfo As SHQUERYRBINFO, msg As VbMsgBoxResult
    RBinInfo.cbSize = Len(RBinInfo)
    SHQueryRecycleBin vbNullString, RBinInfo
    hWndx = Screen.ActiveForm.hWnd
    
    NoConfirmation = True: NoProgress = False: NoSound = False
    flags = (NoConfirmation And SHERB_NOCONFIRMATION) Or (NoProgress And _
        SHERB_NOPROGRESSUI) Or (NoSound And SHERB_NOSOUND)

    'this coz vbnullstring will do all drives
    SHEmptyRecycleBin hWndx, vbNullString, flags
    SHUpdateRecycleBinIcon
End Sub


Public Sub CloseProgram(ByVal Handle As Long)

' Handle = FindWindow(vbNullString, Caption)
' If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub

Public Sub CloseProgramHwnd(ByVal Handle As Long)

 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub



Private Sub Timer2_Timer()
End
End Sub

Private Sub Timer3_Timer()
    
    If SizeR > 0 Then Call EmptyRecycleBinQ
    If SizeR > 0 Then Exit Sub
    On Error Resume Next
    AppActivate "Recycle Bin", False
    Sleep 100
    SendKeys ("%{F4}"), True
    Timer2.Enabled = False
    Timer2.Interval = 0
    Timer2.Interval = 10000
    Timer2.Enabled = True
End Sub

