VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Opus"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   945
   ClientWidth     =   1770
   Icon            =   "RecordGo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   1770
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   375
      Top             =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Opus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Enum WindowStates
  SW_HIDE = 0
  SW_SHOWNORMAL = 1
  SW_normal = 1
  SW_SHOWMINIMIZED = 2
  SW_SHOWMAXIMIZED = 3
  SW_MAXIMIZE = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  sw_RESTORE = 9
  SW_SHOWDEFAULT = 10
  SW_FORCEMINIMIZE = 11
  SW_MAX = 11
End Enum



Dim OldTiger, TigerX

Private Sub Form_Activate()
Form1.Show
Call Timer1_Timer

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End

OldTiger = -2
End Sub

Private Sub Label1_Click()
    
    Dim Tiger As Long
    Tiger = FindWindow("#32770", "Recording Window")
    If Tiger > 0 And WindowVisible(Tiger) = True Then
        WindowVisible(Tiger) = False
    Else
        WindowVisible(Tiger) = True
        'wr = Putfocus(HottMini) '' still dupious if need this one
        'ngI = SetFocuses(HottMini)
        'SetForegroundWindow Tiger
    End If
    'RecordWindow = Now + TimeSerial(0, 3, 0)
    Exit Sub

End Sub

Private Sub Timer1_Timer()
    Tiger = FindWindow("#32770", "Recording Window")
    If Tiger = 0 Then
        Form1.WindowState = 0
    End If
    If Tiger > 0 And OldTiger <> Tiger Then
        Form1.WindowState = 1
    End If
    OldTiger = Tiger

    If TigerX = 0 Then TigerX = FindWindow(vbNullString, "Digital Wave Player")
    If TigerX = 0 Then
        Call LoadRunAsWavePlayer
    End If
End Sub

Private Sub LoadRunAsWavePlayer()
'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt3"
    .sPassword = "matto"
    .sCommand = "C:\Program Files\Olympus\Digital Wave Player\DWP.exe"

'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End Sub




Property Let WindowVisible(hwnd As Long, New_Value As Boolean)
    If hwnd = 0 Then Exit Property
     
    Call ShowWindow(hwnd, IIf(New_Value, SW_SHOW, SW_HIDE))

End Property

Property Get WindowVisible(hwnd As Long) As Boolean

    If hwnd = 0 Then Exit Property
    WindowVisible = (IsWindowVisible(hwnd) > 0)

End Property


