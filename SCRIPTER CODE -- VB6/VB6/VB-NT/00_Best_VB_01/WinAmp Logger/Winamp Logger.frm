VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label Label7 
      Caption         =   "Label1"
      Height          =   270
      Left            =   540
      TabIndex        =   0
      Top             =   210
      Width           =   930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112
Private Const ipc_isplaying As Long = 104
Private Const IPC_GETPLAYLISTFILE  As Long = 211
Private Const IPC_GETOUTPUTTIME  As Long = 105
Private Const IPC_GETINFO As Long = 126
Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WM_CLOSE = &H10
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
Private Const GW_HWNDNEXT = 2



Private Sub Form_Load()


tf2 = FindWindow("Winamp v1.x", vbNullString)

MsgResult3 = SendMessage(tf2, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)
MsgResult2 = SendMessage(tf2, WM_USER, 0, ByVal IPC_GETOUTPUTTIME)
If MsgResult2 = 0 And MsgResult3 = 0 Then
'MsgBox "Not Playing": End
Shell "notepad " + App.Path + "\00 Music Info Logger.txt", vbNormalFocus
End
End If

beeh = (MsgResult2 / Int(MsgResult3 * 1000)) * 100
If Int(beeh) = 100 Then
    Label7 = Format$(beeh, "000") + "%"
Else
    Label7 = Format$(beeh, "0.0") + "%"
End If
If Label7 = "0.0%" Then
    Label7 = "0%"
End If



Dim oWinamp
Set oWinamp = CreateObject("WinampCOM.Application")
Txz1$ = oWinamp.CurrentSongFileName
Set oWinamp = Nothing

f2 = DateDiff("h", 0, MsgResult3) - 1
f3 = DateDiff("n", 0, MsgResult3) - 1
f4 = DateDiff("s", 0, MsgResult3)
f6 = Int(MsgResult3 / 3600)
f7 = Int(MsgResult3 / 60)
f8 = MsgResult3 Mod 60
'Total
lab$ = Format$(TimeSerial(0, f7, f8), "HH:MM:SS")
If Mid$(lab$, 1, 1) = "0" Then lab$ = Mid$(lab$, 2)
If Mid$(lab$, 1, 1) = "0" Then lab$ = Mid$(lab$, 3)
Label4 = lab$


Hg = Int(MsgResult2 / 1000)
f6 = Int((Hg / 60 ^ 2))
f7 = Int(Hg / 60)
f8 = Hg Mod 60
Xtop2 = TimeSerial(0, f7, f8)

'Position Now

lab$ = Format$(Xtop2, "HH:MM:SS")
If Mid$(lab$, 1, 1) = "0" Then lab$ = Mid$(lab$, 2)
If Mid$(lab$, 1, 1) = "0" Then lab$ = Mid$(lab$, 3)
Label3v = Label3
Label3 = lab$
totter = DateDiff("s", Xtop2, Xtop)
Hg = Int(totter)
f6 = Int((Hg / 60 ^ 2))
f7 = Int(Hg / 60)
f8 = Hg Mod 60
totter = Now + TimeSerial(0, f7, f8)
Label18 = Format$(totter, "HH:MM:SS")


tf2 = FindWindow("Winamp v1.x", vbNullString)

Label9 = ttmd$

If tf2 = 0 Then Exit Sub

Rfg$ = GetWindowTitle(tf2)
rfh1 = InStrRev(Rfg$, " - ")
If rfh1 > 0 Then Rfg$ = Mid$(Rfg$, 1, rfh1 - 1)

If tf2 > 0 Then Label11 = Rfg$



'Exit Sub

gug = FreeFile
Open App.Path + "\00 Music Info Logger.txt" For Append As #gug
Print #gug,
Print #gug, String$(70, "-")
Print #gug, Format$(Now, "DDD DD-MM-YYYY HH:MM:SS")
Print #gug, String$(70, "-")
Print #gug, "Play % Now = " + Label7
Print #gug, "Play Lenght = " + Label4
Print #gug, "Play Now = " + Label3
Print #gug, String$(70, "-")
Print #gug, "File Name = " + Txz1$
Print #gug, "Song Name = " + Label11
Print #gug, String$(70, "-")
Close #gug

Shell "notepad " + App.Path + "\00 Music Info Logger.txt", vbNormalFocus



End
End Sub



Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function

