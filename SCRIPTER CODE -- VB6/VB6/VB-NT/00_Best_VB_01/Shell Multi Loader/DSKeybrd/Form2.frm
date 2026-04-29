VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample for using the DSKBHOOK.DLL"
   ClientHeight    =   3795
   ClientLeft      =   4575
   ClientTop       =   1845
   ClientWidth     =   3870
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   3870
   Begin VB.CommandButton Command2 
      Caption         =   "Release hook"
      Height          =   510
      Left            =   225
      TabIndex        =   5
      Top             =   3105
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Swap keys using ""keybd_event"""
      Height          =   510
      Index           =   4
      Left            =   225
      TabIndex        =   4
      Top             =   2520
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Suppress ""a"" when shifted from callback"
      Height          =   510
      Index           =   3
      Left            =   225
      TabIndex        =   3
      Top             =   1935
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Suppress numbers using ""SetKeyParams()"""
      Height          =   510
      Index           =   2
      Left            =   225
      TabIndex        =   2
      Top             =   1350
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Suppress all keystrokes (Normal Hook)"
      Height          =   510
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   765
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Suppress all keystrokes (LL Hook)"
      Height          =   510
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   3435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'     Sample for using the "dskbhook.dll"
'        (c) 2002 by Delphin Software
'***********************************************
Option Explicit
'***********************************************
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVESIZE& = SWP_NOMOVE Or SWP_NOSIZE
Private Const SWP_NOACTIVATE& = &H10
Private Const SWP_SHOWWINDOW& = &H40
Private Const SWP_SHOWNOACTIVATE& = SWP_SHOWWINDOW Or SWP_NOACTIVATE
Private Declare Function SetWindowPos& Lib "user32" _
(ByVal hwnd&, ByVal hWndInsertAfter&, ByVal x&, _
ByVal y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" _
                (ByVal lpClassName$, ByVal lpWindowName$)
Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hwnd&)
'***********************************************
Dim i%, hNote&, BCount%
'***********************************************
Private Sub Command1_Click(Index%)
If IsHook Then Call UnHook
For i = 0 To BCount
  Command1(i).Enabled = False
Next
mode = Index
Call SetHook
Call SetForegroundWindow(hNote)
End Sub

'***********************************************
Private Sub Command2_Click()
If IsHook Then Call UnHook
For i = 0 To BCount
  Command1(i).Enabled = True
Next
End Sub

'***********************************************
Private Sub Form_Load()

Call Shell("notepad", 0)                          '# Launch Notepad
hNote = FindWindow("notepad", vbNullString)
Call SetWindowPos(hNote, 0, 0, 100, 300, 400, SWP_SHOWWINDOW)

Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
                  SWP_NOMOVESIZE Or SWP_SHOWNOACTIVATE)

If IsIDE Then
  ChDrive Left$(App.Path, 1)                      '# Let VB find the DLL, if the VB.exe is
  ChDir App.Path                                  '# located on another drive than the project.
End If
BCount = Command1.Count - 1
End Sub

'***********************************************
Private Sub Form_Unload(Cancel As Integer)
If IsHook Then Call UnHook                        '# Release the hook when closing
End Sub

'***********************************************
