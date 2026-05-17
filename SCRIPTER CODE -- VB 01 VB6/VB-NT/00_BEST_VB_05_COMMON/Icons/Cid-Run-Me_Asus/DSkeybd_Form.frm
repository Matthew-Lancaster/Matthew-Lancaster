VERSION 5.00
Begin VB.Form DSkeybd_F 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " DS KEYSTROKE ANALYZER"
   ClientHeight    =   6105
   ClientLeft      =   1350
   ClientTop       =   1605
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   3210
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   135
      TabIndex        =   15
      Top             =   5490
      Width           =   2970
      Begin VB.CheckBox Check2 
         Caption         =   "NoTL"
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   180
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use LLHook,   if possible"
         Height          =   375
         Left            =   60
         TabIndex        =   16
         Top             =   135
         Value           =   1  'Checked
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   330
      Left            =   1215
      TabIndex        =   14
      Top             =   5130
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   330
      Left            =   2250
      TabIndex        =   13
      Top             =   5130
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unhook"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2250
      TabIndex        =   10
      Top             =   4545
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hook"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2250
      TabIndex        =   9
      Top             =   4155
      Width           =   825
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   1215
      TabIndex        =   6
      Top             =   4050
      Width           =   1005
      Begin VB.OptionButton Option2 
         Caption         =   "No Rep."
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   8
         Top             =   495
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Repeat"
         Height          =   240
         Index           =   1
         Left            =   45
         TabIndex        =   7
         Top             =   765
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Repeat"
         Height          =   240
         Index           =   1
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   135
      TabIndex        =   2
      Top             =   4050
      Width           =   1005
      Begin VB.OptionButton Option1 
         Caption         =   "Chain"
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Top             =   540
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Thread"
         Height          =   240
         Index           =   1
         Left            =   45
         TabIndex        =   4
         Top             =   810
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Discard"
         Height          =   240
         Index           =   2
         Left            =   45
         TabIndex        =   3
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Go To"
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   900
      End
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   135
      TabIndex        =   0
      Top             =   450
      Width           =   2940
   End
   Begin VB.Label Label1 
      Caption         =   "Action     State      VCode     SCode"
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   2670
   End
End
Attribute VB_Name = "DSkeybd_F"
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
Private Declare Function SetWindowPos& Lib "user32" _
(ByVal hWnd&, ByVal hWndInsertAfter&, ByVal x&, _
ByVal y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
'***********************************************
Dim dl&, dc&, rp&, i%, o2%, IsHook As Boolean
'***********************************************
Private Sub Check1_Click()
With Check2
  If Check1 = 0 Then
    .Value = 0
    .Enabled = False
  Else
    .Value = 1
    .Enabled = True
  End If
End With
End Sub

'***********************************************
Private Sub Check2_Click()
Call NoTL(Check2.Value)
End Sub

'***********************************************
Private Sub Command1_Click()
Call SetGlobalParams(rp, dc)                      '# Set Repeat and Discard value
If Check1 = 0 Then Call NoLLHook(1)               '# No LowLevel Hook
If Check2 = 1 Then Call NoTL(1)                   '# No Translation
dl = SetKeyboardHook(-1, 1, AddressOf Callback)   '# Set the hook
'Debug.Print dl
IsHook = True
DSkeybd_F.Check1.Enabled = False
DSkeybd_F.Command1.Enabled = False
DSkeybd_F.Command2.Enabled = True
'DSkeybd_F.List1.SetFocus
End Sub

'***********************************************
Private Sub Command2_Click()
dl = SetKeyboardHook(0, 0, 0)                     '# Release the hook
IsHook = False
Command1.Enabled = True
Command2.Enabled = False
Check1.Enabled = True
End Sub

'***********************************************
Private Sub Command3_Click()
Unload Me
End Sub

'***********************************************
Private Sub Command4_Click()
With List1
  .Clear
  .SetFocus
End With
End Sub

Private Sub Form_Activate()


'Check2.Value = 0
'Call Check2_Click
'Check1.Value = 0
'Call Check1_Click


Call Option2_Click(0)
Call Command1_Click

'Hide

End Sub

'***********************************************
Private Sub Form_Load()
Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                  SWP_NOMOVE Or SWP_NOSIZE)

'# Check the OS version
If IsNT = 0 Then                                  '# Win9X or NT4 SP < 3
  Check1.Value = 0
  Check1.Enabled = False
  Check2.Value = 0
  Check2.Enabled = False
End If

'# In the IDE, the forwarding of
'# keystrokes will be suppressed
If IsIDE Then
  ChDrive Left$(App.Path, 1)                      '# Let VB find the DLL, if the VB.exe is
  ChDir App.Path                                  '# located on another drive than the project.
  Option1(2) = True
  dc = 2                                          '# Discard the keystrokes, when IDE
  For i = 0 To Option1.Count - 1
    Option1(i).Enabled = False
  Next
  Call HookVB(1)                                  '# Explicitely get callback from VB
End If

'Show

End Sub

'***********************************************
Private Sub Form_Unload(Cancel As Integer)
If IsHook Then Call SetKeyboardHook(0, 0, 0)      '# Release the hook when closing
End Sub

'***********************************************
Private Sub Option1_Click(Index As Integer)
For i = 0 To Option1.Count - 1
  If Option1(i) Then Exit For
Next
dc = i                                            '# Get the Discard value to set
Call SetGlobalParams(rp, dc)                      '# Set the (new) Discard value
End Sub

'***********************************************
Private Sub Option2_Click(Index As Integer)

Option2(1) = True

For i = 0 To Option2.Count - 1
  If Option2(i) Then Exit For                     '# Get the Repeat value
Next

rp = i
Call SetGlobalParams(rp, dc)                      '# Set the (new) Repeat value
If Not IsHook Then Command1.Enabled = True        '# Enable hooking
End Sub

'***********************************************

