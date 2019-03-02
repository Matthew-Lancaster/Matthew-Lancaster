Attribute VB_Name = "Module1"
Public GoThrou
Public Kbbdf, KBPress2
'***********************************************
'     Sample for using the "dskbhook.dll"
'        (c) 2002 by Delphin Software
'***********************************************
Option Explicit
'***********************************************
Public Const KEYEVENTF_KEYUP& = &H2
Public Const KEYEVENTF_EXTENDEDKEY& = &H1
'-----------------------------------------------
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags&, ByVal dwExtraInfo&)
'-----------------------------------------------
Declare Function SetKeyboardHook& Lib "dskbhook" (ByVal hwnd&, ByVal Callback&, ByVal Adr&)
Declare Function SetGlobalParams& Lib "dskbhook" (ByVal Repeat&, ByVal Discard&)
Declare Function SetKeyParams& Lib "dskbhook" (ByVal KeyCode&, ByVal SCAE&, ByVal RDC&)
Declare Function UseSendMessage& Lib "dskbhook" (ByVal Send&)
Declare Function SkipInjected& Lib "dskbhook" (ByVal Skip&)
Declare Function NoLLHook& Lib "dskbhook" (ByVal LLHook&)
Declare Function SetStop& Lib "dskbhook" (ByVal Stp&)
Declare Function HookVB& Lib "dskbhook" (ByVal Hook&)
Declare Function NoTL& Lib "dskbhook" (ByVal tl&)
Declare Function IsNT& Lib "dskbhook" ()
'***********************************************
Public mode%, IsHook As Boolean
Dim Shift As Boolean, Ctrl As Boolean, Alt As Boolean
Dim Ext As Boolean, Up As Boolean, Down As Boolean

'***********************************************
'# This function is called from the DLL in general
Public Function Callback&(ByVal action&, ByVal state&, ByVal vcode&, ByVal scode&)
'# Display the passed values
With DSkeybd_F.List1
  If .ListCount > 500 Then .Clear
  '# Add the new entry at the first position and highlight it
  .AddItem "   " & action & Chr$(9) & state & Chr$(9) & vcode & Chr$(9) & scode, 0
  .Selected(.NewIndex) = True
End With

'DPrint action, state, vcode, scode
Down = action And 2: Up = action And 4: Ext = action And 1
Shift = state And 1: Ctrl = state And 2: Alt = state And 4

Callback = 0 '#never suppress any keys yet

If action = 4 Or action = 5 Then Exit Function

'On Error Resume Next
'Open "D:\TEMP\KeyHit-Tidal.txt" For Output As #1
'Print #1, Label23.Caption
'Close #1
Kbbdf = True
KBPress2 = vcode
Call CID_Run_Me.Lastpress

End Function

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

'***********************************************

