Attribute VB_Name = "Module1"
'***********************************************
'     Sample for using the "dskbhook.dll"
'        (c) 2002 by Delphin Software
'***********************************************
Option Explicit
'***********************************************
Declare Function SetKeyboardHook& Lib "dskbhook" (ByVal hwnd&, ByVal Callback&, ByVal Adr&)
Declare Function SetGlobalParams& Lib "dskbhook" (ByVal Repeat&, ByVal Discard&)
Declare Function SetKeyParams& Lib "dskbhook" (ByVal KeyCode&, ByVal SCAE&, ByVal RDC&)
Declare Function HookVB& Lib "dskbhook" (ByVal Hook&)
Declare Function SkipInjected& Lib "dskbhook" (ByVal Skip&)
Declare Function NoTL& Lib "dskbhook" (ByVal TL&)
Declare Function IsNT& Lib "dskbhook" ()
Declare Function NoLLHook& Lib "dskbhook" (ByVal LLHook&)
'***********************************************
'# This function is called from the DLL in general
Public Function Callback&(ByVal action&, ByVal state&, ByVal vcode&, ByVal scode&)
'# Display the passed values
With Form1.List1
  If .ListCount > 5000 Then .Clear
  '# Add the new entry at the first position and highlight it
  .AddItem "   " & action & Chr$(9) & state & Chr$(9) & vcode & Chr$(9) & scode, 0
  .Selected(.NewIndex) = True
End With
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

