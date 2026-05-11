Attribute VB_Name = "DsKeybd_M2"
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
Declare Function NoTL& Lib "dskbhook" (ByVal TL&)
Declare Function IsNT& Lib "dskbhook" ()
'***********************************************
Public mode%, IsHook As Boolean
Dim Shift As Boolean, Ctrl As Boolean, Alt As Boolean
Dim Ext As Boolean, Up As Boolean, Down As Boolean
'***********************************************
Sub Main()
Load Form1
End Sub

'***********************************************
Sub SetHook()
Select Case mode
Case 0: If IsNT = 0 Then MsgBox "No NT Windows!", 64, " DSKBHOOK:": Exit Sub  '# Check whether NT
        Call SetGlobalParams(0, 2)                          '# Discard all keystrokes
Case 1: Call NoLLHook(1): Call SetGlobalParams(0, 2)        '# Install normal keyboard hook
Case 2: Call SetKeyParams(48, 0, 16)                         '# Suppress numbers
        Call SetKeyParams(49, 0, 16)
        Call SetKeyParams(50, 0, 16)
        Call SetKeyParams(51, 0, 16)
        Call SetKeyParams(52, 0, 16)
        Call SetKeyParams(53, 0, 16)
        Call SetKeyParams(54, 0, 16)
        Call SetKeyParams(55, 0, 16)
        Call SetKeyParams(56, 0, 16)
        Call SetKeyParams(57, 0, 16)
Case 3: Call UseSendMessage(1)                              '# Enable discarding from callback
Case 4: Call SetGlobalParams(0, 2): Call SkipInjected(1)    '# Discard all, skip injected keystrokes
End Select
Call SetKeyboardHook(-1, 1, AddressOf Callback)             '# Install system wide keyboard hook
Call HookVB(1)                                              '# Explicitely get callback from VB
IsHook = True
End Sub

'***********************************************
'# This function is called from the DLL in general
Public Function Callback&(ByVal action&, ByVal state&, ByVal vcode&, ByVal scode&)
'DPrint action, state, vcode, scode
Down = action And 2: Up = action And 4: Ext = action And 1
Shift = state And 1: Ctrl = state And 2: Alt = state And 4

Select Case mode
Case 0:
Case 1:
Case 2:
Case 3: If vcode = 65 And Shift Then Callback = 2           '# Suppress "a" when shifted
Case 4: If Down Then Call SendKey(vcode + 1, state + IIf(Ext, 8, 0)) '# Swap keys
End Select
End Function

'***********************************************
'# Generates a keystroke
Sub SendKey(ByVal vk%, ByVal mSCAE%)
Dim Flags&
'DPrint vk, mSCAE
Call SetStop(1)
If mSCAE And 1 Then Call keybd_event(16, 0, 0, 0) '# Shift
If mSCAE And 2 Then Call keybd_event(17, 0, 0, 0) '# Ctrl
If mSCAE And 4 Then Call keybd_event(18, 0, 0, 0) '# Alt
If mSCAE And 8 Then Flags = KEYEVENTF_EXTENDEDKEY '# Extended

Call keybd_event(vk, 0, Flags, 0)
Call keybd_event(vk, 0, Flags Or KEYEVENTF_KEYUP, 0)

Flags = KEYEVENTF_KEYUP
If mSCAE And 4 Then Call keybd_event(18, 0, Flags, 0)
If mSCAE And 2 Then Call keybd_event(17, 0, Flags, 0)
If mSCAE And 1 Then Call keybd_event(16, 0, Flags, 0)
Call SetStop(0)
End Sub

'***********************************************
Sub UnHook()
Call SetKeyboardHook(0, 0, 0)
IsHook = False
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

