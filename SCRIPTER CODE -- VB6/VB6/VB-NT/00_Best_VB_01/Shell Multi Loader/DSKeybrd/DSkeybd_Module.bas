Attribute VB_Name = "DSkeybd_M"
Public GoThrou
Public Actionz, Statez, Vcodez, Scodez

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long



'***********************************************
'     Sample for using the "dskbhook.dll"
'        (c) 2002 by Delphin Software
'***********************************************
'Option Explicit
'***********************************************
Public Const KEYEVENTF_KEYUP& = &H2
Public Const KEYEVENTF_EXTENDEDKEY& = &H1
'-----------------------------------------------
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags&, ByVal dwExtraInfo&)
'-----------------------------------------------
Declare Function SetKeyboardHook& Lib "dskbhook" (ByVal hWnd&, ByVal Callback&, ByVal Adr&)
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
    
If KeyClean = True Then Callback = 2: Exit Function
    
Actionz = action
Statez = state
Vcodez = vcode
Scodez = scode

Kbbdf5 = False

ATidalX.Label15.Caption = action & "-" & state & "-" & vcode & "-" & scode

If action = 4 Or action = 5 Then Exit Function 'key release

On Local Error Resume Next

Dim Nx As Long
Dim Ny As Long
ATidalX.FindCursor Nx, Ny
xx3 = 0
mwnd = WindowFromPoint(Nx, Ny)
'VB code window
If mwnd = FindWindow("wndclass_desked_gsk", vbNullString) Then xx3 = 1
If mwnd = FindWindow("#32770", "Microsoft Visual Basic") Then xx3 = 1
If mwnd = FindWindow("VbaWindow", vbNullString) Then xx3 = 1

mwnd = GetForegroundWindow
'VB code window
If mwnd = FindWindow("wndclass_desked_gsk", vbNullString) Then xx3 = 1
If mwnd = FindWindow("#32770", "Microsoft Visual Basic") Then xx3 = 1
If mwnd = FindWindow("VbaWindow", vbNullString) Then xx3 = 1

'If IsIDE = True Then xx3 = 0

If xx3 = 0 Then
'For Num Pad
'If Actionz = 2 And Statez = 0 And Vcodez = 188 Then Sluty3 = 3 ' , = Next Track
'If Actionz = 3 And Statez = 0 And Vcodez = 111 Then Sluty3 = 2 ' / = Pause

'For Num Pad
'If Actionz = 2 And Vcodez = 107 Then Sluty3 = 4 ' + = Vol Up
'If Actionz = 2 And Vcodez = 109 Then Sluty3 = 5 ' - =  Vol Down

If Actionz = 2 And Statez = 0 And Vcodez = 121 Then Sluty3 = 4    'F10 = 'Vol Up
If Actionz = 2 And Statez = 0 And Vcodez = 120 Then Sluty3 = 5  'F9 = 'Vol Down
'If Actionz = 2 And Statez = 0 And Vcodez = 27 Then Sluty3 = 5    'Esc = 'Vol Down

'------------------------------------------------------------------
If Actionz = 2 And Statez = 0 And Vcodez = 119 Then Sluty3 = 2 'F8 = Pause
If Actionz = 2 And Statez = 0 And Vcodez = 118 Then Sluty3 = 3 'F7 = Next Track

End If

'on Mobile Remote
If Actionz = 2 And Statez = 6 And Vcodez = 50 Then Sluty3 = 3 'F7 = Next Track
If Actionz = 2 And Statez = 6 And Vcodez = 49 Then Sluty3 = 2 'F8 = Pause


'If Actionz = 3 And (Statez = 0 Or Statez = 2) And Vcodez = 38 Then 'Up Arrow Mobile
'    Call KnockerLogg
'End If

'If Actionz = 2 And Vcodez = 106 Then '* Num Pad
    'Call KnockerLogg
'End If

'If kbi = 117  -- 'F6 = 1 Min delay Pause
'If Actionz = 2 And Statez = 0 And Vcodez = 117 Then MinDelayPause = Now + TimeSerial(0, 30, 0)

'Control Key
If Actionz = 2 And Statez = 2 And Vcodez = 118 Then Sluty3 = 3 'F7 = Next Track
If Actionz = 2 And Statez = 2 And Vcodez = 119 Then Sluty3 = 2 'F8 = Pause
If Actionz = 2 And Statez = 2 And Vcodez = 120 Then Sluty3 = 5    'F9 = 'Vol Down
If Actionz = 2 And Statez = 2 And Vcodez = 121 Then Sluty3 = 4    'F10 = 'Vol Up

'If Actionz = 2 And Statez = 0 And Vcodez = 122 Then Sluty4 = 1      'F11 = Mark Messages By Conversation
If Actionz = 2 And Vcodez = 109 Then Sluty4 = 1      '- Num Pad = Mark Messages By Conversation

If vcode = 123 Then 'f12
    If FindWindow("IEFrame", vbNullString) = GetForegroundWindow And _
        GetWindowState(FindWindow("IEFrame", vbNullString)) <> vbMinimized Then
     Sluty = 1    ' F12
    Else
        tiger = FindWindow("Notepad2", vbNullString)
        hh$ = GetWindowTitle(FindWindow("Notepad2", vbNullString))
        If InStr(hh$, "RutLand Hosso.txt - Notepad2") > 0 Then xtiger = 1 Else xtiger = 0
        
        edge = 0
        If xtiger = 1 Then
            If GetWindowState(tiger) = vbMinimized Then
                ShowWindow tiger, SW_SHOWNORMAL
                edge = 1
            End If
            If GetWindowState(tiger) = vbMaximized Or GetWindowState(tiger) = -1 Then
                If edge = 0 Then ShowWindow tiger, SW_MINIMIZE
            End If
        Else
            Shell "D:\VB6\VB-NT\00_Best_VB_01\Shell NotePad Loader\Shell NotePad Loader.exe Load_RutLand"
        End If
    End If
End If
        
'If GetForegroundWindow = FindWindow("Notepad2", vbNullString) Then
'    If vcode = 115 Then Sluty = 1 'f4 Copy - Ie Paste
'End If
        
'If vcode = 44 Then Sluty3 = 8 '* Print Screen Pinnos


'Check ----------------- Timer8WinAmpVol_Timer
If Sluty3 > 0 Then Call KeyWindowCheck
If Sluty > 0 Then Call KeyWindowCheck
If Sluty4 > 0 Then Call KeyWindowCheck

'Work to Turn Lock on an Off
If vcode = 44 And state = 2 Then Call ATidalX.Bollocks
If vcode = 44 And state = 3 Then Call ATidalX.Bollocks


If vcode = 111 And action = 3 Then
    Call SendKey(220, 0) '# Swap keys /=\
    Callback = 2
End If

If vcode = 107 And (xx3 = 1 Or ATidalX.hWnd = GetForegroundWindow) Then
    Call SendKey(3, 2) '# Swap NumPad + to Break
    Callback = 2
End If

If vcode = 122 And action = 2 Then
    Call SendKey(65, 2 + IIf(Ext, 8, 0)) '# inject control A to select all
'    Call SendKey(67, 2)  '# inject control C to Copy to Clipboard
    DSkeybd_F.Timer1.Enabled = True
    Callback = 2
End If

If FindWindow("Notepad2", vbNullString) > 0 Then
If vcode = 118 Or vcode = 119 Or vcode = 120 Then Callback = 2
End If

hh$ = GetWindowTitle(GetForegroundWindow)
If InStr(hh$, "IceView (Unregistered)") > 0 Then
    If vcode = 118 Then Callback = 2
End If

tiger = FindWindow("Notepad2", vbNullString)
If tiger = 0 Then tiger = FindWindow("Scintilla", vbNullString)
'If tiger = 0 Then GetForegroundWindow = FindWindow("Notepad2", vbNullString)
If tiger = GetForegroundWindow And vcode = 123 Then
    Callback = 2
End If



Keyy = Keyy + 1

ATidalX.Label23.Caption = Str$(Keyy)

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
'# This function is called from the DLL in general
Public Function Callback33&(ByVal action&, ByVal state&, ByVal vcode&, ByVal scode&)
'DPrint action, state, vcode, scode
Down = action And 2: Up = action And 4: Ext = action And 1
Shift = state And 1: Ctrl = state And 2: Alt = state And 4

Select Case mode
Case 0:
Case 1:
Case 2:
Case 3: If vcode = 65 And Shift Then Callback33 = 2           '# Suppress "a" when shifted
Case 4: If Down Then Call SendKey(vcode + 1, state + IIf(Ext, 8, 0)) '# Swap keys
End Select
End Function


