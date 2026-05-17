Attribute VB_Name = "APISupport"
Option Explicit
'This module contains all the API stuff so it is easier to find/update
Private Const CB_FINDSTRING            As Long = &H14C
Private Const CB_FINDSTRINGEXACT       As Long = &H158
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Function EscPressed() As Boolean

  '*PURPOSE: detect Esc has been pressed

  EscPressed = (GetAsyncKeyState(vbKeyEscape) < 0)

End Function

Public Function PosInCombo(ByVal StrA As String, _
                           ByVal C As ComboBox, _
                           Optional CaseSensitive As Boolean = True) As Long

  'find if strA is in Combolist
  'returns -1 if not found

  PosInCombo = SendMessage(C.hWnd, IIf(CaseSensitive, CB_FINDSTRINGEXACT, CB_FINDSTRING), 0, ByVal StrA)

End Function

':) Roja's VB Code Fixer V1.1.2 (7/07/2003 2:03:18 AM) 6 + 20 = 26 Lines Thanks Ulli for inspiration and lots of code.

