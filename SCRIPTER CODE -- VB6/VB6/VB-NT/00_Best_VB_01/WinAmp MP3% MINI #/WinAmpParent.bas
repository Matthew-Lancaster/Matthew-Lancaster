Attribute VB_Name = "WinAmpParent"
'Option Explicit
Public Day1, Day2, OldDay, LastURL, CuteURL, Writing

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1
Public vFileName As String

Private Const WM_GETTEXTLENGTH = &HE
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Public OldWinTitle, OldAsh, Ash$
Public TrueExit
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) _
    As Long
    
Private Const WM_SETTEXT            As Long = &HC
Private Const WM_GETTEXT = &HD

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142




Function WinAmpParent(Test_Hwnd) As Long
    
    Dim hWndForm        As Long
    Dim hWndTextBox     As Long
    Dim RetVal As Long
    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
    Dim WinClass As String, WinTitle2 As String
    'Dim Ash$
    
    'URL = "": WinTitle = ""
    
    'ttg = 0
    
    'hWndFormi = FindWindow("IEFrame", vbNullString)
    hWndForm = Test_Hwnd
    
    'If hWndFormi = GetForegroundWindow Then ttg = 1
    'hWndFormE = FindWindow("ExploreWClass", vbNullString)
    'If hWndFormE = GetForegroundWindow Then ttg = 2
    'If ttg = 1 Then hWndForm = hWndFormi
    'If ttg = 2 Then hWndForm = hWndFormE
    
    'If ttg = 0 Then
    '    Exit Sub
    'End If
    
    
    
    Dim hWndIEChild As Long
    
    hWndIEChild = FindWindowEx(hWndForm, 0&, "Winamp v1.x", vbNullString)
    'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ReBarWindow32", vbNullString)
    'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Address Band Root", vbNullString)
    'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBoxEx32", vbNullString)
    'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBox", vbNullString)
    'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Edit", vbNullString)
    
    WinAmpParent = hWndIEChild
    
    Exit Function
    
    If hWndIEChild > 0 Then
        
    window_hwnd = hWndIEChild
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
        
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, 255, ByVal WinTitleBuf)
    WinTitle = StripNulls(WinTitleBuf)
    'Ash$ = GetWindowTitle(hWndForm)
    
End If
    
End Function

Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
End Function

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim s As String
   l = GetWindowTextLength(hWnd)
   s = Space(l + 1)
   GetWindowText hWnd, s, l + 1
   GetWindowTitle = Left$(s, l)

End Function

