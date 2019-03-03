Attribute VB_Name = "FindWindowDLL"





'----
'FreeVBCode code snippet: Get hWnd of All Visible Windows on Screen.
'http://www.freevbcode.com/ShowCode.asp?ID=8408
'----
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long





'----
'Closing a firefox tab with Visual Basic (and JavaScript if needed)
'https://www.experts-exchange.com/questions/23469794/Closing-a-firefox-tab-with-Visual-Basic-and-JavaScript-if-needed.html
'----

Private Declare Function EnumWindows Lib _
   "user32" (ByVal lpEnumFunc As Long, lParam _
   As Any) As Long

Private Declare Function _
   GetWindowThreadProcessId Lib "user32" (ByVal _
   hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function GetWindowLong Lib _
   "user32" Alias "GetWindowLongA" (ByVal hwnd _
   As Long, ByVal nIndex As Long) As Long

'Private Declare Function PostMessage Lib _
   "user32" Alias "PostMessageA" (ByVal hWnd As _
   Long, ByVal wMsg As Long, ByVal wParam As _
   Long, lParam As Any) As Long

Private Const GWL_HWNDPARENT = (-8)

'Private Const WM_CLOSE = &H10



' Initiate PostCloseMessage callbacks like this:

' Call EnumWindows(AddressOf PostCloseMessage,

' ByVal ProcessID)

'



Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long  'MODULE 1135



'DECLARATIONS:

'------------------------------------------
'----
'Get browser tab captions & close tabs-VBForums
'http://www.vbforums.com/showthread.php?691711-Get-browser-tab-captions-amp-close-tabs
'----

'Public Declare Function BeepAPI Lib "kernel32" Alias "Beep" (ByVal dwFrequency As Long, ByVal dwMilliseconds As Long) As Long
'Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Public Declare Function GetActiveWindow Lib "user32" () As Long
'Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
'Public Const GW_HWNDNEXT = 2
'Public Const WM_CLOSE = &H10

'------------------------------------------



Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  
  'MODULE 1135

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long



Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hwnd)
   S = Space(L + 1)
   GetWindowText hwnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function



Function FindWindowPart(Test) As Long

'Testing
'Test$

FindWindowPart = False

'Dim RECT8 As RECT

'Dim Ash$
Dim test_hwnd As Long

'dim test_pid As Long, _
    test_thread_id As Long

'Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one

'Dim Huge

test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
'br$ = ""
Do While test_hwnd <> 0
    
    'hwnd9 = GetWindowRect(test_hwnd, RECT8)
        
     '   Ash$ = GetWindowTitle(test_hwnd)
'        If InStr(ash$, "Double") > 0 Then Stop
'        gws = GetWindowState(test_hwnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"

    
    'If (RECT8.Top > 0 And RECT8.Left > 0) Or gws = vbMinimized Then
        
        If InStr(GetWindowTitle(test_hwnd), Test) > 0 Then
            FindWindowPart = test_hwnd
            'Exit Do
        End If
        
        'If Ash$ <> "" Then 'And InStr(br$, "-- " + Ash$ + " -- ") = 0 Then
            'br$ = br$ + "-- " + Ash$ + " -- "
            'On Error Resume Next
            'AppActivate Ash$
            'On Error GoTo 0
            'Huge = Huge + 1
            'ef = SetForegroundWindow(hwnd9)
            'ef = Putfocus(hwnd9)
'            ecute = Timer + 0.1
'
'            Do
'                'Safety IN CASE TME RESETS BACK TO ZERO
'                If Timer < ecute - 30 Then Exit Do
'
'            Loop Until Timer > ecute
        
        'End If
    'End If
        
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

'If Q = False Then MsgBox Str(Huge) + " Windows Brought To Front"

'FindWinPartFront = Huge

End Function

Function FindWindowPartVB(Test) As Long

'Testing
'Test$
'TEST$ 123

'VB WANTS TWO FOR MAIN FORM

V1 = 0: V2 = 0

FindWindowPartVB = False

Dim test_hwnd As Long

test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
Do While test_hwnd <> 0
        
        If InStr(GetWindowTitle(test_hwnd), Test) > 0 Then
'            If V2 > 0 Then MsgBox "TOO MANY VALUE FOR FindWindowPartVB"
'            If V1 > 0 Then V2 = test_hwnd
'            If V1 = 0 Then V1 = test_hwnd
'
            FindWindowPartVB = test_hwnd
            Exit Function

        End If
        
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop


End Function


Function FindWindowPart_BASEBAR(Test) As Long

'Testing
'Test$
'TEST$ 123



FindWindowPart_BASEBAR = False

Dim test_hwnd As Long
Dim cText As String
Dim RECT8 As RECT


test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
Do While test_hwnd <> 0
        

        
        cText = Space$(255)
        G = GetClassName(test_hwnd, cText, 255)
        cText = StripNulls(cText)
        If cText = Test Then
            hwndtest = GetWindowRect(test_hwnd, RECT8)
            If RECT8.Top < 35 And RECT8.Bottom > 700 And RECT8.Left = 0 Then
            
                FindWindowPart_BASEBAR = test_hwnd: Exit Function

            End If
        End If
        
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop


End Function


Function FindWinPartExplorerGone(Q) As Long
'AND BRING ALL WINDOWS FORWARD
'ONLY VB SUFFERS FROM THIS

FindWinPartExplorerGone = False

Dim RECT8 As RECT



Dim ASH$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one

Dim Huge

test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
br$ = ""
Do While test_hwnd <> 0
    
    
    DoEvents
    hwnd9 = GetWindowRect(test_hwnd, RECT8)
        ASH$ = GetWindowTitle(test_hwnd)
'        If InStr(Ash$, "Double") > 0 Then Stop
        gws = GetWindowState(test_hwnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"
'
'
    If (RECT8.Top > 0 And RECT8.Left > 0) Or gws = vbMinimized Then
        ASH$ = GetWindowTitle(test_hwnd)
        If ASH$ <> "" And InStr(br$, "-- " + ASH$ + " -- ") = 0 Then
            br$ = br$ + "-- " + ASH$ + " -- "
'            On Error Resume Next
'            AppActivate Ash$, 0
'            On Error GoTo 0
            Huge = Huge + 1
            ef = SetForegroundWindow(hwnd9)
            'ef = Putfocus(hwnd9)
            ECUTE = Timer + 0.1
         
            Do
            
                DoEvents
            
                'Safety IN CASE TME RESETS BACK TO ZERO
                If Timer < ECUTE - 30 Then Exit Do
                
            Loop Until Timer > ECUTE
        
        End If
    End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

If Q = False Then MsgBox Str(Huge) + " Windows Brought To Front", vbMsgBoxSetForeground

FindWinPartExplorerGone = Huge

End Function



Private Sub CPC_404_CHILDREN()

'------------------------------------------
'----
'Get browser tab captions & close tabs-VBForums
'http://www.vbforums.com/showthread.php?691711-Get-browser-tab-captions-amp-close-tabs
'----




' lstChildHWND.List is where the handles of children are loaded
' lstKW.List is where the keywords to compare with window captions are stored
' lstHWND.List is where the handles are loaded from the EnumWindows function
' wbSave is a sub that saves relevant information to an external INI

' [EXPLANATION OF EVENTS]
' Top-level window handles are loaded into lstHWND.List
' Each handle is checked for Child windows
' If Child windows exist, get captions and check against keywords (lstKW.List)
' If no Child windows, get top-level captions and check against keywords
' If matches are found, close via SendMessage/WM_Close

'lstHWND.Clear
'EnumWindows AddressOf EnumWindowsProc, ByVal 0&
'
'Dim handCnt, kwCnt, opt As Integer
'Dim childCnt As Integer
'
'For handCnt = 0 To lstHWND.ListCount - 1
'    For kwCnt = 0 To lstKW.ListCount - 1
'    Call GetChildHandles(lstChildHWND, lstHWND.List(handCnt))
'        If lstChildHWND.ListCount > 0 Then
'            For childCnt = 0 To lstChildHWND.ListCount - 1
'                If InStr(LCase(CaptionGetFromHandle(lstChildHWND.List(childCnt))), lstKW.List(kwCnt)) > 0 Then
'                    '---BEGIN EXEMPTIONS--------
'                    If Trim(LCase(CaptionGetFromHandle(lstChildHWND.List(childCnt)))) Like Trim(LCase(Me.Caption)) Then GoTo SkipChild
'                    If Trim(LCase(CaptionGetFromHandle(lstChildHWND.List(childCnt)))) Like Trim(LCase(frmSecurity.Caption)) Then GoTo SkipChild
'                    If Trim(LCase(CaptionGetFromHandle(lstChildHWND.List(childCnt)))) Like "registry editor" Then GoTo SkipChild
'                    If Trim(LCase(CaptionGetFromHandle(lstChildHWND.List(childCnt)))) Like "windows task manager" Then GoTo SkipChild
'                    '---END OF EXEMPTIONS-------
'                    While CaptionGetFromHandle(lstChildHWND.List(childCnt)) <> ""
'                        PostMessage lstChildHWND.List(childCnt), WM_CLOSE, CLng(0), CLng(0)
'                    Wend
'                    lblBlockedCnt = Val(lblBlockedCnt) + 1
'                    Call wbSave
'                    For opt = 0 To 2
'                        If optSound(opt).Value = True Then
'                            Select Case opt
'                                Case 0: 'Silent
'                                Case 1: BeepAPI 700, 75
'                                Case 2: Call MediaWAVplay(txtSound)
'                            End Select
'                        End If
'                    Next opt
'                End If
'SkipChild:
'            Next childCnt
'        Else
'            If InStr(LCase(CaptionGetFromHandle(lstHWND.List(handCnt))), lstKW.List(kwCnt)) > 0 Then
'                '---BEGIN EXEMPTIONS--------
'                If Trim(LCase(CaptionGetFromHandle(lstHWND.List(handCnt)))) Like Trim(LCase(Me.Caption)) Then GoTo SkipTop
'                If Trim(LCase(CaptionGetFromHandle(lstHWND.List(handCnt)))) Like Trim(LCase(frmSecurity.Caption)) Then GoTo SkipTop
'                If Trim(LCase(CaptionGetFromHandle(lstHWND.List(handCnt)))) Like "registry editor" Then GoTo SkipTop
'                If Trim(LCase(CaptionGetFromHandle(lstHWND.List(handCnt)))) Like "windows task manager" Then GoTo SkipTop
'                '---END OF EXEMPTIONS-------
'                While CaptionGetFromHandle(lstHWND.List(handCnt)) <> ""
'                    PostMessage lstHWND.List(handCnt), WM_CLOSE, CLng(0), CLng(0)
'                Wend
'                lblBlockedCnt = Val(lblBlockedCnt) + 1
'                Call wbSave
'                For opt = 0 To 2
'                    If optSound(opt).Value = True Then
'                        Select Case opt
'                            Case 0: 'Silent
'                            Case 1: BeepAPI 700, 75
'                            Case 2: Call MediaWAVplay(txtSound)
'                        End Select
'                    End If
'                Next opt
'            End If
'        End If
'    Next kwCnt
'SkipTop:
'Next handCnt
'
End Sub


'Sub CPC_404_002()

'----
'Closing a firefox tab with Visual Basic (and JavaScript if needed)
'https://www.experts-exchange.com/questions/23469794/Closing-a-firefox-tab-with-Visual-Basic-and-JavaScript-if-needed.html
'----

'Private Declare Function EnumWindows Lib _
'   "user32" (ByVal lpEnumFunc As Long, lParam _
'   As Any) As Long
'
'Private Declare Function _
'   GetWindowThreadProcessId Lib "user32" (ByVal _
'   hWnd As Long, lpdwProcessId As Long) As Long
'
'Private Declare Function GetWindowLong Lib _
'   "user32" Alias "GetWindowLongA" (ByVal hWnd _
'   As Long, ByVal nIndex As Long) As Long
'
'Private Declare Function PostMessage Lib _
'   "user32" Alias "PostMessageA" (ByVal hWnd As _
'   Long, ByVal wMsg As Long, ByVal wParam As _
'   Long, lParam As Any) As Long
'
'Private Const GWL_HWNDPARENT = (-8)
'
'Private Const WM_CLOSE = &H10

'

' Initiate PostCloseMessage callbacks like this:

' Call EnumWindows(AddressOf PostCloseMessage,

' ByVal ProcessID)

'

Private Function PostCloseMessage(ByVal hwnd As Long, ByVal lParam As Long) As Long

   Static PID As Long

   ' Check if hWnd belongs to PID: lParam

   Call GetWindowThreadProcessId(hwnd, PID)

   If PID = lParam Then

      ' Belongs to target process, try close

      ' message if this doesn't have an owner. VB

      ' apps that display a msgbox from

      ' QueryUnload will lock a 9x box up hard,

      ' if we don't do this test!

      If GetWindowLong(hwnd, GWL_HWNDPARENT) = 0 Then

         Call PostMessage(hwnd, WM_CLOSE, 0&, ByVal 0&)

      End If

   End If

   ' Return true to continue enumeration.

   PostCloseMessage = True

End Function



'----
'FreeVBCode code snippet: Get hWnd of All Visible Windows on Screen.
'http://www.freevbcode.com/ShowCode.asp?ID=8408
'----

Function FindWinPart_404_REMOVE_TAB_CHROME(HUGE1, HUGE2) As Long


    'This code submitted by Sarun101.
    Dim DeskTophWnd As Long, WindowhWnd As Long
    Dim Buff As String * 255, WindowsCaption() As String
    Dim WindowsHandle() As Long
    ReDim WindowsCaption(0)
    ReDim WindowsHandle(0)
    'DeskTophWnd = GetDesktopWindow
    'WindowhWnd = GetWindow(DeskTophWnd, 5)
    WindowhWnd = FindWindowDLL(ByVal 0&, ByVal 0&)
    Do While (WindowhWnd <> 0)
        HUGE1 = HUGE1 + 1
                    
        GetWindowText WindowhWnd, Buff, 255
        'If (Trim(Buff) <> "") Then 'And (IsWindowVisible(WindowhWnd) > False) Then
            'ShowWindowAsync WindowhWnd, 0
'            ReDim Preserve WindowsCaption(UBound(WindowsCaption) + 1)
'            ReDim Preserve WindowsHandle(UBound(WindowsHandle) + 1)
'            WindowsCaption(UBound(WindowsCaption)) = Buff
'            WindowsHandle(UBound(WindowsHandle)) = WindowhWnd
            
            'HUGE1 = HUGE1 + 1
            
            
           hwnd2 = FindWindowEx(WindowhWnd, 0, "Chrome_WidgetWin_1", "404 Page Not Found | CPC - Google Chrome")
           If hwnd2 <> 0 Then Stop
           
           
           'End If
            
            
            
            '
            If InStr(Buff, "404 Page") > 0 Then Stop
            


        'End If
        WindowhWnd = GetWindow(WindowhWnd, GW_HWNDNEXT)
        'Debug.Print WindowhWnd
    Loop
    'The caption of window is in WindowsCaption()
    'The handle of window is in WindowsHandle()



Exit Function

Dim ASH$, ECUTE As Double, TIME_VAR2 As Double
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

Dim lngHwnd As Long

'Find the first window
'need this is you gona use this procedure get from CIDRun ME an another one
'Dim HUGE1, HUGE2

test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)





Do While test_hwnd <> 0
    
    DoEvents
    
        'ASH$ = GetWindowTitle(test_hwnd)
'        If InStr(ash$, "Double") > 0 Then Stop
        'gws = GetWindowState(test_hwnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"

        ASH$ = GetWindowTitle(test_hwnd)
'        If ASH$ <> "" Then Stop

        If test_hwnd = 1181660 Then Stop
        If test_hwnd = 662148 Then Stop

        cText = Space$(255)
        G = GetClassName(test_hwnd, cText, 255)
        cText = StripNulls(cText)
        'If cText = Test Then
'        Debug.Print cText
'        If InStr(ASH$, "404") > 0 Then Stop
        'If 1 = 1 Or ASH$ <> "" And InStr(ASH$, "404") > 0 Then
         
         If ASH = "Chrome Legacy Window" Then Stop
             
         If 1 = 1 Then
            On Error Resume Next
            'AppActivate Ash$
            On Error GoTo 0
            
            HUGE1 = HUGE1 + 1
                    
            'TargetHwnd = PostMessage(test_hwnd, WM_CLOSE, 0&, 0&)
 
           lngHwnd = FindWindowEx(test_hwnd, 0, vbNullString, vbNullString)
        
        
           Do Until lngHwnd = 0
                If lngHwnd = 1181660 Then Stop
                
                GWT = GetWindowTitle(lngHwnd)
                If InStr(GWT, "404 Page") > 0 Then
                    Debug.Print lngHwnd
                    Debug.Print GWT
                    
                
                    HUGE2 = HUGE2 + 1
                    'SendKeys "^(W)", False
                    'TargetHwnd = PostMessage(lngHWnd, WM_CLOSE, 0&, 0&)
                
                End If
                lngHwnd = FindWindowEx(Form1.hwnd, lngHwnd, vbNullString, vbNullString)
           
           Loop
            
            'ef = SetForegroundWindow(test_hwnd)
            'ef = Putfocus(test_hwnd)
            
            'ECUTE = Int(Now) + (Timer + 5)
            'NOW Safety IN CASE TME RESETS BACK TO ZERO
            'Do
            '    DoEvents
            '    TIME_VAR2 = Int(Now) + Timer
            '    If TIME_VAR2 > ECUTE Then
            '        Exit Do
            '    End If
            '    If GetForegroundWindow = test_hwnd Then Exit Do
            'Loop Until 1 = 2
        
'            If GetForegroundWindow = test_hwnd Then
'                HUGE2 = HUGE2 + 1
'                'SendKeys "^(W)", False
'                TargetHwnd = PostMessage(test_hwnd, WM_CLOSE, 0&, 0&)
'
'
'                Sleep 500
'            End If
        
        
        End If
        
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    
    Form1.MNU_404_CPC_PAGES.Caption = CPC_404_CPC + " -- " + Str(HUGE1) + " Checked --" + Str(HUGE2) + " Removed"

Loop

Form1.MNU_404_CPC_PAGES.Caption = CPC_404_CPC + " -- " + Str(HUGE1) + " Checked --" + Str(HUGE2) + " Removed"
               
End Function




Function FindWinPart_404_REMOVE_TAB_CHROME_002(HUGE1, HUGE2) As Long

Dim ASH$, ECUTE As Double, TIME_VAR2 As Double
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

Dim lngHwnd As Long

'Find the first window
'need this is you gona use this procedure get from CIDRun ME an another one
'Dim HUGE1, HUGE2

test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)

Do While test_hwnd <> 0
    
    DoEvents
    
        'ASH$ = GetWindowTitle(test_hwnd)
'        If InStr(ash$, "Double") > 0 Then Stop
        'gws = GetWindowState(test_hwnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"

        ASH$ = GetWindowTitle(test_hwnd)
'        If ASH$ <> "" Then Stop

        If test_hwnd = 1181660 Then Stop
        If test_hwnd = 662148 Then Stop

        cText = Space$(255)
        G = GetClassName(test_hwnd, cText, 255)
        cText = StripNulls(cText)
        'If cText = Test Then
'        Debug.Print cText
'        If InStr(ASH$, "404") > 0 Then Stop
        'If 1 = 1 Or ASH$ <> "" And InStr(ASH$, "404") > 0 Then
         
         If ASH = "Chrome Legacy Window" Then Stop
             
         If 1 = 1 Then
            On Error Resume Next
            'AppActivate Ash$
            On Error GoTo 0
            
            HUGE1 = HUGE1 + 1
                    
            'TargetHwnd = PostMessage(test_hwnd, WM_CLOSE, 0&, 0&)
 
           lngHwnd = FindWindowEx(test_hwnd, 0, vbNullString, vbNullString)
        
        
        
           Do Until lngHwnd = 0
                If lngHwnd = 1181660 Then Stop
                
                GWT = GetWindowTitle(lngHwnd)
                If InStr(GWT, "404 Page") > 0 Then
                    Debug.Print lngHwnd
                    Debug.Print GWT
                    
                
                    HUGE2 = HUGE2 + 1
                    'SendKeys "^(W)", False
                    'TargetHwnd = PostMessage(lngHWnd, WM_CLOSE, 0&, 0&)
                
                End If
                lngHwnd = FindWindowEx(Form1.hwnd, lngHwnd, vbNullString, vbNullString)
           
           Loop
            
            'ef = SetForegroundWindow(test_hwnd)
            'ef = Putfocus(test_hwnd)
            
            'ECUTE = Int(Now) + (Timer + 5)
            'NOW Safety IN CASE TME RESETS BACK TO ZERO
            'Do
            '    DoEvents
            '    TIME_VAR2 = Int(Now) + Timer
            '    If TIME_VAR2 > ECUTE Then
            '        Exit Do
            '    End If
            '    If GetForegroundWindow = test_hwnd Then Exit Do
            'Loop Until 1 = 2
        
'            If GetForegroundWindow = test_hwnd Then
'                HUGE2 = HUGE2 + 1
'                'SendKeys "^(W)", False
'                TargetHwnd = PostMessage(test_hwnd, WM_CLOSE, 0&, 0&)
'
'
'                Sleep 500
'            End If
        
        
        End If
        
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    
    Form1.MNU_404_CPC_PAGES.Caption = CPC_404_CPC + " -- " + Str(HUGE1) + " Checked --" + Str(HUGE2) + " Removed"

Loop

Form1.MNU_404_CPC_PAGES.Caption = CPC_404_CPC + " -- " + Str(HUGE1) + " Checked --" + Str(HUGE2) + " Removed"
               
End Function


Function FindWinPartFront(Q) As Long

FindWinPartFront = False

Dim RECT8 As RECT

Dim ASH$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one

Dim Huge

test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
br$ = ""
Do While test_hwnd <> 0
    
    DoEvents
    
    hwnd9 = GetWindowRect(test_hwnd, RECT8)
        ASH$ = GetWindowTitle(test_hwnd)
'        If InStr(ash$, "Double") > 0 Then Stop
        gws = GetWindowState(test_hwnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"

    
    If (RECT8.Top > 0 And RECT8.Left > 0) Or gws = vbMinimized Then
        ASH$ = GetWindowTitle(test_hwnd)
        If ASH$ <> "" And InStr(br$, "-- " + ASH$ + " -- ") = 0 Then
            br$ = br$ + "-- " + ASH$ + " -- "
            On Error Resume Next
            'AppActivate Ash$
            On Error GoTo 0
            Huge = Huge + 1
            ef = SetForegroundWindow(hwnd9)
            'ef = Putfocus(hwnd9)
            
            ECUTE = Int(Now) + (Timer + 0.1)
            'NOW Safety IN CASE TME RESETS BACK TO ZERO
            Do
                DoEvents
                TIME_VAR2 = Int(Now) + Timer
                If TIME_VAR2 > ECUTE Then
                    Exit Do
                End If
            Loop Until 1 = 2
        
        End If
    End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

If Q = False Then MsgBox Str(Huge) + " Windows Brought To Front"
FindWinPartFront = Huge

End Function


'Ambigous in ShellExecute
'Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
'    If IsWindow(lngHwnd) = 1 Then
'        GetWindowState = -1
'    If IsIconic(lngHwnd) <> 0 Then
'        GetWindowState = vbMinimized
'    ElseIf IsZoomed(lngHwnd) <> 0 Then
'        GetWindowState = vbMaximized
'    End If
'    End If
'End Function



