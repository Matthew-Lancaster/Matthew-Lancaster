Attribute VB_Name = "Module1"
' ---------------------------------------------------------------------
' MAKE SEPARATE ROUTINE AS
' SHELLEXECUTEEX_OPEN_PROPERTIE
' ---------------------------------------------------------------------
' WHEN EXECUTE COMMAND
' IF CLOSE APP THAT LAUNCH THEM
' THEY CLOSE ALSO
' ALSO ON THE INTERNET
' MANY TALKER
' AND THE REQUEST WAY
' IS USE CREATE PROCESS WHICH RETURN A LOT MORE VARIABLE
' BUT THAT NOT FEATURE THE ITEM REQUIREING
' SO BATTLE THOUGHT THAT MIND FIELD
' THE APP WAS CLOSE AT TOP TALK AND CLOSE THE REQUEST APP THAT RUN
' THAT IDEA FIND THE HWND FROM THE PID
' A BIT RUSTY ABOUT THAT ONE
' AND LOT WORKER -- PROCESS NIGHT MARE
' EVEN THEN WAS WAIT UNTIL PRCOCESS DONE
' I GOT FEW SHOP WANT THIS ONE
' RESOLVER BY CLEVER CHANGE THE HWND AND ATTACH TO ANOTHER APP THE DESKTOP WINDOW
' AFTER THAT RESULT WANTER
' SetParent SEI.hWnd, lHnd
' THIS HERE CODE
' WILL OPEN UP A PROPERTIES DIALOG
' SO ABLE ADJUST THE PERMISSION NOT WITH EXPLORER
' AND WILL HANG A LITTLE BIT AFTER A WHILE
' IS RUOTINE TO LOAD THEM ALL
' OR CALLER PROGRAM AND USE OTHER THING
' LOAD ONE CALLER APP LOAD THE SET ALL THEM
' ---------------------------------------------------------------------

' NEWER ADDITION PUBLISHER
' ---------------------------------------------------------------------

' TIME TO FIND HERE BEFORE WORKER
' --------------------------------------------------------------------
' 0001
' Fri 26-Jun-2020 10:44:17
' Fri 26-Jun-2020 12:00:00 -- 1 HOUR 15 MINUTE
' --------------------------------------------------------------------
' SHELLEXECUTE TO SPECIFY THE PROPERTIES VERB VB6 - GOOGLE SEARCH
' https://www.google.co.uk/search?safe=active&sxsrf=ALeKk02PtyW3kzi1wOQHAHk707nxgZsszg%3A1593167672498&ei=OM_1Xp-DHtujgAbcyamgAg&q=ShellExecute+to+specify+the+properties+verb+VB6&oq=ShellExecute+to+specify+the+properties+verb+VB6&gs_lcp=CgZwc3ktYWIQAzoHCAAQRxCwAzoHCCMQrgIQJzoFCCEQoAFQjdkOWOfeDmCF4g5oAXAAeACAAZ0BiAGvA5IBAzIuMpgBAKABAaoBB2d3cy13aXo&sclient=psy-ab&ved=0ahUKEwif_76JpJ_qAhXbEcAKHdxkCiQQ4dUDCAw&uact=5
' ---------------------------------------------------------------------
' 0001
' Fri 26-Jun-2020 10:44:17
' Fri 26-Jun-2020 12:28:10 -- 1 HOUR 43 MINUTE
' ---------------------------------------------------------------------
' TOTAL -- 1 + 2 -- 43+42+47+27 -- 159 -- 2 HOUR 39 + 1 + 2 -- 5 HOUR 39
' ---------------------------------------------------------------------
' Fri 26-Jun-2020 10:44:17 -- START TIME
' Sat 27-Jun-2020 06:22:51 -- END -- 19 HOUR 38 MINUTE
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' 0002
' Fri 26-Jun-2020 14:34:56 --
' Fri 26-Jun-2020 15:17:48 -- 42 MINUTE
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' 0003
' Fri 26-Jun-2020 14:23:31
' Fri 26-Jun-2020 17:11:08 -- 2 HOUR 47 MINUTE
' HEAD EXPLODE
' GO WORK ANOTHER ITEM UPSTAIR BEDROOM -- GOODSYNC STUFF WORK
' SUCCESSFUL RESULT
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' 0004
' Sat 27-Jun-2020 05:54:59
' Sat 27-Jun-2020 06:22:51 -- 27 MINUTE
' END TIME FIND RESULT TO SETHWND HANDLE TO ANOTHER APP
' THAT WAY MAKE WHEN QUIT PROGRAM HERE NOT PULL THE OTHER APP TO ALSO
' ---------------------------------------------------------------------
' TOTAL -- 1 + 2 -- 43+42+47+27 -- 159 -- 2 HOUR 39 + 1 + 2 -- 5 HOUR 39
' Fri 26-Jun-2020 10:44:17 -- START TIME
' Sat 27-Jun-2020 06:22:51 -- END -- 19 HOUR 38 MINUTE
' ---------------------------------------------------------------------


' ---------------------------------------------------------------------
' TWO SESSION -- 1 HOUR 43 MINUTE __+__ 42 MINUTE
' FIND ERROR CODE NOT CLOSE APP -- AND DESTROY CALL APP WINDOW
' WHEN CALL BY APP CLOSE OTHER DO
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' Fri 26-Jun-2020 10:44:17 -- TOTAL SMALL WITH GAP BETWEEN
' Fri 26-Jun-2020 15:17:48 -- 4 HOUR 33 MINUTE
' ---------------------------------------------------------------------
' --------------------------------------------------------------------
' TIME TO FIND HERE BEFORE WORKER
' --------------------------------------------------------------------
' Fri 26-Jun-2020 10:44:17
' Fri 26-Jun-2020 12:00:00 -- 1 HOUR 15 MINUTE
' --------------------------------------------------------------------
' ShellExecute to specify the properties verb VB6 - Google Search
' https://www.google.co.uk/search?safe=active&sxsrf=ALeKk02PtyW3kzi1wOQHAHk707nxgZsszg%3A1593167672498&ei=OM_1Xp-DHtujgAbcyamgAg&q=ShellExecute+to+specify+the+properties+verb+VB6&oq=ShellExecute+to+specify+the+properties+verb+VB6&gs_lcp=CgZwc3ktYWIQAzoHCAAQRxCwAzoHCCMQrgIQJzoFCCEQoAFQjdkOWOfeDmCF4g5oAXAAeACAAZ0BiAGvA5IBAzIuMpgBAKABAaoBB2d3cy13aXo&sclient=psy-ab&ved=0ahUKEwif_76JpJ_qAhXbEcAKHdxkCiQQ4dUDCAw&uact=5
' --------
' ShellExecute and "properties" verb
' http://forums.codeguru.com/showthread.php?2125-ShellExecute-and-properties-verb
' --------------------------------------------------------------------
Dim COMMAND_LINE
Dim FSO

Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H28
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFF
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_TIMEOUT = &H102
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Const GWL_STYLE = -16

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' ---- OPTIONAL FIELDS
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess    As Long
  hThread     As Long
  dwProcessId As Long
  dwThreadId  As Long
End Type

Private Type STARTUPINFO
  cb              As Long
  lpReserved      As String
  lpDesktop       As String
  lpTitle         As String
  dwX             As Long
  dwY             As Long
  dwXSize         As Long
  dwYSize         As Long
  dwXCountChars   As Long
  dwYCountChars   As Long
  dwFillAttribute As Long
  dwFlags         As Long
  wShowWindow     As Integer
  cbReserved2     As Integer
  lpReserved2     As Long
  hStdInput       As Long
  hStdOutput      As Long
  hStdError       As Long
End Type

Private Enum Priorities
  p_RealTime = &H100
  p_Hight = &H80
  p_Normal = &H20
  p_Idle = &H40
End Enum

Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, hModule As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Sub Main()
    COMMAND_LINE = Command$
    ' COMMAND_LINE = "E"
    If Trim(COMMAND_LINE) = "" Then End
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Call SHELLEXECUTEEX_OPEN_PROPERTY
End Sub

Sub SHELLEXECUTEEX_OPEN_PROPERTY()
    Dim FileName
    Dim PROC As PROCESS_INFORMATION
    Dim SEI As SHELLEXECUTEINFO
    
    On Error Resume Next
    DRIVE_LETTER = Mid(COMMAND_LINE, 1, 1)
    Err.Clear
    Set z = FSO.GetDrive(FSO.GetDriveName(FSO.GetAbsolutePathName(DRIVE_LETTER + ":")))
    Select Case z.DriveType
        Case 0: T = "Unknown"
        Case 1: T = "Removable"
        Case 2: T = "Fixed"
        Case 3: T = "Network"
        Case 4: T = "CD-ROM"
        Case 5: T = "RAM Disk"
    End Select
    
    If Err.Number = 0 Then
        TT = z.DriveLetter + ":" + "- Type - " + T + " --- Serial - " + Hex$(z.SerialNumber) + " - Vol.Name - " + z.VolumeName + vbCrLf
        TG = TG + 1
        Y1 = Y1 + TT
        FileName = z.DriveLetter + ":\" ' + z.VolumeName
        
        
        lHnd = GetDesktopWindow()
        With SEI
            ' SIZE OF THE STRUCTURE
            .cbSize = Len(SEI)
            ' USE THE OPTIONAL HPROCESS ELEMENT OF THE STRUCTURE.
            .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST
            ' HANDLE TO THE WINDOW CALLING THIS FUNCTION.
            .hWnd = lHnd
            ' THE ACTION TO PERFORM: OPEN THE FILE.
            .lpVerb = "open"
            .lpVerb = "properties"
            ' THE FILE TO OPEN.
            .lpFile = "woteveryourfilewas.msg"
            .lpFile = FileName
            ' NO ADDITIONAL PARAMETERS ARE NEEDED HERE.
            .lpParameters = ""
            ' THE DEFAULT DIRECTORY -- NOT REALLY NECESSARY IN THIS CASE.
            ' .LPDIRECTORY = "C:\"
            ' SIMPLY DISPLAY THE WINDOW.
            .nShow = SW_SHOWNORMAL
            ' THE OTHER ELEMENTS OF THE STRUCTURE ARE EITHER NOT USED
            ' OR WILL BE SET WHEN THE FUNCTION RETURNS.
        End With
        
        ' OPEN THE FILE USING ITS ASSOCIATED PROGRAM.
        lRVal = ShellExecuteEx(SEI)
        If lRVal = 0 And 123 = 321 Then
            'THERE WAS AN ERROR
            Select Case SEI.hInstApp
                Case SE_ERR_FNF
                sMsg = "File not found"
                Case SE_ERR_PNF
                sMsg = "Path not found"
                Case SE_ERR_ACCESSDENIED
                sMsg = "Access denied"
                Case SE_ERR_OOM
                sMsg = "Out of memory"
                Case SE_ERR_DLLNOTFOUND
                sMsg = "DLL not found"
                Case SE_ERR_SHARE
                sMsg = "A sharing violation occurred"
                Case SE_ERR_ASSOCINCOMPLETE
                sMsg = "Incomplete or invalid file association"
                Case SE_ERR_DDETIMEOUT
                sMsg = "DDE Time out"
                Case SE_ERR_DDEFAIL
                sMsg = "DDE transaction failed"
                Case SE_ERR_DDEBUSY
                sMsg = "DDE busy"
                Case SE_ERR_NOASSOC
                sMsg = "No association for file extension"
                Case ERROR_BAD_FORMAT
                sMsg = "Invalid EXE file or error in EXE image"
                Case Else
                sMsg = "Unknown error"
            End Select
            MsgBox sMsg
        End If
        
        ' ------------------------------------------------
        ' A WORK AROUND METHOD
        ' REALLY I WOULD GET A HWND TO PID
        ' PROBLEM TO SOLVE
        ' IS WHEN CALLING PROGRAM QUIT SO DO REQUEST TO RUN ITEM
        ' MY WORK AROUND
        ' IS TRANSFER THE HWND TO ANOTHER APP LIKE DESKTOP
        ' AND WORK EASY
        ' THE SEI ABLE SET HWND
        ' AND EVEN THOUGH SET IT TO DESKTOP
        ' IT NOT REALLY CONNECTED THAT WAY
        ' OBVIOUS DESKTOP IS ONE OF IT OWN
        ' AND SEI RETURN IT OWN ONE ONCE MADE
        ' NOT TO WORRY
        ' HERE IS SOLUTION
        ' ------------------------------------------------
        SetParent SEI.hWnd, lHnd
    
        ' SET PARENT SOLUTION CAME FROM
        ' ------------------------------------------------
        ' How to get Process ID from hwnd?-VBForums
        ' http://www.vbforums.com/showthread.php?765305-How-to-get-Process-ID-from-hwnd
        ' ------------------------------------------------
    
    End If
End Sub


Private Sub MNU_OPEN_PROPERTIE_FOR_ALL_DRIVE_Click()
    ' Me.WindowState = vbMinimized
    Dim TT, TG, Y1 ' ---- NOT REQUIRE INFO OTHER EXAMPLE
    Dim FileName
    Dim s As SHELLEXECUTEINFO
    On Error Resume Next
    For R = 3 To 25
        Err.Clear
        Set z = FSO.GetDrive(FSO.GetDriveName(FSO.GetAbsolutePathName(Chr$(R + 64) + ":")))
        Select Case z.DriveType
            Case 0: T = "Unknown"
            Case 1: T = "Removable"
            Case 2: T = "Fixed"
            Case 3: T = "Network"
            Case 4: T = "CD-ROM"
            Case 5: T = "RAM Disk"
        End Select
        If Err.Number = 0 Then
            TT = z.DriveLetter + ":" + "- Type - " + T + " --- Serial - " + Hex$(z.SerialNumber) + " - Vol.Name - " + z.VolumeName + vbCrLf
            TG = TG + 1
            Y1 = Y1 + TT
            FileName = z.DriveLetter + ":\" ' + z.VolumeName
            s.cbSize = LenB(s)
            s.lpFile = FileName
            s.nShow = SW_SHOW
            s.fMask = SEE_MASK_INVOKEIDLIST
            s.lpVerb = "properties"
            ShellExecuteEx s
            
            lHnd = GetDesktopWindow()
            SetParent SEI.hWnd, lHnd
        End If
    Next
    ' Unload Me
End Sub

