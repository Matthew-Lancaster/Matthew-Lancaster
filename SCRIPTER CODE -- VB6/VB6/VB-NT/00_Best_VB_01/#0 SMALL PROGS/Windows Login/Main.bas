Attribute VB_Name = "modMain"
'*************************************************************************************
'* The purpose of this app is to have a way of logging userlogon and logof on a      *
'* Win9x machine. At the scool I administer the student computers and I had problems *
'* with students installing all sorts of program on the computers via MS Powerpoint. *
'* This method omits the policy settings and was giving me a real headache.          *
'* But with this program running in the background I can get the date on programs    *
'* installed to a system, check the log and se who used the computer at the current  *
'* time and nail her/him ;-)                                                         *
'* The program is invisible to the user and it accepts commandline arguments.        *
'* Start it with /? for list of arguments.                                           *
'* You are allowed to use this code as you want.                                     *
'*                                                                                   *
'* Trond Sørensen, trond.sorensen@bi.no                                              *
'*************************************************************************************

Option Explicit

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Public Const RSP_SIMPLE_SERVICE = 1     ' For hiding app in tasklist
Public Const RSP_UNREGISTER_SERVICE = 0 ' For hiding app in tasklist
Public Const logFile = "E-Winlog.sys"   ' Logfilename constant.

' Declaring fixed lenght variables
Public Ip As String * 16                ' The ip used
Public HOST As String * 16              ' Name of host computer
Public Ident As String * 16             ' Logon ident
Public TimeNow As String * 21           ' The time now
Public TimeOn As String * 25            ' Total time between logon and logof
Public TimeLogOn As String * 21         ' The time loged on or started computer
Public TimeLogOff As String * 21        ' The time loged off or shutdown of computer

Public LogfileDir As String             ' Path to directory to make logfile
Public WinSysDir As String              ' Path to system directory
Public WindDir As String                ' Path to windows directory

Public Sub Main()
       
    ' to debug, execute next line to show main form. The app can be ended from it.
'    frmMain.Show
    
    Load frmMain
    
    WinSysDir = SysDir(True)        ' Path to Windows System Directory with leading "\"
    WindDir = WinDir(True)          ' Path to Windows Directory with leading "\"
    
    ' Checking commandline parameters
    If Left(Command$, 2) = "/?" Then
        ' Display the commandline help form.
        frmHelp.Show
        ' No need to end the app here,
        ' it ends by clicking the OK button on the help form
    ElseIf Left(Command$, 2) = "/q" Then
        ' First check if app is running. If it is, make file.
        ' If not, do not make file. If we do the app will end when it finds
        ' the file next time it starts.
        If App.PrevInstance Then
            ' Make e-winlog.dat file in system dir.
            ' The already running app will then terminate.
            Open WinSysDir & "e-winlog.dat" For Output As #2
            Print #2, "E-WinLog terminate data"
            Close #2
            End
        Else
            End
        End If
    ElseIf Left(Command$, 2) = "/i" Then
        ' Display info about app already running or not.
        If App.PrevInstance Then
            MsgBox App.ProductName & " is running", vbInformation
            End
        Else
            MsgBox App.ProductName & " is NOT running", vbInformation
            End
        End If
    ElseIf Left(Command$, 2) = "/r" Then
        ' Redirect the logfile to another location than the default system dir.
        LogfileDir = Mid(Command$, 4)
        ' Test to se if the user actualy wrote a pathname behind the argument
        If LogfileDir = "" Then
            MsgBox "No pathname given!", vbCritical
            End     ' quit
        End If
        ' Test for the existance of the pathname given.
        If (Dir(LogfileDir, vbDirectory) = "") Then
            MsgBox "The folder " & "'" & LogfileDir & "'" & " does not exist!" _
            & Chr(13) & "Terminating application!", vbCritical
            End     ' quit
        End If
        ' Test for leading backslash in the pathname given.
        ' If none found, add one.
        If Right(LogfileDir, 1) = "\" Then
            LogfileDir = LogfileDir
        Else
            LogfileDir = LogfileDir & "\"
        End If
        Call startApp
    ElseIf Left(Command$, 2) = "/l" Then
        ' Make e-winlogFileInfo.dat file in system dir
        ' The already running app will then display info about
        ' where its current logfil resides.
        ' If app is running, make file.
        If App.PrevInstance Then
            Open WinSysDir & "e-winlogFileInfo.dat" For Output As #2
            Print #2, "E-WinLog FileInfo data"
            Close #2
            End
        Else
            ' If app is not running. do not make file. If we do, next time
            ' the app starts it will display info about the logfile.
            MsgBox App.ProductName & " is NOT running", vbInformation
            End
        End If
    ElseIf Left(Command$, 2) = "/v" Then
        ' Make e-winlogViewLog.dat file in system dir
        ' The already running app will then launch notepad with
        ' the current logfil.
        ' If app is running, make file.
        If App.PrevInstance Then
            Open WinSysDir & "e-winlogViewLog.dat" For Output As #2
            Print #2, "E-WinLog Viewlog data"
            Close #2
            End
        Else
            ' App is not running.
            ' Open the logfile in Notepad from default loocation
            Dim RetVal
            RetVal = Shell(WindDir & "NOTEPAD.EXE " & WinSysDir & logFile, 1)
            End
        End If
    ElseIf Left(Command$, 2) = "" Then
        ' No arguments passed. Start the program with default setings.
        LogfileDir = WinSysDir
        
        Call startApp
    End If
End Sub

Public Function SysDir(Optional ByVal AddSlash As Boolean = False) As String

    ' Returns the path to the system directory on the local system
    Dim t As String * 255
    Dim i As Long
    i = GetSystemDirectory(t, Len(t))
    SysDir = Left(t, i)

    If (AddSlash = True) And (Right(SysDir, 1) <> "\") Then
        SysDir = SysDir & "\"
    ElseIf (AddSlash = False) And (Right(SysDir, 1) = "\") Then
        SysDir = Left(SysDir, Len(SysDir) - 1)
    End If
End Function

Public Function WinDir(Optional ByVal AddSlash As Boolean = False) As String
    
    ' Returns the path to the windows directory on the local system
    Dim t As String * 255
    Dim i As Long
    i = GetWindowsDirectory(t, Len(t))
    WinDir = Left(t, i)
    
    If (AddSlash = True) And (Right(WinDir, 1) <> "\") Then
        WinDir = WinDir & "\"
    ElseIf (AddSlash = False) And (Right(WinDir, 1) = "\") Then
        WinDir = Left(WinDir, Len(WinDir) - 1)
    End If
End Function
Public Sub startApp()

    ' This is the place the app actualy starts.
    ' Check to se if the program is already running
    If App.PrevInstance Then
        End     ' Quit
    End If
    
    ' Start the program in normal mode
    Dim test As Boolean

    ' test for existance of logfile
    test = FileExists(LogfileDir & logFile)
    If test = False Then    ' logfile does not exist
        Call MakeLogHeader
    End If
    
    ' Hide the app fron the tasklist
    Call MakeMeService
       
    Call Logon

End Sub

Public Function FileExists(ByVal PathName As String) As Boolean
    
    ' Ths function checks for the existence of a named file
    FileExists = IIf(Dir$(PathName) = "", False, True)

End Function
Public Sub MakeLogHeader()
        
    ' This sub makes a header in the log file.
    ' Will only be executed if there is no existing logfile
    
    Open LogfileDir & logFile For Append As #1
    Print #1, "E-WinLog " & "v." & App.Major & "." & App.Minor & "." & App.Revision & " log file:"
    Print #1, "-------------------------------"
    Print #1,
    Print #1, "E-WinLog is Freeware"
    Print #1, "Copyright © 2000 Trond Sørensen"
    Print #1, "All Rights Reserved"
    Print #1, "E-mail: trond.sorensen@bi.no"
    Print #1,
    Print #1, "User            Log in               Log out              Time logged on           IP used         Host computer"
    Print #1, "--------------- -------------------- -------------------- ------------------------ --------------- ----------------"
    Close #1

End Sub

Public Sub Logon()

    ' This sub writes the first part of the session line in the logfile.
    ' If the computer crashed last time, E-Winlog was terminated before
    ' this sub was executed and the app was unable to write the last part
    ' of the sessionline.
    ' So this sub must check if there is a full line at the end of the
    ' logfile. We check this by testing if the last character in the
    ' logfile is linebreak (chr(13)). If it is not we must add it first.
    
    Ident = NetUserName
    TimeNow = Date & " - " & Time
    TimeLogOn = Now
    Dim MaxSize, cPos, MyChar
    
    Open LogfileDir & logFile For Input As #1   ' Open file for input.
    MaxSize = LOF(1)        ' Get size of file in bytes.
    cPos = MaxSize - 1      ' Set cursor possition to the end of file.
    Seek #1, cPos           ' Set position to the second last character.
    MyChar = Input(1, #1)   ' Read the next character (the last).
    
    ' Now close the file. Next we are going to print to the file.
    ' To do that we must open it in append mode.
    Close #1
    
    Open LogfileDir & logFile For Append As #1
    If MyChar = Chr(13) Then
        Print #1, Ident;
        Print #1, TimeNow;
    Else
        Print #1,   ' this is the same as Chr(13)
        Print #1, Ident;
        Print #1, TimeNow;
    End If
    Close #1        ' Close file.

End Sub

Public Sub LogOf()
    
    ' This sub writes the last part of the session line in the logfile.
    
    Ip = GetIPAddress()
    HOST = ComputerName
    Ident = NetUserName
    TimeLogOff = Now
    TimeOn = FormatCount(DateDiff("s", TimeLogOn, TimeLogOff))

    ' For some reason it seems as the app not always writes the first part of the line
    ' when starting. Therefore we must check to se if it has been written.
    ' If not, add spaces first to get the last part in the right place.
    ' I think this bug is sorted out by unregistering the service before
    ' quiting the app, but I leave the code here.
    
    Dim MaxSize, cPos, MyChar
    
    Open LogfileDir & logFile For Input As #1   ' Open file for input.
    MaxSize = LOF(1)        ' Get size of file in bytes.
    cPos = MaxSize - 4      ' Set cursor possition to the ":" at last line.
    Seek #1, cPos           ' Set position to the second last character.
    MyChar = Input(1, #1)   ' Read the next character (the last).
    
    ' Now close the file. Next we are going to print to the file.
    ' To do that we must open it in append mode.
    Close #1

    Open LogfileDir & logFile For Append As #1
    If MyChar = ":" Then    ' Then the first part of line already is there
        Print #1, TimeNow;
        Print #1, TimeOn;
        Print #1, Ip;
        Print #1, HOST      ' skip the semicolon here. Makes the print statement print linebrake.
    Else                    ' Then the first part of line is NOT there
        Print #1, Ident;
        Print #1, "Not registered       ";
        Print #1, TimeNow;
        Print #1, TimeOn;
        Print #1, Ip;
        Print #1, HOST      ' skip the semicolon here. Makes the print statement print linebrake
    End If
    Close #1    ' Close file.

End Sub

Function ComputerName() As String
    
    ' Returns the name of the computer.
    ' Call like this: Host = ComputerName
    
    Dim buffer As String * 512, length As Long
    length = Len(buffer)
    If GetComputerName(buffer, length) Then
        ' this API returns non-zero if successful,
        ' and modifies the length argument
        ComputerName = Left$(buffer, length)
    End If
End Function

Function NetUserName() As String
    ' This function returns the current username
    
    Dim Rtn As Long 'declare the needed variables
    Dim UserName As String * 255
    
    Rtn = WNetGetUser("", UserName, 255)
       
    If Rtn = 0 Then
       NetUserName = Left(UserName, InStr(UserName, Chr$(0)) - 1)
    Else
       NetUserName = "No User Name"
    End If
       
End Function

Function FormatCount(Count As Long) As String

    ' This function formats the data (seconds) the DateDiff returns
    ' to days, hours, minutes and seconds.
    
    Dim Days As Integer, Hours As Long, Minutes As Long, Seconds As Long
    
    Days = Count \ (24& * 3600&)
    If Days > 0 Then Count = Count - (24& * 3600& * Days)
    Hours = Count \ 3600&
    If Hours > 0 Then Count = Count - (3600& * Hours)
    Minutes = Count \ 60
    Seconds = Count Mod 60

    FormatCount = Days & " d, " & Hours & " h, " & Minutes & " m, " & Seconds & " s "
End Function

Public Sub MakeMeService()
    ' Remove your program from the Ctrl+Alt+Delete list (tasklist), make it a service
    ' Don't forget to unregister your application as a service before it closes to
    ' free up system resources by calling UnMakeMeService.

    Dim pid As Long
    Dim regserv As Long
    
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub

Public Sub UnMakeMeService()
    ' Restore your application to the Ctrl+Alt+Delete list (takslist).
    Dim pid As Long
    Dim regserv As Long

    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
End Sub
