Attribute VB_Name = "Module1"

Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function




Sub Main()



'     -r        Dump log from least recent to most recent.
'no

'     -s        Records are listed on one line each with delimited
'     -x        Dump extended data.
'     -a        Dump records timestamped after specified date.
'     -b        Dump records timestamped before specified date.

On Error Resume Next
Open App.Path + "\DateInfo.txt" For Input As #1
Line Input #1, file3
Close #1
On Error GoTo 0


et = Now
file1 = Format$(et, "YYYY-MM-DD HH-MM-SS DDD")
file2 = "-a " + Format$(et - 1, "MM/DD/YYYY") ' HH:MM:SS")
Open App.Path + "\DateInfo.txt" For Output As #1
Print #1, file2
Close #1

'file3 = "-a 07/15/2009"
'file3 = ""




On Error Resume Next
MkDir "D:\00 Loggers Text\Event Logg--" + GetComputerName

tt1$ = """C:\Program Files\SysInternals\PsTools\psloglist.exe"" -x " + file3 + " System >""D:\00 Loggers Text\Event Logg--" + GetComputerName + "\Event Logg - " + file1 + " - 01 System.txt"""
tt2$ = """C:\Program Files\SysInternals\PsTools\psloglist.exe"" -x " + file3 + " Application >""D:\00 Loggers Text\Event Logg--" + GetComputerName + "\Event Logg - " + file1 + " - 02 Application.txt"""
tt3$ = """C:\Program Files\SysInternals\PsTools\psloglist.exe"" -x " + file3 + " Security >""D:\00 Loggers Text\Event Logg--" + GetComputerName + "\Event Logg - " + file1 + " - 03 Security.txt"""
tt4$ = """C:\Program Files\SysInternals\PsTools\psloglist.exe"" -x " + file3 + " ""Internet Explorer"" >""D:\00 Loggers Text--" + GetComputerName + "\Event Logg\Event Logg - " + file1 + " - 04 Internet Explorer.txt"""

Open App.Path + "\EventBatchEXE.bat" For Output As #1
    Print #1, tt1$
    Print #1, tt2$
    Print #1, tt3$
    Print #1, tt4$
Close #1

Shell App.Path + "\EventBatchEXE.bat", vbHide


End Sub


Sub Commands()

'PsLoglist v2.64 - local and remote event log viewer
'Copyright (C) 2000-2006 Mark Russinovich
'Sysinternals - www.sysinternals.com
'
'PsLogList dumps event logs on a local or remote NT system.
'
'Usage: psloglist [\\computer[,computer2[,...] | @file] [-u username [-p password]]] [-s [-t delimiter]] [-m #|-n #|-d #|-h #|-w][-c][-x][-r][-a 'mm/dd/yy][-b mm/dd/yy] [-f filter] [-i ID,[ID,...]] | -e ID,[ID,...]] [-o event source[,event source[,...]]] [-q event source[,event source[,...]]] '[[-g|-l] event log file] <event log>
'     @file     Psloglist will execute the command on each of the computers
'               listed in the file.
'     -a        Dump records timestamped after specified date.
'     -b        Dump records timestamped before specified date.
'     -c        Clear event log after displaying.
'     -d        Only display records from previous n days.
'     -e        Exclude events with the specified ID or IDs (up to 10).
'     -f        Filter event types, using starting letter
'               (e.g. "-f we" to filter warnings and errors).
'     -g        Export an event log as an evt file. This can only be used
'               with the -c switch (clear log).
'     -h        Only display records from previous n hours.
'     -i        Show only events with the specified ID or IDs (up to 10).
'     -l        Dump the contents of the specified saved event log file.
'     -m        Only display records from previous n minutes.
'     -n        Only display n most recent records.
'     -o        Show only records from the specified event source or sources
'               (e.g. "-o cdrom").
'     -p        Specifies password for user name.
'     -q        Omit records from the specified event source or sources
'               (e.g. "-q cdrom").
'     -r        Dump log from least recent to most recent.
'     -s        Records are listed on one line each with delimited
'               fields, which is convenient for string searches.
'     -t        The default delimiter for the -s option is a comma,
'               but can be overriden with the specified character. Use "\t"
'               to specify tab.
'     -u        Specifies optional user name for login to
'               remote computer.
'     -w        Wait for new events, dumping them as they generate (local system
'               only.)
'     -x        Dump extended data.
'     eventlog  Specifies event log to dump. Default is system. If the
'               -l switch is present then the event log name specifies
'               how to interpret the event log file.

End Sub
