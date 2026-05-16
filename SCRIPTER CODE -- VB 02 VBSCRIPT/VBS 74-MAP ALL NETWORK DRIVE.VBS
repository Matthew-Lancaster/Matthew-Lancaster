' Option Explicit

Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

' -------------------------------------------------------------------
' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF ARE
' SOURCE GIVEN
' -------------------------------------------------------------------
' vbscript - how to check vbs script in windows is running or not? - Stack Overflow
' https://stackoverflow.com/questions/7849699/how-to-check-vbs-script-in-windows-is-running-or-not
' -------------------------------------------------------------------
' [ Wednesday 13:08:00 Pm_20 March 2019 ]
' -------------------------------------------------------------------

' -------------------------------------------------------------------

DIM GETUSERNAME
DIM GETCOMPUTERNAME
DIM COMPUTERNAME
DIM WSHSHELL
SET WSHSHELL = CreateObject("WScript.Shell")
GETUSERNAME = CreateObject("WScript.Network").USERNAME
GETCOMPUTERNAME = CreateObject("WScript.Network").COMPUTERNAME
COMPUTERNAME = CreateObject("WScript.Network").COMPUTERNAME
' ------------------------------------------------------------------
' ANOTHER EXAMPLE
' ------------------------------------------------------------------
' COMPUTERNAME = WSHSHELL.ExpandEnvironmentStrings("%computername%")
' ------------------------------------------------------------------


strComputer = "."
strComputer = ComputerName


Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutFile = objFSO.CreateTextFile(".\MappedDrives.txt")

Set colDrives = objWMIService.ExecQuery _
    ("Select * From Win32_LogicalDisk Where DriveType = 4")

For Each objDrive in colDrives
    objOutFile.WriteLine(objDrive.DeviceID & " (" & _
      objDrive.ProviderName & ")")
Next

objOutFile.Close



Wscript.QUIT

