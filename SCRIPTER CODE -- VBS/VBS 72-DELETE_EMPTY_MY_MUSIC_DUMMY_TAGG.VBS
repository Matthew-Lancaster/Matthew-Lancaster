Option Explicit

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
DIM CUR_USER
DIM ComputerName
DIM objShell
DIM objWMIService
DIM colProcesses
DIM i1
DIM i2
DIM objProcess
DIM VBV
VBV=VBCRLF


CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

' -------------------------------------------------------------------
' 01 OF 02 
' SOME CONFLICT WAR IN CODE STYLE 
' SOME PEOPLE TALK THIS HAS ERROR 
' THIS ONE IS CORRECT 
' -------------------------------------------------------------------
On Error Resume Next
Dim iMyPID : iMyPID = GetObject("winmgmts:root\cimv2").Get("Win32_Process.Handle='" & CreateObject("WScript.Shell").Exec("mshta.exe").ProcessID & "'").ParentProcessId
On Error Goto 0

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")
' -------------------------------------------------------------------


Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

i1 = 0 ' ANY PROGRAM WSCRIPT 
i2 = 0 ' MY  PROGRAM WSCRIPT 
FOR EACH objProcess IN colProcesses
    IF NOT (IsNull(objProcess.CommandLine )) THEN
		strScriptName = Trim(Right(objProcess.CommandLine,Len(objProcess.CommandLine) - InstrRev(objProcess.CommandLine,"\")))
        strScriptName = Left(strScriptName, Len(strScriptName) - 1)
		PID_Script = objProcess.ProcessID
		
		IF strScriptName = Wscript.ScriptName THEN
			i2=i2+1
		END IF
		IF strScriptName <> Wscript.ScriptName THEN
			i1=i1+1
		END IF
		
		IF i1 > 0 THEN ' ANY PROGRAM WSCRIPT 
			' Wscript.Quit
		END IF
		IF i2 > 0 THEN ' MY  PROGRAM WSCRIPT 
			' -------------------------------------------------------
			' KILL PROCESS NUMBER OF ALL OTHER INSTANCE 
			' HERE SCRIPT WHEN START
			' -------------------------------------------------------
			' IF lngMyProcessId<>iMyPID THEN 
				' I= "ERROR MISMATCH TRUE MYPID RESULT "
				' MSGBOX I +VBV+ CSTR(iMyPID) +VBV+ CSTR(lngMyProcessId)
			' -------------------------------------------------------
			IF iMyPID<>PID_Script THEN 
				objProcess.Terminate() 
				objShell.Run "TASKKILL /F /IM """+CSTR(PID_Script)+"""", DONTShowWindow, WaitUntilFinished
				' MSGBOX CSTR(iMyPID) +VBV+ strScriptName +VBV+ CSTR(PID_Script)  +VBV+ Wscript.ScriptName
			END IF
		END IF
    END IF
NEXT



DIM FILE_NAME , PARAM
Dim STRPATH
DIM ERR_COUNTER
DIM TEST_COUNTER
DIM EXECUTE_LANUCHED
DIM ERROR_WITH_SOME

DIM READY_01
DIM READY_02

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")


STRPATH="D:\0 00 MUSIC ---"

'STRPATH=""
'STRPATH=""
'STRPATH=""
'STRPATH=""

Dim objFolder
DIM objFolderERR

IF LEN(STRPATH) = 2 AND MID(STRPATH,2,1) = ":" THEN STRPATH = STRPATH + "\"

IF FSO.FOLDEREXISTS(STRPATH) = TRUE THEN 
		WScript.Echo STRPATH & VBCRLF & "TO GO _ KILL PROCESS IF NOT SURE ____"
		CALL DeleteEmptyFolders (FSO.GetFolder(STRPATH))
		EXECUTE_LANUCHED = TRUE
	ELSE
		WScript.Echo STRPATH & VBCRLF & "__ NOT A VALID FOLDER GIVEN __ EXIT END __"
END IF



Wscript.QUIT

Sub DeleteEmptyFolders(folder)
    Dim subfolder
	ON ERROR RESUME NEXT
	ERR.CLEAR
    For Each subfolder in folder.SubFolders
		DeleteEmptyFolders subfolder
    Next
	
	DONE_HERE_2=FALSE
	FOR EACH file in subfolder.Files
		DONE_HERE=FALSE
		IF MID(FILE.PATH,1,3)="0#0" THEN DONE_HERE=TRUE
		IF DONE_HERE=FALSE THEN 
			EXTENSION_T = UCASE(FSO.getextensionname(file.path))
			MSGBOX FILE.PATH
			IF EXTENSION_T<>"MP3" THEN 
				MSGBOX FILE.PATH
			END IF
		END IF
	NEXT

	
	' ---------------------------------------------------------------
	' MORE ERROR SAFE METHOD
	' REMEMBER DECLARE ON AS OPTION EXPLICIT 
	' AND ERROR RESULT FROM DECLARE ERROR MAKE NOT PROPER TRAPP IT
	' ---------------------------------------------------------------
	READY_01=-1
	READY_02=-1
	ERR.CLEAR
	READY_01=folder.SubFolders.Count
	READY_02=folder.Files.Count
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
    If READY_01 = 0 And READY_02 = 0 AND ERR.NUMBER=0 Then
		ON ERROR GOTO 0
		TEST_COUNTER=TEST_COUNTER+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		' FSO.DeleteFolder folder.Path, TRUE
	End If
End Sub

