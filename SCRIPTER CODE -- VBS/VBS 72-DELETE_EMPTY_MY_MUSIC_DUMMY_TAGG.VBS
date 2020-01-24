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
DIM strScriptName
DIM PID_Script

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")


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

' -------------------------------------------------------------------
' GOOD CODE TO PASTE ON VB6
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

Dim objFolder
DIM objFolderERR

DIM T
DIM I
DIM DELETE_COUNTER
DELETE_COUNTER=0

IF FSO.FOLDEREXISTS(STRPATH) = TRUE THEN 
		CALL DeleteEmptyFolders (FSO.GetFolder(STRPATH))
	ELSE
		WScript.Echo STRPATH & VBCRLF & "__ NOT A VALID FOLDER GIVEN __ EXIT END __"
END IF

MSGBOX "DELETE_COUNTER"+VBCRLF+CSTR(DELETE_COUNTER)


Wscript.QUIT

Sub DeleteEmptyFolders(folder)
    Dim subfolder
	Dim file
	DIM DONE_H_1
	DIM DONE_H_2
	DIM EXTENSION_T
	
    For Each subfolder in folder.SubFolders
		DeleteEmptyFolders subfolder
    Next

	DONE_H_1=FALSE
	FOR EACH file in folder.Files
		IF MID(FILE.NAME,1,3)="0#0" THEN 
			DONE_H_1=TRUE
		END IF 
	NEXT
	
	DONE_H_2=FALSE
	FOR EACH file in folder.Files
		IF DONE_H_1=TRUE AND MID(FILE.NAME,1,3)<>"0#0" THEN 
			EXTENSION_T = UCASE(FSO.getextensionname(file.path))+" "
			EXTENSION_V = "MP3 MOD MID MIDI AMR SRF WAV MP4 MID M3U M3U8 WMV FLV WMA AVI M4A OGG RMI KAR MMF MOD SFV PAR2 QP2 "
			IF INSTR(EXTENSION_V,EXTENSION_T)>0 THEN 
				DONE_H_2=TRUE
			END IF
		END IF
	NEXT
	
	' FILE TYPE IN FOLDER MUSIC
	' -------------------------
	' 1 -- .TXT
	' 2 -- .MP3
	' 3 -- .RAR
	' 4 -- .SRF
	' 5 -- .LNK
	' 6 -- .JPG
	' 7 -- .AMR
	' 8 -- .WAV
	' 9 -- .MP4
	' 10 -- .BMP
	' 11 -- .URL
	' 12 -- .NVC
	' 13 -- .XLS
	' 14 -- .MID
	' 15 -- .M3U
	' 16 -- .CRDOWNLOAD
	' 17 -- .ZIP
	' 18 -- .BAT
	' 19 -- .HTM
	' 20 -- .PNG
	' 21 -- .M3U8
	' 22 -- .HTML
	' 23 -- .GIF
	' 24 -- .08 SHOW ON DEVIANTAUDIO - DEVIANT AUDIO
	' 25 -- . LENGTH - 32
	' 26 -- .07 - STYLE
	' 27 -- . 11
	' 28 -- .LOG
	' 29 -- ._GS
	' 30 -- .GSL
	' 31 -- .WMV
	' 32 -- .FLV
	' 33 -- .WMA
	' 34 -- .DLL
	' 35 -- .EXE
	' 36 -- .MAP
	' 37 -- .TAB
	' 38 -- .INI
	' 39 -- .AVI
	' 40 -- .DOC
	' 41 -- .WHTT
	' 42 -- .LOCK
	' 43 -- .LST
	' 44 -- .REF
	' 45 -- .HTACCESS
	' 46 -- .TMP
	' 47 -- .M4A
	' 48 -- .BZ2
	' 49 -- .GZ
	' 50 -- .ICO
	' 51 -- .OGG
	' 52 -- .LZX
	' 53 -- .INF
	' 54 -- .VQF
	' 55 -- .LHA
	' 56 -- .CF
	' 57 -- .PIC
	' 58 -- .PL
	' 59 -- .Q1
	' 60 -- .RMI
	' 61 -- .KAR
	' 62 -- .MMF
	' 63 -- .MOD
	' 64 -- .PAR
	' 65 -- .NFO
	' 66 -- .MHT
	' 67 -- .RTF
	' 68 -- .BAS
	' 69 -- .VBP
	' 70 -- .PDF
	' 71 -- .SFV
	' 72 -- .PAR2
	' 73 -- .QP2
	' 74 -- .PK
	' 75 -- .DAT
	' 76 -- .AIFF
	' 77 -- .FRM
	' 78 -- .CSS
	' 79 -- .JS


	FLAG_ONE=FALSE
	IF DONE_H_1=TRUE THEN 
	IF DONE_H_2=FALSE THEN 
		FOR EACH file in folder.Files
			IF MID(FILE.NAME,1,3)="0#0" AND FLAG_ONE=FALSE THEN 
				FLAG_ONE=TRUE
				T=T+FOLDER+VBCRLF
				I=I+1
			END IF
		NEXT
	END IF
	END IF
	
	IF DONE_H_1=TRUE THEN 
	IF DONE_H_2=FALSE THEN 
		FOR EACH file in folder.Files
			IF MID(FILE.NAME,1,3)="0#0"  THEN 

				TTI=FILE.PATH
				FSO.DELETEFILE FILE.PATH
				DELETE_COUNTER=DELETE_COUNTER+1
				' MSGBOX TTI

			END IF
		NEXT
	END IF
	END IF
	
	IF I>20 THEN 
		WSHShell.Run """"+"EXPLORER.EXE"+"""" + " /SELECT, " + ""+FOLDER.PATH+"" , ShowWindow, WaitUntilFinished
		MSGBOX T
		I=0
		T=""
	END IF 
	
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
		' ON ERROR RESUME NEXT
		' ERR.CLEAR
		' FSO.DeleteFolder folder.Path, TRUE
	End If
End Sub

