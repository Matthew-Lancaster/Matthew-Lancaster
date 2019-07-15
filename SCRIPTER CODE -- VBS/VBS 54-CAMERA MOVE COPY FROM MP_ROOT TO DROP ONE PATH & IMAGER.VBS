'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 52-INSYNC HIKVISION COPY DUPLICATE FOLDER.VBS
'# __ 
'# __ VBS 52-INSYNC HIKVISION COPY DUPLICATE FOLDER.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Tue 11-Jun-2019 13:25:39 ]
'# END FINAL TIME [ Tue 11-Jun-2019 15:14:00 ]
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' -------------------------------------------------------------------
' 2 HOUR AND 49 MINUTE TO COMPLETE ANOTHER PROJECT IN
' -------------------------------------------------------------------
' NOW INCLUDE 
' WILL COPY SOURCE TO TARGET BY FOLDER AND DELETE AFTER 
' UNABLE DO MOVE EFFECTIVE-AH IN VBS
' AND INCORPORATE CODE FROM BEFORE THAT DELETE EMPTY
' -------------------------------------------------------------------
'# START     TIME [ Tue 11-Jun-2019 13:25:39 ]
'# END FINAL TIME [ Tue 11-Jun-2019 15:14:00 ]
'#-------------------------------------------------------------------

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
' DOUBLE CODE VERSION SECOND HALF OF CODE TAKEN FROM ANOTHER PROJECT
' -------------------------------------------------------------------
' VBS 13-COPY MOVE SHELL FOLDING.VBS
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' __ CODE BY MATTHEW LANCASTER
' __ DATE BEGINNER 5 AM -- 5 HOURS
' __ DATE BEGINNER AND ENDER ----------------------------------------
' [ Saturday 10:18:00 Am_30 December 2017 ] The UK
' [ Saturday 10:18:00 Am_30 December 2017 ] GMT UTC+0 WET
' __ DATE BEGINNER AND ENDER ----------------------------------------
' -------------------------------------------------------------------
' THIS WILL COPY SOURCE TO TARGET BY FOLDER AND OPTION DELETE AFTER 
' -------------------------------------------------------------------
' ------------------------------------------------------------------
' -----------------------------------------------------------------
' ----------------------------------------------------------------

DIM DELETE_AFTER
DIM objFSO
Dim objShell
DIM ComputerName
DIM obj_SOURCE_Folder
DIM obj_TARGET_Folder
DIM Version_OS

Set objShell = Wscript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
WindowsDir = objShell.ExpandEnvironmentStrings("%windir%")
AppDataLocation=objShell.ExpandEnvironmentStrings("%APPDATA%")
ComputerName=objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

CALL GET_WINDOWS_VERSION

IF Version_OS=<5 THEN 
	'Wscript.Echo "begin."
	'Wscript.Echo "end."
	WScript.Quit(0)
END IF
' ----------------------------------------------------------------
' -----------------------------------------------------------------
' ------------------------------------------------------------------
' -------------------------------------------------------------------

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

i1 = 0 ' ANY PROGRAM WSCRIPT 
i2 = 0 ' MY  PROGRAM WSCRIPT 
FOR EACH objProcess IN colProcesses
    IF NOT (IsNull(objProcess.CommandLine )) THEN
		strScriptName = Trim(Right(objProcess.CommandLine,Len(objProcess.CommandLine) - InstrRev(objProcess.CommandLine,"\")))
		strScriptName = Left(strScriptName, Len(strScriptName) - 1)                             
		IF strScriptName = Wscript.ScriptName THEN
			i2=i2+1
		else
			i1=i1+1
		END IF
		IF i1 > 0 THEN ' ANY PROGRAM WSCRIPT 
				' Wscript.Quit
		END IF
		IF i2 > 1 THEN ' MY  PROGRAM WSCRIPT 
				Wscript.Quit
		END IF
    END IF
NEXT

' Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
' Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")
' KNOCK UP THIS FOR USE IN MY OUTLOOK SCRIPT CODING
IF 1=2 THEN 
	FOR EACH objProcess IN colProcesses
		IF NOT (IsNull(objProcess.ProcessID )) THEN
			strScriptProcessID = objProcess.ProcessID
			MSGBOX "MY PROCESS NUMBER PID IS -- "+ CSTR(strScriptProcessID)
		END IF
	NEXT
END IF




' -------------------------------------------------------------------
' ALL THE -- IF ____ Wscript.Quit ____ ROUTINE SET DONE
' -------------------------------------------------------------------

Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

'--------------------------------------------------------
' AFTER THE EFFORT 
' VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
' IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
' HERE VBS DOES SINGLE 
'--------------------------------------------------------

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

Dim objFolder

FOLDER_STRING_1="D:\0 CLOUD\GD-INSYNC\rub.rim@gmail.com\snapshot ("
FOLDER_STRING_2="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\snapshot ("
TARGET_FOLDER_1="D:\0 CLOUD\GD-INSYNC\rub.rim@gmail.com\snapshot"
TARGET_FOLDER_2="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\snapshot"

FOR RV=1 TO 2
	FOR RO=2 TO 20
		XXC=CSTR(RO)
		IF RV=1 THEN 
			FOLDER_STRING=FOLDER_STRING_1+XXC+")"
			TARGET_STRING=TARGET_FOLDER_1
		END IF
		IF RV=2 THEN 
			FOLDER_STRING=FOLDER_STRING_2+XXC+")"
			TARGET_STRING=TARGET_FOLDER_2
		END IF

		IF FSO.FOLDERExists(FOLDER_STRING) THEN
			Set objFolder = FSO.GetFolder(FOLDER_STRING)
				' MSGBOX "--"+FOLDER_STRING

			If objFolder.SubFolders.Count > 0 Then 
				CALL COPY_FOLDER_ONE_UP_ROOT_LEVEL
			END IF
		END IF
	NEXT
NEXT




FOLDER_STRING_1="D:\0 CLOUD\GD-INSYNC\rub.rim@gmail.com\snapshot ("
FOLDER_STRING_2="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\snapshot ("

FOR RV=1 TO 2
	FOR RO=2 TO 20
		XXC=CSTR(RO)
		IF RV=1 THEN 
			FOLDER_STRING=FOLDER_STRING_1+XXC+")"
			TARGET_STRING=TARGET_FOLDER_1
		END IF
		IF RV=2 THEN 
			FOLDER_STRING=FOLDER_STRING_2+XXC+")"
			TARGET_STRING=TARGET_FOLDER_2
		END IF

		IF FSO.FOLDERExists(FOLDER_STRING) THEN

			Set objFolder = FSO.GetFolder(FOLDER_STRING)

			'MSGBOX objFolder.SubFolders.Count
			'MSGBOX objFolder.Files.Count
			
			If objFolder.SubFolders.Count = 0 Then 
				IF objFolder.Files.Count = 0 Then
					' -------------------------------------------
					' THE ROOT FOLDER IS NOT EVER DELETE HERE 
					' AND NEXT ROUTINE IS SPECIAL DESIGN FOR THAT
					' ALL DELETE EMPTY WILL BE GONE
					' -------------------------------------------
					fso.DeleteFolder objFolder.Path
					IF FSO.FOLDERExists(FOLDER_STRING) THEN
						fso.DeleteFolder FOLDER_STRING
					END IF
				END IF
			END IF
		END IF
	NEXT
NEXT

' Wscript.Echo "ENDER"

WScript.Quit(0)

'--------------------------------------------------------------------
'END STOP -----------------------------------------------------------
'--------------------------------------------------------------------





Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

' MsgBox i_TIME
' MsgBox Format(Now, "YYYY-MMM-DD -- HH:MM:SS - HH:MM:SS-AM/PM DDD")
' 2019-Mar-18 -- 14:00:40 - 02:00:40-PM - Sun


' SOURCE CREDIT DELETE EMPTY
' ----
' Deleting Empty Folders – Random Musings of Jeremy Jameson
' https://blogs.msdn.microsoft.com/jjameson/2009/11/03/deleting-empty-folders/
' ----



'-------------------------------------
Sub COPY_FOLDER_ONE_UP_ROOT_LEVEL
'-------------------------------------
	DELETE_AFTER=TRUE
	obj_SOURCE_Folder = FOLDER_STRING
	obj_TARGET_Folder = TARGET_STRING ' +"\"
	
	' MSGBOX obj_SOURCE_Folder+VBCRLF +obj_TARGET_Folder

	SET_GO=TRUE
	If objFSO.FOLDERExists(obj_SOURCE_Folder) = FALSE Then SET_GO=FALSE
	If objFSO.FOLDERExists(obj_TARGET_Folder) = FALSE Then SET_GO=FALSE
	
	' MSGBOX SET_GO

	IF SET_GO=TRUE THEN 
		Set objSuperFolder = objFSO.GetFolder(obj_SOURCE_Folder)
		Call ShowSubfolders (objSuperFolder)
	END IF
END SUB

'-------------------------------------
Sub ShowSubFolders(fFolder)
'-------------------------------------
    Set objFolder = objFSO.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
    
	SET_GO=FALSE

	' --------------------------------------
	' COUNT BACKSLASH 
	' --------------------------------------
	' WANT ONLY FOLDER THAT ARE ONE UP FROM ROOT OF PATH REQUEST COPIER
	' --------------------------------------
    X1 = obj_SOURCE_Folder
    X2 = objFolder.PATH
	X1 = Len(X1) - Len(Replace(X1, "\", ""))
    X2 = Len(X2) - Len(Replace(X2, "\", ""))

	IF X1+1=X2 THEN SET_GO=TRUE
	
	IF SET_GO=TRUE THEN 
		ERR.CLEAR
		ON ERROR RESUME NEXT

		objFSO.COPYFOLDER objFolder.PATH, obj_TARGET_Folder+"\"+objFolder.NAME
		
		IF DELETE_AFTER=TRUE AND ERR.NUMBER=0 THEN 
			objFSO.DELETEFOLDER objFolder.PATH
		END IF
	End If
	
    For Each Subfolder in fFolder.SubFolders
        ShowSubFolders(Subfolder)
    Next
End Sub

'--------------------------------------------------------
SUB GET_WINDOWS_VERSION
'--------------------------------------------------------
	Version_OS=0
	ON ERROR RESUME NEXT
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
	for each System in SystemSet 
	' WScript.Echo System.Caption 
	' WScript.Echo System.Manufacturer 
	' WScript.Echo System.BuildType 
	' WScript.Echo " Version: " + System.Version 
	' WScript.Echo " Locale: " + System.Locale 
	' WScript.Echo " Windows Directory: " + System.WindowsDirectory 
	' WScript.Echo " Total memory: " + System.TotalVisibleMemorySize + " bytes" 
	' WScript.Echo " Serial Number: " + System.SerialNumber 
	' Wscript.Echo "" 
	Version_OS = CINT(MID(System.Version,1,INSTR(System.Version,".")-1))
	'WINDOWS 7 = 6.1.7601
	next 
	IF ERR.NUMBER>0 THEN Version_OS=0
	' CATCH THE ERROR
	' ---------------------------
	' Windows Script Host
	' ---------------------------
	' Script:	C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 13-COPY MOVE SHELL FOLDING.VBS
	' Line:	150
	' Char:	2
	' Error:	ActiveX component can't create object
	' Code:	800A01AD
	' Source: 	Microsoft VBScript runtime error
	' ---------------------------
END SUB

