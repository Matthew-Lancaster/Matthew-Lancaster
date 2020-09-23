'--------------------------------------------------------------
' __ CODE BY MATTHEW LANCASTER
' __ DATE BEGINNER 04 AM --------------------------------------
' __ WITH SOME HEAVY WIND 
' __ THROWING BITS FROM ONE PLACE TO ANOTHER 
' __ I AM THE WEAVER __ I AM THE WIND
' __ DATE ENDER -----------------------------------------------
' [ Friday 12:58:00 Pm_29 December 2017 ] The UK
' [ Friday 12:58:00 Pm_29 December 2017 ] GMT UTC+0 WET
'--------------------------------------------------------------
' END PROBLEM WITH CODE
' OBJECTIVE ABLE TO SEE FROM THE PROBLEM ALMOST
'--------------------------------------------------------------
' 1.
' VICEVERSA PRO SOFTWARE WORKING FROM MIXED FOLDER
' DOESN'T DO A REAL SYNCHRONISE WHEN MULTI FOLDER PATH SET
' ONLY KIND OF UPDATE
' IF REMOVE SOMETHING THE OTHER FOLDER WILL BRING IT BACK AGAIN
' SO NOT ABLE TO DO A MATRIX TYPE THING AFTER ALL
' DOES SORT IT OUT IF WANT A TIDY UP IN MANUAL MODE
' BUT AS I SAY ONLY BETWEEN ONE COMPUTER SO HARD WORK THERE
' 2.
' UNLESS MAKE SOME CROSS REFERENCING GOING ON
' 3.
' MOST THE CODE IS IN PLACE TO HELP WITH AWFUL THAT
' 4.
' MAYBE BEEN BETTER IN GOODSYNC BUT I GET TIRED OF TYPING 
' ALL THE PATH SET AND A LOT OF CROSS REFERENCING BUT TRUE SYNC
' 5. 
' AND THIS CODE GETS OVER PROBLEM OF MANY FOLDER AND DISH OUT THE 
' PLATE MENU WHEN THERE ARE STARTUP FOLDER AND THING THIS ONE 
' WAS DONE FOR SEND_TO 1ST
' SAVES WRITING MY OWN CODE AND HAVE SYNCHRONISE WORKER HARD WORK
' THE CODE WILL HAVE TO RUN TWICE FOR A FULL UPDATE SYNCHRONISER
' PRETTY MUCH LIKE ALL SYNCHRONISER WORK DO
' BE NICER OF GOODSYNC WORKED IN THE _.TIX FILE ENDING RATHER 
' THAN _.TIC ' AND THE BE _.KML EDITABLE THERE IS A SLIGHT WAY 
' AROUND IT WITH SAVE A JOB AS EXPORT TO XML AND BACK IMPORT AGAIN
'--------------------------------------------------------------
' 18 URLS OPEN WHILE WORKER HERE
'--------------------------------------------------------------

' Set WshShell = WScript.CreateObject("WScript.Shell")
' If WScript.Arguments.Length = 0 Then
  ' Set ObjShell = CreateObject("Shell.Application")
  ' ObjShell.ShellExecute "wscript.exe" , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  ' WScript.Quit
' End if
' ------------------------------------------------------------------
' REQUESTING RUNAS ADMINISTRATOR IN WIN XP BRING UP RUNAS DIALOG BOX
' ------------------------------------------------------------------
' SOME PART OF THIS CODE REQUEST ADMIN BUT HAVE BEEN CODED NOT TO IF 
' WIN XP
' ------------------------------------------------------------------

DIM USER_INTERACTIVE_MODE

Set args = Wscript.Arguments
For Each arg In args
	XF0 = XF0 + arg
Next

IF XF0="" THEN 
	USER_INTERACTIVE_MODE=TRUE
ELSE
	USER_INTERACTIVE_MODE=FALSE
END IF 
	
DIM FSO
SET FSO = CreateObject("Scripting.FileSystemObject")
DIM VV_PATH_1
DIM VV_PATH_2
DIM INDEX_VAR

DIM VV_VAR(3)
DIM VV_VAR_TRACKING(3)
' DONT RUN VV_VAR(1)
' DONT RUN VV_VAR(2)
VV_VAR(1)="SENDTO __ MATT 01 TO 04 __ "
VV_VAR(2)="START_MENU PROGRAM_DATA __ "

VV_VAR(3)="VV C DRIVE ROOT __ "


FOR i = 1 to UBOUND(VV_VAR_TRACKING)
	COUNTER_VAR=CSTR(I)
	IF LEN(COUNTER_VAR)=1 THEN COUNTER_VAR="0"&COUNTER_VAR
	VV_VAR_TRACKING(i)="VV_DB__"&COUNTER_VAR&"__"
NEXT

VV_EXE_PATH=""
VV_PATH_1="C:\SCRIPTER\SCRIPT_02 SYNC_ VICEVERSA\# VV COMPUTERNAME SCRIPTOR\"
VV_PATH_2="C:\SCRIPTER\SCRIPT_02 SYNC_ VICEVERSA\# TRACKING DATABASE\"

DIM RN_VAR(8)
RN_VAR(1)="2-ASUS-EEE"
RN_VAR(2)="1-ASUS-X5DIJ"
RN_VAR(3)="3-LINDA-PC"
RN_VAR(4)="4-ASUS-GL522VW"
RN_VAR(5)="5-ASUS-P2520LA"
RN_VAR(6)=""
RN_VAR(7)="7-ASUS-GL522VW"
RN_VAR(8)="8-MSI-GP62M-7RD"

CUR_USER = CreateObject("WScript.Network").UserName
COMPUTER_NAME_VAR = CreateObject("WScript.Network").Computername

FOR VV_4=1 TO UBOUND(RN_VAR)
	IF RN_VAR(VV_4)=COMPUTER_NAME_VAR THEN
		INDEX_VAR=RN_VAR(VV_4)
		EXIT FOR
	END IF
NEXT

DIM Version_OS
Const ForReading = 1, ForWriting = 2, ForAppending = 8

CALL GET_WINDOWS_VERSION

' MSGBOX Version_OS
' WSCRIPT.QUIT

IF Version_OS=>6 OR 1=1 THEN 
	CALL VICE_VERSA_COPY_OVER
	CALL WRITE_THE_TRACKING_DATABASE_IF_NOT_EXIST 
	CALL RUN_THE_VICE_VERSA_SCRIPT_THAT__CORRESPONDS_TO_THIS_COMPUTERNAME
END IF

WSCRIPT.QUIT

'--------------------------------------------------------
SUB GET_WINDOWS_VERSION
'--------------------------------------------------------
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
END SUB

'--------------------------------------------------------
SUB VICE_VERSA_COPY_OVER
'--------------------------------------------------------
DIM TT_EDITOR
DIM FILE_2
		
FOR VV_2=1 TO UBOUND(VV_VAR)
	FOR RN=1 TO UBOUND(RN_VAR)
		IF RN_VAR(RN)<>"" AND RN_VAR(RN)<>INDEX_VAR THEN 
		
		FILEN_1=VV_PATH_1 & VV_VAR(VV_2) & INDEX_VAR & ".fsf"
		FILEN_2=VV_PATH_1 & VV_VAR(VV_2) & RN_VAR(RN) & ".fsf"
		
		If FSO.FileExists(FILEN_2) = FALSE Then
			FSO.COPYFILE FILEN_1, FILEN_2
		END IF
		
		Set MyFile=FSO.getFile(FILEN_2) 
		SZ=MyFile.size
		set TS=MyFile.openAsTextStream(1)
		TT_EDITOR=TS.read(SZ)
		TS.close
		
		INDEX_VAR_2=""
		RN_VAR_2=""
		FOR RT_2=1 TO LEN(INDEX_VAR)
			INDEX_VAR_2=INDEX_VAR_2 & MID(INDEX_VAR,RT_2,1)+CHR(0)
		NEXT
		FOR RT_2=1 TO LEN(RN_VAR(RN))
			RN_VAR_2=RN_VAR_2 & MID(RN_VAR(RN),RT_2,1)+CHR(0)
		NEXT
		
		X = 1
		Do
			T1 = InStr(X, TT_EDITOR, INDEX_VAR_2)
			If T1 > 0 Then
				TT_EDITOR = Replace(TT_EDITOR, INDEX_VAR_2 , RN_VAR_2 , 1 , 1)
				Counter_VAR = Counter_VAR + 1
			End If
			X = T1 + 1
		Loop Until T1 = 0
		
		' PUT ANY OLD TRACKING DATABASE IN AND IT WILL FIND BY THE EXTENSION
		' AND ENABLE TO BE DONE BY YOU WHICH IS DEFAULT WHEN MAKE A DATABASE
		'-------------------------------------------------------------------
		'--------------------------------------------------
		Sync_Tracking_Database_VAR_1=VBCRLF&"Sync Tracking Database="
		Sync_Tracking_Database_VAR_2=Sync_Tracking_Database_VAR_1
		INDEX_VAR_2=""
		FOR RT_2=1 TO LEN(Sync_Tracking_Database_VAR_2)
			INDEX_VAR_2=INDEX_VAR_2 & MID(Sync_Tracking_Database_VAR_2,RT_2,1)+CHR(0)
		NEXT
		Sync_Tracking_Database_VAR_2=INDEX_VAR_2
		
		'UNICODE WON'T TAKE AN (= AND) NULL _ STRIP IT OFF
		Sync_Tracking_Database_MOD_VAR=MID(Sync_Tracking_Database_VAR_2,1,LEN(Sync_Tracking_Database_VAR_2)-2)
		
		INDEX_VAR_2=""
		VAR_tdb=".tdb"
		FOR RT_2=1 TO LEN(VAR_tdb)
			INDEX_VAR_2=INDEX_VAR_2 & MID(VAR_tdb,RT_2,1)+CHR(0)
		NEXT
		VAR_tdb=INDEX_VAR_2
		
		T1 = InStr(1, TT_EDITOR, Sync_Tracking_Database_MOD_VAR)
		T2 = InStr(1, TT_EDITOR, VAR_tdb)
		'msgbox t1
		'WSCRIPT.QUIT
		
		N3 = Mid(TT_EDITOR, T1, (T2 - T1))
		
		REPLACE_VAR=Sync_Tracking_Database_VAR_1&VV_PATH_2&VV_VAR_TRACKING(VV_2)&RN_VAR(RN)
		INDEX_VAR_2=""
		FOR RT_2=1 TO LEN(REPLACE_VAR)
			INDEX_VAR_2=INDEX_VAR_2 & MID(REPLACE_VAR,RT_2,1)+CHR(0)
		NEXT
		REPLACE_VAR=INDEX_VAR_2
		
		TT_EDITOR = Replace(TT_EDITOR, N3, REPLACE_VAR)
		'--------------------------------------------------
		
		Set MyFile=FSO.getFile(FILEN_2) 
		SZ=MyFile.size
		set TS=MyFile.openAsTextStream(1)
		TT_TEMP=TS.read(SZ)
		TS.close
				
		IF TT_TEMP<>TT_EDITOR THEN 
			set TS=MyFile.openAsTextStream(2)
			TS.write(TT_EDITOR)
			TS.close
		END IF 
		
		'MsgBox "Done Total " & Counter_VAR & " Proecesses"

		END IF
	NEXT
NEXT
END SUB


'--------------------------------------------------------
SUB WRITE_THE_TRACKING_DATABASE_IF_NOT_EXIST
'--------------------------------------------------------
	FOR VV_2=1 TO UBOUND(VV_VAR)
		FOR RN=1 TO UBOUND(RN_VAR)
			IF RN_VAR(RN)<>"" AND RN_VAR(RN)<>INDEX_VAR THEN 
				FILE_NAME_CHECKER=VV_PATH_2 & VV_VAR_TRACKING(VV_2) & RN_VAR(RN) & ".tdb"
				If FSO.FileExists(FILE_NAME_CHECKER) = FALSE Then
					Set objFile = FSO.CreateTextFile(FILE_NAME_CHECKER,True)
					objFile.Close
				END IF
			END IF
		NEXT
	NEXT		
END SUB

'-------------------------------------------------------------------
SUB RUN_THE_VICE_VERSA_SCRIPT_THAT__CORRESPONDS_TO_THIS_COMPUTERNAME
'-------------------------------------------------------------------

Set wshShell = CreateObject( "WScript.Shell" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

FOR VV_2=3 TO UBOUND(VV_VAR)
	FOR RN=1 TO UBOUND(RN_VAR)
		IF RN_VAR(RN)=strComputerName THEN 
		
		FILEN_2=VV_PATH_1 & VV_VAR(VV_2) & RN_VAR(RN) & ".fsf"
		
		Dim objShell
		Set objShell = Wscript.CreateObject("WScript.Shell")
		'Set ObjShell = Wscript.CreateObject("Shell.Application")
		' --------------------------------------------------
		' WELL I CAN'T FIGURE THIS BIT OUT AND I'M GIVING UP
		' --------------------------------------------------
		
		'WSCRIPT.QUIT

		' REQUESTING RUNAS ADMINISTRATOR IN WIN XP BRING UP RUNAS DIALOG BOX
		' ------------------------------------------------------------------
		'IF Version_OS=5 THEN

		IF USER_INTERACTIVE_MODE=FALSE THEN 
			iReturn=objShell.Run (CHR(34)& "C:\Program Files\ViceVersa Pro\ViceVersa.exe" & CHR(34)& " " & CHR(34)& FILEN_2 & CHR(34) & " /hiddenautoexec /wait",0,TRUE)
			
			'= 5 = CONFLICT
			'--------------
			IF IRETURN=5 THEN 
				iReturn=objShell.Run (CHR(34)& "C:\Program Files\ViceVersa Pro\ViceVersa.exe" & CHR(34)& " " & CHR(34)& FILEN_2 & CHR(34) & " /autocompare",0,TRUE)
			END IF

			IF IRETURN >0 AND IRETURN <>5 THEN 
				msgbox "VICE VERSA C ROOT = " & iReturn,vbMsgBoxSetForeground,Wscript.ScriptName
				iReturn=objShell.Run (CHR(34)& "C:\Program Files\ViceVersa Pro\ViceVersa.exe" & CHR(34)& " " & CHR(34)& FILEN_2 & CHR(34) & " /autocompare",0,TRUE)
			END IF
		ELSE
			msgbox "VICE VERSA C ROOT = USER_INTERACTIVE_MODE=TRUE",vbMsgBoxSetForeground,Wscript.ScriptName 
			iReturn=objShell.Run (CHR(34)& "C:\Program Files\ViceVersa Pro\ViceVersa.exe" & CHR(34)& " " & CHR(34)& FILEN_2 & CHR(34) & " /autocompare",0,TRUE)
		END IF
		
		
		'ELSE
			' objShell.Run "CMD", CHR(34)& "C:\Program Files\ViceVersa Pro\ViceVersa.exe" & CHR(34)& " " & CHR(34)& FILEN_2 & CHR(34) & " /hiddenautoexec /wait", , "runas", 1
		'END IF
		
		'/hiddenautoexec /wait"
		'/dialogautoexec /autoclose /wait", 6
		'/autoexec /autoclose
		
		Set objShell = Nothing
		
		END IF
	NEXT
NEXT
		
END SUB





'---------------------------
'REFERENCE SET
'---------------------------
'----
'vbscript - How to determine windows version from a VB script? - Stack Overflow
'https://stackoverflow.com/questions/4542284/how-to-determine-windows-version-from-a-vb-script
'----	

'----
'Find and replace NUL character using VBScript - Stack Overflow
'https://stackoverflow.com/questions/45933250/find-and-replace-nul-character-using-vbscript
'----

' -------------------------------------------------------------
' FORGET REF BEFORE USELESS
' NEXT UP - HOW TO OPEN A FILE AND FILE THE SAME AS IT WAS AFTER SOME MANIPULATION
' WITHOUT UNICODE HASSLE
' NEW CODE OPEN AS BINARY SO SEARCH AND REPLACE ARE UNICODE 
' FILE TODAY HERE WAS A BYTE BETWEEN EACH CHARACTER A 2 BYTE LENGTH A TYPE OF BASIC UNICODE
' -------------------------------------------------------------
' ----
' TRY TO WRITE FILE IN VBS WITH NULL CHARACTERS - Google Search
' https://www.google.co.uk/search?num=50&q=TRY+TO+WRITE+FILE+IN+VBS+WITH+NULL+CHARACTERS&spell=1&sa=X&ved=0ahUKEwjVk9-y5a7YAhUIBsAKHYBEC9oQvwUIJigA&biw=1536&bih=644
' --------
' handling binary files in VBScript - Google Groups
' https://groups.google.com/forum/#!topic/microsoft.public.scripting.vbscript/Ui_JyiuCP5k
' ----

' Set fold = FSO.GetFolder(strPath)
' Set files = fold.files


'----
'Run a vbscript from another vbscript - Stack Overflow
'https://stackoverflow.com/questions/1686454/run-a-vbscript-from-another-vbscript
'----
' Dim ProgramPath, WshShell, ProgramArgs, WaitOnReturn,intWindowStyle
' Set WshShell=CreateObject ("WScript.Shell")
' ProgramPath="c:\test run script.vbs"
' ProgramArgs="/hello /world"
' intWindowStyle=1
' WaitOnReturn=True
' WshShell.Run Chr (34) & ProgramPath & Chr (34) & Space (1) & ProgramArgs,intWindowStyle, WaitOnReturn

' intWindowStyle
' Description
' 0
' Hides the window and activates another window.
' 1
' Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
' 2
' Activates the window and displays it as a minimized window.
' 3
' Activates the window and displays it as a maximized window.
' 4
' Displays a window in its most recent size and position. The active window remains active.
' 5
' Activates the window and displays it in its current size and position.
' 6
' Minimizes the specified window and activates the next top-level window in the Z order.
' 7
' Displays the window as a minimized window. The active window remains active.
' 8
' Displays the window in its current state. The active window remains active.
' 9
' Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
' 10
' Sets the show-state based on the state of the program that started the application.