;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
;# __ 
;# __ Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME [ Tue 00:52:00 Am_12 Dec 2017 ]
;# LAST EDIT TIME [ Fri 16:08:00 Pm_27 Apr 2018 ]
;# __ 
;  =============================================================


; 002
; --------------------------------------------------------------
; FROM TO Sun 15-Apr-2018 19:26:10
; --------------------------------------------------------------
; THIS IS OUR PREFERRED DEFAULT OPTIONS FOR INSTALLING NOTEPAD++
; THE MIDDLE CHECKBOX IS SELECTED
; Reference's at End
; --------------------------------------------------------------

; 003
; --------------------------------------------------------------
; New work the GoodSync Routine to Set Number of Hour in Options
; From Default 2 to 4
; Straight Form The Help File Nothing Checked Online
; --------------------------------------------------------------
; FROM -- Fri 27-Apr-2018 13:44:49
; TO ---- Fri 27-Apr-2018 14:28:00 -- 50 MINUTE
; --------------------------------------------------------------
; --------------------------------------------------------------
; Well That Hard Work To Set the Content of an Edit Control 
; You Must Delete Whats in There First to Do with Caret Positioning
; And I Had to Look Up on the Internet
; Better than Using Sendinput key and Having to Be In Focus 
; Active Window __ Very Easy in the End
; Credit Due
; ----
; Delete edit control text ahk - Google Search
; https://www.google.co.uk/search?num=50&rlz=1C1CHWL_enGB794&ei=HDjjWvX4KI63gQab46zQDQ&q=Delete+edit+control+text+ahk&oq=Delete+edit+control+text+ahk&gs_l=psy-ab.3..35i39k1.5379.5647.0.6669.2.2.0.0.0.0.143.203.1j1.2.0....0...1c.1.64.psy-ab..0.2.201....0.9781WNbiPGQ
; --------
; Delete contents of edit control? - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/80399-delete-contents-of-edit-control/
; ----
; --------------------------------------------------------------
; FROM -- Fri 27-Apr-2018 14:28:00
; TO ---- Fri 27-Apr-2018 15:51:00 -- 1 HOUR 20 MINUTE
; SESSION PAIR 2 HOURS 10 MINUTE-ER
; --------------------------------------------------------------

; 004
; MORE WORK WANTED A SOUND EFFECT IF A WINDOW CAME OUT OF NOT RESPONDING 
; TO LET LEARN THE WAIT IS OVER FOR GOODSYNC NETWORK PATHS TIME CONSUMING
; TOOK A WHILE BECAUSE SOUND SYSTEM EVENTS WERE SWITCHED OFF
; --------------------------------------------------------------
; FROM -- Fri 27-Apr-2018 18:53:27
; TO ---- Fri 27-Apr-2018 19:33:58 -- 40 MINUTE
; SESSION PAIR 2 HOURS 10 MINUTE-ER
; --------------------------------------------------------------

; 005
; NEXT WANTER THE CHECKBOX SET WHEN GOODSYNC HAS THE TEXT BOX FILLED
; AND IT CHECK HWND OF WINDOW AND DOES ONCE 
; IT HARD TO TEST BECAUSE WHEN CHECKBOX NOT CHECKED THE TEXTBOX FIGUAR 
; ENTRY IS ENABLED=FALSE
; --------------------------------------------------------------
; FROM -- Fri 27-Apr-2018 19:33:58
; TO ---- Fri 27-Apr-2018 19:53:05 -- 20 MINUTE
; SESSION PAIR 2 HOURS 10 MINUTE-ER
; --------------------------------------------------------------

; 00*
; SESSION CODER _ ADD LOADER OF CAMERA REEL OFFLOAD
; TIMER_DRIVE_GET_CAMERA
; --------------------------------------------------------------
; FROM -- Thu 14-Jun-2018 17:15:27
; TO ---- Thu 14-Jun-2018 18:44:00 -- 1 HOUR 30 MINUTE
; --------------------------------------------------------------


; 007
; DFX REQUIRED WORKING PROPERLY AND EXTRA FOR IT
; --------------------------------------------------------------
; FROM -- Wed 20-Jun-2018 14:50:48
; TO ---- Wed 20-Jun-2018 15:02:00
; --------------------------------------------------------------

;# ------------------------------------------------------------------
; Location OnLine
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set Google Drive
; Possible Censorship of Code Detected By Google as Malicious Happen Here
; unlike DropBox that has All Available
; https://drive.google.com/open?id=0BwoB_cPOibCPTnRZZVFuRFpHOTg
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set DropBox
; https://www.dropbox.com/sh/ntghoncyb8py1tf/AACWYrfkVn9PlqpYzNNSMcpMa?dl=0
;--------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
;# ------------------------------------------------------------------


#Warn
#NoEnv
#SingleInstance Force
;--------------------
#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
;--------------------

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; SCRIPT ============================================================

DetectHiddenWindows, on
SetStoreCapslockMode, off

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

Set_Var_Responding_1=FALSE
Set_Var_Responding_2=FALSE
O_HWND_1=0
O_Status_GSDATA=1
SCRIPT_NAME_VAR=

GLOBAL VAR_STORE_CAMERA_LABEL
GLOBAL Secs_CAMERA
GLOBAL Label_CAMERA

VAR_STORE_CAMERA_LABEL=

TIMER_SUB_OWNER_SAVE_TIMER=0

; Each array must be initialized before use:
COMPUTER_NAME_ARRAY := []

UniqueID_Old=0

OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6
	

SCRIPT_NAME_VAR:=SubStr(A_ScriptName, 1, -4)
SCRIPT_NAME_VAR=%A_ScriptDir%\%SCRIPT_NAME_VAR%_TIMER_%A_ComputerName%.txt
SCRIPT_NAME_VAR=%SCRIPT_NAME_VAR%

If FileExist(SCRIPT_NAME_VAR)
{
	FileReadLine, TIMER_SUB_OWNER_SAVE_TIMER, %SCRIPT_NAME_VAR%, 1
}
	

setTimer TIMER_SUB_1,200

setTimer TIMER_SUB_EliteSpy, OFF

setTimer TIMER_SUB_GOODSYNC,1000
setTimer TIMER_SUB_GOODSYNC_OPTIONS,1000

setTimer TIMER_SUB_NOTEPAD_PLUS_PLUS,1000

setTimer TIMER_SUB_WINDOWS_UPDATE,OFF
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_WINDOWS_DESKTOP_ICON,1000
; STARTS AS 1 SECOND AND THEN GOES TO 10 SECOND

setTimer TIMER_SUB_VICE_VERSA,OFF
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_SCRIPT_SHELL_FOLDERING,1000
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

;setTimer TIMER_SUB_OWNER, % -1 * 1000 * 60 * 60 ; After1Hours
setTimer TIMER_SUB_OWNER, 1000 ; After1Hours
; STARTS AS AFTER 1 HOUR AND THEN GOES TO EVERY 24 HOUR

;setTimer TIMER_SUB_WSCRIPT,100

;setTimer TIMER_SUB_CMD_KILL, OFF
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_LOGGER, 1000
; LOGGING BLUETOOTH AND DUPLICATE CLEANER LOGGER TRUNCATE
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB__SendSMTP__0__LOG_BAT,1000
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_I_VIEW32_CONVERT_CCSE,1000
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HALF HOUR

setTimer TIMER_SUB__MY_IP, 1000
; STARTS AS 1 SECOND AND THEN GOES TO EVERY 10 MINUTE

setTimer TIMER_PREVIOUS_INSTANCE, 1
; -------------------------------------------------------------------
; I used PREVIOUS_INSTANCE Detection For My Scripts Now 
; Otherwise they Maybe Two or More Loaded When You Run a Whole Bunch Together Quickly at Bootup or Just to run Again after Killed the Lot or Something
; Sometimes On Less Quicker Machine Computer
; There Might Be a Bug with This at Display the Icon in System Tray Not Sure Yet Maybe it was at a Time when My System was Instability
; -------------------------------------------------------------------



SETTIMER TIMER_DRIVE_CAMERA_UPLOAD_DROPBOX,4000

RETURN

TIMER_DRIVE_CAMERA_UPLOAD_DROPBOX:
IfWinExist Camera Upload
{
	WinActivate
	sendinput, !{F4}		; CLOSE
	SoundBeep , 1000 , 50
}
RETURN


TIMER_SUB_WINDOWS_DESKTOP_ICON:

setTimer TIMER_SUB_WINDOWS_DESKTOP_ICON,59000

;C:\Windows10Upgrade\Windows10UpgraderApp.exe /ClientID "Win10Upgrade:VNL:NHV13SIH:{}"

FN_VAR:="E:\01 Desktop\#_%A_ComputerName%\Windows 10 Update Assistant.lnk"
IfExist, %FN_VAR%
	{
		;SoundBeep , 2000 , 200
		;FileDelete, %FN_VAR%
	}
	

FN_VAR:="E:\01 Desktop\#_%A_ComputerName%\GoodSync Explorer.lnk"
IfExist, %FN_VAR%
	{
		SoundBeep , 2000 , 200
		FileDelete, %FN_VAR%
	}
RETURN


Multiple_Thread_Port_Scanner_ROUTINE:
; ---------------------------------------------------
; NOT USED IN CODE YET BUT PRACTICED FOR CODE BUILDER
; ---------------------------------------------------
; FROM    Thu 03-May-2018 16:16:39
; TO      Thu 03-May-2018 17:58:00 __ 1 HOUR 40 MINUTE
; ---------------------------------------------------
; I WROTE THIS CODE FROM IDEA WORKING ON CODE BEFORE
; AND SIDE TRACKED INTO THIS ONE
; AND NEAR THE END DECIDED NOT GOING TO USE IT
; BUT PRACTICE WAS HELPFUL
; ---------------------------------------------------

IfNotExist, %A_TEMP%\IPTEST.TXT
{
FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\Multiple_Port_Scanner.exe"
IfExist, %FN_VAR%
	{
		Run %comspec% /c ""%FN_VAR%" "/ALL" "" >"%A_TEMP%\IPTEST.TXT , , MIN
	}
}

ArrayCount = 0
; Write to the array:
Loop, Read, %A_TEMP%\IPTEST.TXT
{
	SET_GO=FALSE
	IF SubStr(A_LoopReadLine, 1 , 2)="\\"
		SET_GO=TRUE
	
	IP_LINE_1=%A_LoopReadLine%
	IP_LINE_2:=SubStr(A_LoopReadLine, 3)
	IP_LINE_2:=StrReplace(IP_LINE_2, "-" , "_")
	;StringLower, IP_LINE_2, IP_LINE_2
	;StringLower, IP_LINE_1, IP_LINE_1
	
	IP_LINE_3:="_03_fat32_4gb"
	
	IF SET_GO=TRUE
		{
		SET_GO=FALSE
		ifExist, %IP_LINE_1%\%IP_LINE_2%%IP_LINE_3%\*.*
			{
			SET_GO=TRUE
			IP_LINE_4=%IP_LINE_1%\%IP_LINE_2%%IP_LINE_3%
			;MSGBOX % IP_LINE_4
			}
		}
			
	IF SET_GO=TRUE
	{
		ArrayCount += 1  
		COMPUTER_NAME_ARRAY%ArrayCount% = %IP_LINE_4%
	}
}

; Read from the array:
Loop %ArrayCount%
{
    MsgBox % "Element number " . A_Index . " is " . COMPUTER_NAME_ARRAY%A_Index%
}

RETURN


;----------------------------------------
TIMER_SUB_NOTEPAD_PLUS_PLUS:
;----------------------------------------
DetectHiddenWindows, on
SetTitleMatchMode 2  ; Avoids Specify Full path.

IfWinExist Notepad++
{
	ControlGetText, OutputVar, Allow plugins to be loaded from, Notepad++
	IF OutputVar 
	{	
		ControlGet, Status, Checked,, Button5
		If Status = 0
		{
			Control, Check,, Button5
			SoundBeep , 4000 , 100
		}
	}
}
Return

TIMER_SUB_EliteSpy:
;----------------------------------------

DetectHiddenWindows, on
SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

IfWinNotExist EliteSpy+ by Andrea
{
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
	IfExist, %FN_VAR%
		{
			Run, "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
		}
}
Return

;----------------------------------------
TIMER_SUB_GOODSYNC_OPTIONS:
;----------------------------------------
DetectHiddenWindows, on
SetTitleMatchMode 2  ; Avoids Specify Full path.

WinGet, HWND_1, ID, ] Options ahk_class #32770
	IF (HWND_1>0)
	{
		; WinGet, OutputVar, ControlList, ahk_id %HWND_1%
		; Tooltip, % OutputVar ; List All Controls of Active Window
		;---------------------------------------------------------
		ControlGettext, OutputVar_2, Button16, ahk_id %HWND_1%

		ControlGet, OutputVar_1, Line, 1, Edit9, ahk_id %HWND_1%
		
		WinGetTitle OutputVar_3,ahk_id %HWND_1%
		
		If (OutputVar_1 = 2
			and OutputVar_2="Periodically (On Timer), every")
			{
				ControlSetText, Edit9,, ahk_id %HWND_1%
				Control, EditPaste, 4, Edit9, ahk_id %HWND_1%
				SoundBeep , 4000 , 100

		}
		if O_HWND_1<>%HWND_1%
		{
			ControlGet, OutputVar_4, Visible, , Button16, ahk_id %HWND_1%
			ControlGet, Status, Checked,, Button16, ahk_id %HWND_1%
			If Status=0
			{
				Control, Check,, Button16, ahk_id %HWND_1%
				IF OutputVar_4=1
					SoundBeep , 4000 , 100
			}
		}
			
		ControlGettext, OutputVar_2, Button20, ahk_id %HWND_1%
		ControlGet, OutputVar_1, Line, 1, Edit2, ahk_id %HWND_1%
		
		If (!OutputVar_1 
			and OutputVar_2="Wait for Locks to clear, minutes")
			{
				ControlSetText, Edit12,, ahk_id %HWND_1%
				Control, EditPaste, 10, Edit2, ahk_id %HWND_1%
				SoundBeep , 4000 , 100

		}
		ControlGet, Status, Checked,, Button20, ahk_id %HWND_1%
		If Status=1
		{
			Control, UnCheck,, Button22, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}

		;------------------------------------------------------------
		; Button5 ---- Save deleted/replaced files to History f (...)
		; Button6 ---- Cleanup _history_ folder after this many (...)
		; IF BUTTON 5 SET THEN ALSO SET BUTTON 6 _ DON'T LEAVE INFINITE
		;------------------------------------------------------------
		ControlGet, Status, Checked,, Button5, ahk_id %HWND_1%
		If Status = 1
		{
			ControlGet, Status, Checked,, Button6, ahk_id %HWND_1%
			If Status = 0
			{
			Control, Check,, Button6, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
			}
		}

		;------------------------------------------------------------
		; ---- HISTORY SET DAYS TO 30
		;------------------------------------------------------------
		Var_check=[VB
		if (SubStr(OutputVar_3, 1, 3)=Var_check)
		{
		ControlGet, Status, Checked,, Button4, ahk_id %HWND_1%
		If Status = 1
		{
			ControlGet, OutputVar_1, Line, 1, Edit2, ahk_id %HWND_1%
			ControlGet, OutputVar_4, Enabled, , Edit2, ahk_id %HWND_1%
			
			If (Trim(OutputVar_1)<>30 and OutputVar_4=1)
				{
					ControlSetText, Edit2,, ahk_id %HWND_1%
					Control, EditPaste, 30, Edit2, ahk_id %HWND_1%
					SoundBeep , 4000 , 100
			}
		}
		}
		
		;-----------------------------------------------------
		; Button3
		; Text:	Save deleted/replaced files to Recycle B (...)
		; Button4
		; Text:	Cleanup _saved_ folder after this many d (...)
		;-----------------------------------------------------
		Var_check=[VB
		if (SubStr(OutputVar_3, 1, 3)=Var_check)
		{
		ControlGet, Status, Checked,, Button4, ahk_id %HWND_1%
		If Status = 1
		{
			sleep, 500
			ControlGet, OutputVar_1, Line, 1, Edit1, ahk_id %HWND_1%
			ControlGet, OutputVar_4, Visible, , Edit1, ahk_id %HWND_1%
			If (Trim(OutputVar_1)<>30 and OutputVar_4=1)
				{
					ControlSetText, Edit1,, ahk_id %HWND_1%
					Control, EditPaste, 30, Edit1, ahk_id %HWND_1%
					SoundBeep , 4000 , 100
			}
		}
		}
		
		;-----------------------------------------------------
		; CAN'T DO THIS ONE 
		; IT'S EITHER HISTORY OR SAVED NOT BOTH
		; WELL YOU CAN DO BUTTON4 SAVE DO-ER INFINITE
		; Button3
		; Text:	Save deleted/replaced files to Recycle B (...)
		; Button4
		; Text:	Cleanup _saved_ folder after this many d (...)
		; Button6
		; Cleanup _history_ folder after this many (...)
		; BUTTON4 AND BUTTON6 CAN ALWAYS BE ON
		;-----------------------------------------------------
		;ControlGet, Status, Checked,, Button3, ahk_id %HWND_1%
		;If Status = 0
		;{
		;	Control, Check,, Button3, ahk_id %HWND_1%
		;	SoundBeep , 4000 , 100
		;}
		ControlGet, Status, Checked,, Button4, ahk_id %HWND_1%
		If Status = 0
		{
			Control, Check,, Button4, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}
		ControlGet, Status, Checked,, Button6, ahk_id %HWND_1%
		If Status = 0
		{
			Control, Check,, Button6, ahk_id %HWND_1%
			SoundBeep , 4000 , 100
		}
		
		
		IF O_HWND_1<>%HWND_1%
		{
			O_Status_GSDATA=0
		}
		;------------------------------------------------------------
		; ---- No _gsdata_ folder here
		;------------------------------------------------------------
		Var_check=[VB
		if (SubStr(OutputVar_3, 1, 3)=Var_check)
		{
			ControlGet, Status, Checked,, Button41, ahk_id %HWND_1%
			If Status = 1
			IF O_Status_GSDATA<>%Status%
			{
				Control, unCheck,, Button41, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
				ControlGet, Status, Checked,, Button41, ahk_id %HWND_1%
				If Status=0
					O_Status_GSDATA=1
			}
		}
		
		;------------------------------------------------------------
		;Button10 ---	Text:	Exclude empty folders
		;Button11 --- Text:	Exclude Hidden files and folders
		;Button12 -- Text:	Exclude System files and folders
		;------------------------------------------------------------
		Var_check=[VB EXE SYNC
		if (SubStr(OutputVar_3, 1, 12)=Var_check)
		{
			ControlGet, Status, Checked,, Button10, ahk_id %HWND_1%
			If Status = 0
			{
				Control, Check,, Button10, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}
			ControlGet, Status, Checked,, Button11, ahk_id %HWND_1%
			If Status = 0
			{
				Control, Check,, Button11, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}
			ControlGet, Status, Checked,, Button12, ahk_id %HWND_1%
			If Status = 0
			{
				Control, Check,, Button12, ahk_id %HWND_1%
				SoundBeep , 4000 , 100
			}
		}
		
		
		
		O_HWND_1=%HWND_1%
		
}

;Are you sure you want to move _GSDATA_ folder back to the sync folder?
;This option is for Advanced users and you should read the manual first.
;Yes
;No

;WinGet, HWND_1, ID, GoodSync Warning ahk_class #32770
WinGet, HWND_1, ID, GoodSync ahk_class #32770
WinGetText OutputVar_3,ahk_id %HWND_1%
;tooltip % HWND_1

IfInString, OutputVar_3, Are you sure you want to move _GSDATA_
{
	#WinActivateForce, ahk_id %HWND_1%
	ControlClick, Button2,ahk_id %HWND_1%
	SoundBeep , 4000 , 100
}

IfInString, OutputVar_3, Are you sure you want to keep _GSDATA_
{
	#WinActivateForce, ahk_id %HWND_1%
	ControlClick, Button2,ahk_id %HWND_1%
	SoundBeep , 4000 , 100
}
Return

TIMER_SUB_GOODSYNC:
;----------------------------------------
;setTimer TIMER_SUB_GOODSYNC, OFF
IF (TRUE=FALSE)
{
	Process, Exist, GoodSync-v10.exe
	If Not ErrorLevel
		{
		FN_VAR:="C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
		IfExist, %FN_VAR%
			{
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			Run, "%FN_VAR%" , , MIN
			}
		}
}

DetectHiddenWindows, oFF
SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

IfWinExist GoodSync - Preparing Crash Report
{
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	Run, "TASKKILL.exe" /F /IM GoodSync-v10.exe /T , , HIDE
}

IfWinExist Reporting a Crash
{
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	SoundBeep , 4000 , 100
	SoundBeep , 3000 , 100
	Run, "TASKKILL.exe" /F /IM GoodSync-v10.exe /T , , HIDE
}

;OK
;GoodSync has crashed just now and we are sorry for that.
;Click OK to assemble Crash Report.
;You will be given option to submit it to Siber Systems.
;Submitting Crash Reports helps us in fixing these crashes.

DetectHiddenWindows, OFF

IF (TRUE=FALSE)
{
	IfWinExist GoodSync
	{
		ControlGetText, OutputVar, GoodSync has crashed just now , GoodSync
		IF OutputVar 
		{
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			SoundBeep , 4000 , 100
			SoundBeep , 3000 , 100
			ControlClick, OK, GoodSync
			;WINHIDE
			
			;Run, "TASKKILL.exe" /F /IM GoodSync-v10.exe /T , , HIDE
		}
	}
}
Return

;----------------------------------------
TIMER_SUB_WSCRIPT:
DetectHiddenWindows, ON

IfWinExist Windows Script Host
{
	ControlClick, OK, Windows Script Host
	SoundBeep , 2500 , 100
}
Return


TIMER_SUB__MY_IP:
setTimer TIMER_SUB__MY_IP, % -1 * 1000 * 60 * 10 ; After10Minute

FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 23-MY IP.VBS"
IfExist, %FN_VAR%
	{
		Run, %FN_VAR%
	}
RETURN

TIMER_SUB__SendSMTP__0__LOG_BAT:
setTimer TIMER_SUB__SendSMTP__0__LOG_BAT, % -1 * 1000 * 60 * 60 ; 1 HOUR

FN_VAR:="C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP__0__LOG.BAT"
IfExist, %FN_VAR%
	{
		Run, %FN_VAR%, , MIN
	}
RETURN

; -------------------------------------------------------------------
TIMER_SUB_I_VIEW32_CONVERT_CCSE:

setTimer TIMER_SUB_I_VIEW32_CONVERT_CCSE, % -1 * 1000 * 60 * 30 ; HALF HOUR

if OSVER_N_VAR<10
{
	setTimer TIMER_SUB_I_VIEW32_CONVERT_CCSE,off
	RETURN
}


FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 24-I_VIEW32 CONVERT_CCSE.AHK"
IfExist, %FN_VAR%
	{
		Run, %FN_VAR%
	}
RETURN



;----------------------------------------
;----------------------------------------
;----------------------------------------
;----------------------------------------
TIMER_SUB_1:
	
DetectHiddenWindows, on
SetTitleMatchMode 2



WinGet, ID, list,ahk_class ConsoleWindowClass
IF (ID)
{
	Process, Exist, TASKLIST.EXE
	NewPID = %ErrorLevel%  ; Save the value immediately ErrorLevel is often changed
	If NewPID > 0
	{
		SoundBeep , 3000 , 100
		Process, priority, %NewPID%, Realtime
	}

	Process, Exist, TASKKILL.EXE
	NewPID = %ErrorLevel%  ; Save the value immediately ErrorLevel is often changed
	If NewPID > 0
	{
		SoundBeep , 2000 , 100
		Process, priority, %NewPID%, Realtime
	}
}



if (WinExist("Output ahk_class #32770"))
{
WinGet, HWND, ID, Output ahk_class #32770
WinGet, PATH, ProcessName, ahk_id %HWND%
IF PATH=ViceVersa.exe
	WINCLOSE
}


; -------------------------------------------------------------------
; [ Sunday 22:34:10 Pm_15 April 2018 ]
; Time Spend 
; Sun 15-Apr-2018 20:38:30 <> 
; Sun 15-Apr-2018 22:38:00 = 2 Hours
; I Frequently Want to swap between show Hidden Files in the System
; But it comes up with a Nagg Screen Warning
; But It Doesn't Anymore
; Question With The Answer _ Answer Sewer
; ---------------------------------------------------
; Hide Protected Operating System Files (Recommended)
; ---------------------------------------------------
; I wasn't Able to Make it Check Against Any Text Other than Warning
; Not any Control Text or Other Windows Text
; But control Copy Gets the Text for The Clipboard 
; There Was the Ability to Check if a Yes No Question was on the MsgBox
; But I Had to Do the Best I Able Do to Safeguard Hitting Another Window
; Project with More Subroutine Set
; Located Here Folder and Source File
; https://www.dropbox.com/sh/h2ebk12dksaq7j3/AAD9Ow_SbBP33JKmuALRkO1_a?dl=0
; https://www.dropbox.com/s/yuspt7npqj0jpuo/Autokey%20--%2019-SCRIPT_TIMER_UTIL.ahk?dl=0
; -------------------------------------------------------------------

if (WinExist("Warning ahk_class #32770"))
{
	; Test Line Remmed
	; WinGet, OutputVar, ControlList, ahk_id %HWND%
	; Tooltip, % OutputVar

	WinGet, HWND, ID, Warning ahk_class #32770
	WinGetClass, This_Class, ahk_id %HWND%
	WinGet, path, ProcessName, ahk_id %HWND%
	WinGetText, OutputVar, ahk_id %HWND%
	WinGetPos, WinLeft, WinTop, WinWidth, WinHeight, ahk_id %HWND%

	IF (OSVER_N_VAR=5) 
	{
		WinWidthOS=WinWidth
		WinHeightOS=WinHeight
	}
	IF (OSVER_N_VAR=6)  
	{
		WinWidthOS=572
		WinHeightOS=201
	}
	IF (OSVER_N_VAR=10)
	{
		WinWidthOS=713
		WinHeightOS=256
	}
	
	; Include Multiple If-And Statement With Separated Lines
	; ---------------------------------
	IF (This_Class="#32770" 
	and PATH="explorer.exe" 
	and OutputVar="&Yes`r`n&No`r`n" 
	and WinWidth=WinWidthOS
	and WinHeight=WinHeightOS)
	{
		Control, Check,, Button1
		SoundBeep , 4000 , 100
	}
}

DetectHiddenWindows, off

;Would you like to switch to the following audio playback device?
IfWinExist FxSound Message
{
	;WinGet, OutputVar, ControlList, FxSound Message
	;Tooltip, % OutputVar ; List All Controls of Active Window
	;---------------------------------------------------------
	ControlGettext, OutputVar_2, Static1,  FxSound Message
	
	IfInString, OutputVar_2, Would you like to
	{
		SoundBeep , 3000 , 100
		SoundBeep , 2000 , 100
	    Control, Check,, Button4, FxSound Message ahk_class #32770
	    ControlClick, Button2, FxSound Message ahk_class #32770
	}
}



IfWinExist File Access Denied ahk_class #32770
{
	;WinGet, OutputVar, ControlList, FxSound Message
	;Tooltip, % OutputVar ; List All Controls of Active Window
	;---------------------------------------------------------
	ControlGettext, OutputVar_2, SysLink1, File Access Denied ahk_class #32770
	
	; MSGBOX % OutputVar_2
	
	IfInString, OutputVar_2, You'll need to provide administrator
	{
	    ControlClick, Button1, File Access Denied ahk_class #32770
		SoundBeep , 3000 , 100
		SoundBeep , 2000 , 100
	}
}

IfWinExist File Access Denied ahk_class OperationStatusWindow
{

	WINACTIVATE, File Access Denied ahk_class OperationStatusWindow
	SENDINPUT {ENTER}
	
	; ControlClick, Continue, File Access Denied ahk_class OperationStatusWindow
	SoundBeep , 3000 , 100
	SoundBeep , 2000 , 100
}


DetectHiddenWindows, on


ifWinNotExist, ahk_class wndclass_desked_gsk
{
	if (WinExist("(Not Responding)") and Set_Var_Responding_1="FALSE")
	{
		Set_Var_Responding_1=TRUE
		SoundBeep , 1000 , 100
		SoundBeep , 2000 , 100
	}

	if (!WinExist("(Not Responding)") and Set_Var_Responding_1="TRUE")
	{
		Set_Var_Responding_1=FALSE
		SoundBeep , 3000 , 100
		SoundBeep , 2000 , 100
		SoundBeep , 2500 , 100
		SoundBeep , 2000 , 100
	}
}

if (WinExist("Page Unresponsive") and Set_Var_Responding_2="FALSE")
{
	Set_Var_Responding_2=TRUE
	SoundBeep , 1000 , 100
	SoundBeep , 2000 , 100
}

if (!WinExist("Page Unresponsive") and Set_Var_Responding_2="TRUE")
{
	Set_Var_Responding_2=FALSE
	SoundBeep , 3000 , 100
	SoundBeep , 2000 , 100
	SoundBeep , 2500 , 100
	SoundBeep , 2000 , 100
}

DetectHiddenWindows, off

IfWinExist RoboForm Upgrade
{
	WINCLOSE
	;WinActivate
	;sendinput, !{F4}		I
	;SoundBeep , 2500 , 100
}


IfWinExist ahk_class #32770
{

	HWND_ID_1 := WinExist("ahk_class #32770")
	HWND_ID_2 := WinExist("AutoSave Login - RoboForm ahk_class #32770")
	HWND_ID_3 := WinExist("Save - RoboForm ahk_class #32770")
	HWND_ID_4 := WinExist("Install RoboForm ahk_class #32770")
	HWND_ID_5 := WinExist("Sync RoboForm Data Folder")
	HWND_ID_6 := WinExist("AutoFill - RoboForm")
	
	ControlGetText, OutputVar_1, _RoboForm_Dialog_1100973_, ahk_class #32770
	
	; -------------------------------------------------------
	; THIS IS THE THIN TASK BAR UNDERNEATH THE APP'S
	; NOT MUCH TO GO ON TAKEN MAX THOUGH _ FILTER THIS ONE TO 
	; IGNORE TO ACT ON
	; MORE EFFORT ONLY ABLE USE DON'T SAVE WORD
	; -------------------------------------------------------
	; Save
	; Don't Save
	; _RoboForm_Dialog_1100973_
	; -------------------------------------------------------
	ControlGetText, OutputVar_2, Don't Save, ahk_class #32770
	; ControlGetText, OutputVar_2, Save`r`nDon't Save`r`n_RoboForm_Dialog_1100973_, ahk_class #32770
	
	; Fill Empty Fields &Only
	; _RoboForm_Dialog_1100973_
	; -------------------------------------------------------
	ControlGetText, OutputVar_3, Fill Empty Fields, ahk_class #32770
	
	SET_GO=FALSE
	IF OutputVar_1
		SET_GO=TRUE
	IF OutputVar_2
		SET_GO=FALSE
	IF OutputVar_3
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_2%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_3%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_4%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_5%
		SET_GO=FALSE
	IF HWND_ID_1=%HWND_ID_6%
		SET_GO=FALSE
		
	IF SET_GO=TRUE
	{	
		SoundBeep , 2500 , 100
		;MSGBOX REQUIRE HELP HERE ROBOFORM SEARCHING _RoboForm_Dialog_1100973_
		;ControlClick, No, ahk_class #32770
	}
}

IfWinExist RoboForm Update ahk_class #32770
{
	SoundBeep , 2500 , 100
	ControlClick, No, RoboForm Update ahk_class #32770
}

IfWinExist ahk_class #32770
{
	ControlGetText, OutputVar, Button1, ahk_class #32770
	IfInString, OutputVar, Install
	{	
	WINCLOSE ahk_class #32770 ahk_exe robotaskbaricon.exe
	;ControlClick, Button1, ahk_class #32770
	}
}

;This operation will affect the entire list, not just the current filter. Continue?
IfWinExist DuplicateCleaner ahk_class #32770
{
	WinGetText, OutputVar, DuplicateCleaner ahk_class #32770
	IfInString, OutputVar, This operation will affect the entire list
	{	
	ControlClick, Button1, DuplicateCleaner ahk_class #32770
	}
}


DetectHiddenWindows, on



;6 days 22 hours left in trial
;ahk_class #32770
;ahk_exe dfx.exe
DetectHiddenWindows, off
IfWinExist left in trial
{
	WinActivate
	sendinput, !{F4}		I
	SoundBeep , 1000 , 50
}


SetTitleMatchMode 3  ; Exactly

; DFXSTATIC
; ..\..\UTIL\doubleBuff\doubleBuffInit.cpp
IfWinExist 38 ahk_class #32770
{
	; ControlGetText, OutputVar, doubleBuffInit , 38 ahk_class #32770
	; ControlGetText,Static1,doubleBuffInit
	; ControlGettext, OutputVar, Static1, 38 ahk_class #32770
	; MSGBOX % OutputVar 
	; IF OutputVar 

	ControlClick, OK, 38 ahk_class #32770
	SoundBeep , 1000 , 50
}
IfWinExist 58 ahk_class #32770
{
	ControlClick, OK, 58 ahk_class #32770
	SoundBeep , 1000 , 50
}
IfWinExist 158 ahk_class #32770
{
	ControlClick, OK, 158 ahk_class #32770
	SoundBeep , 1000 , 50
}

SetTitleMatchMode 2

;---------------------------------------
;---------------------------------------
;____ TEAM_VIEWER
DetectHiddenWindows, on
IfWinExist Sponsored session
{
	;WinActivate
	;sendinput, !{F4}		I
	SoundBeep , 2500 , 100
	
	;0x10 CLOSE
	;PostMessage, 0x10, 0, 0,, Sponsored session
	
	;PostMessage, 0x112, 0xF060,,, WinTitle, WinText  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
	;PostMessage, 0x112, 0xF060,,, ,Sponsored session,   ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE

	ControlClick, OK, Sponsored session
}
;--------------------------------------------------------------------
; TEAM VIEWER CLOSE AFTER CONNECTION ADVERT
;--------------------------------------------------------------------

DetectHiddenWindows, OFF
IfWinExist ahk_class ATL:03A50E50
{
	ControlClick, OK, ahk_class ATL:03A50E50
}

IfWinExist Commercial use ahk_class #32770
{
	;ControlClick, Button4, Commercial use ahk_class #32770
	ControlClick, OK, Commercial use ahk_class #32770
}


IfWinExist Session timeout! ahk_class #32770
{
	ControlClick, OK, Session timeout! ahk_class #32770
}



DetectHiddenWindows, off
SetTitleMatchMode 3  ; Exactly
;Visual BASIC
IfWinExist Visual Component Manager
{
	ControlClick, OK, Visual Component Manager
	SoundBeep , 2500 , 100
}

DetectHiddenWindows, on
SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

DetectHiddenWindows, off
;Visual BASIC
IfWinExist Data View
{
	SoundBeep , 2500 , 100
	ControlClick, OK, Data View
}

DetectHiddenWindows, on

IfWinExist hubiC
{
	ControlGetText, OutputVar, hubiC is already running. , hubiC
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, OK, hubiC
	}
}

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.
IfWinExist hubiC
{
	ControlGetText, OutputVar, If you exit hubiC now , hubiC
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, &Yes, hubiC
	}
}

IfWinExist RoboForm Question
{
	ControlGetText, OutputVar, Do you want to delete
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, OK, RoboForm Question
	}
}


IfWinExist Microsoft OneDrive
{
	ControlGetText, OutputVar, Close OneDrive , Microsoft OneDrive
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, Close OneDrive, Microsoft OneDrive
	}
}

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

IfWinNotExist, Replace
	UniqueID_Old=0

IfWinExist, Replace
{
	; Replace Search and Replace Move to Better Position In Visual Basic
	; or any other editor
	ControlGetText, OutputVar, Current &Procedure , Replace
	IF OutputVar 
	{	
		UniqueID := WinExist("Replace")
		;tooltip %UniqueID%
		WinGetPos,,YPos,,, Replace
		if (YPOS>(A_ScreenHeight/2))
			UniqueID_Old=0
		if UniqueID_Old<>%UniqueID%
		{
			;WinMove, Replace,, (A_ScreenWidth/2)-(Width/2), (A_ScreenHeight/2)-(Height/2)
			SoundBeep , 2500 , 50
			WinGetPos,,YPos, Width, Height, Replace
			WinMove, Replace,, (A_ScreenWidth)-(Width), 0

			WinGetPos,,YPos, Width, Height, Replace
			if (YPOS<>0)
				UniqueID=0
		}
		
		UniqueID_Old=%UniqueID%
	}
}


;DuplicateCleaner
IfWinExist Finished deleting files.
{
	SoundBeep , 2500 , 100
	ControlClick, OK, Finished deleting files.
}


;Are you sure you want to cancel the scan?
IfWinExist DuplicateCleaner
{
	ControlGetText, OutputVar, Are you sure you want to cancel , DuplicateCleaner
	IF OutputVar 
	{	
		SoundBeep , 2500 , 100
		ControlClick, &Yes, DuplicateCleaner
	}
}

;DuplicateCleaner
IfWinExist Scan cancelled
{
	SoundBeep , 2500 , 100
	ControlClick, Close, Scan cancelled
}



SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

IfWinExist MSDN Library Visual
{
	ControlGetText, OutputVar, Active Su&bset , MSDN Library Visual
	IF OutputVar 
	{	
		;Esc::
		;{
		;	SoundBeep , 3500 , 100
		;	PostMessage, 0x112, 0xF060,,, MSDN Library Visual  ; 0x112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
		;}
		;return
	}
}


;----
;Joystick Test Script
;https://autohotkey.com/docs/scripts/JoystickTest.htm
;----

;----
;Detect if a variable is empty/blank - AutoHotkey Community
;https://autohotkey.com/boards/viewtopic.php?t=5080
;----
	

	
Return
;--------------------------------------------------------------------
; END OF TIMER_1
;--------------------------------------------------------------------

;----------------------------------------
TIMER_SUB_VICE_VERSA:

setTimer TIMER_SUB_VICE_VERSA, % -1 * 1000 * 60 * 60 ; After1Hours

VICE_VERSA_EXE=C:\Program Files\ViceVersa Pro\ViceVersa.exe
VICE_VERSA_PATH=C:\SCRIPTER\SCRIPTER -- ViceVersa PRO\# VV C DRIVE ROOT

; Run, "%VICE_VERSA_EXE%" "%VICE_VERSA_PATH%\VV C DRIVE ROOT __ %A_ComputerName%.fsf" /hiddenautoexec

; /hiddenautoexec
; /dialogautoexec /autoclose 

Return

;----------------------------------------
TIMER_SUB_CMD_KILL:

;setTimer TIMER_SUB_CMD_KILL, % -1 * 1000 * 60 * 60 ; After1Hours

If (A_Now>START_CMD_KILL)
{
	IF (OSVER_N_VAR = 5)  ; WIN XP
	{
		START_CMD_KILL=%A_Now%
		START_CMD_KILL+=10*60 ; 10 MINUTES

		;Run, "TASKKILL.exe" /F /IM CMD.EXE /T , , HIDE
	}
}
Return



;----------------------------------------
TIMER_SUB_SCRIPT_SHELL_FOLDERING:
; STARTS AS 1 SECOND AND THEN GOES TO EVERY HOUR

setTimer TIMER_SUB_SCRIPT_SHELL_FOLDERING,% -1 * 1000 * 60 * 60 ; After1Hours

; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 13-COPY MOVE SHELL FOLDING.VBS" , , hide

Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 10-VICEVERSA _ SHELL FOLDERING__.AHK" /RUN

; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 10-VICEVERSA _ SHELL FOLDERING__.AHK" , , hide
; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 10-VICEVERSA _ SHELL FOLDERING__.VBS

Return


;----------------------------------------
TIMER_SUB_LOGGER:

setTimer TIMER_SUB_LOGGER, % -1 * 1000 * 60 * 60 ; After24Hours

Run, "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 31-Duplicate Cleaner Logger.BAT" , , hide

Run, "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\BlueToothView_Desc.VBS" , , hide

Return

;----------------------------------------
TIMER_SUB_WINDOWS_UPDATE:

;----
;Downloads / Tools by other people / Windows 10 Update Disabler
;https://winaero.com/download.php?view.1932
;----

RETURN

IF (A_ComputerName="7-ASUS-GL522VW") 
	RETURN

setTimer TIMER_SUB_WINDOWS_UPDATE, % -1 * 1000 * 60 * 60 ; After24Hours

OSVER_N_RESULT=0
IF (OSVER_N_VAR = 5)  ; WIN XP
	OSVER_N_RESULT=1
IF (OSVER_N_VAR = 6)  ; WIN 07
	OSVER_N_RESULT=1
IF (OSVER_N_VAR < 10) ; WIN 10
	OSVER_N_RESULT=1  ; SET NOT RUN IF THIS LESS THAN WINDOWS 10 OS SYS
	
COMP_N=0
;IF (A_ComputerName="7-ASUS-GL522VW") 
;	COMP_N=1        ; SET NOT RUN IF THIS COMPUTER

ALLOW_UPDATE_KILLER=1
IF (OSVER_N_RESULT=1) or (COMP_N=1) 
	ALLOW_UPDATE_KILLER=0

If (ALLOW_UPDATE_KILLER=1)
{
	Run, "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 12-KEEP WINDOWS UPDATE STOPPED.BAT" /QUITE , , HIDE
}
Return


;----------------------------------------
TIMER_SUB_OWNER:
DetectHiddenWindows, on

;Process, Exist, ICACLS.EXE
;If ErrorLevel > 0
;	{
;	;MSGBOX IT DOES EXIST
;	}


IfWinExist TeamViewer Panel ahk_class TV_ControlWin
{
	Process, Exist, ICACLS.EXE
	If ErrorLevel > 0
	{
		Process, Close, ICACLS.EXE
		SoundBeep , 2000 , 100
		;MSGBOX HH
	}
}

IF (OSVER_N_VAR < 6 ) ; THAN XP
	RETURN
Process, Exist, ICACLS.EXE
If ErrorLevel > 0
	RETURN
		
IF TIMER_SUB_OWNER_SAVE_TIMER<%A_NOW%
{	
	TIMER_SUB_OWNER_SAVE_TIMER=%A_NOW%
	TIMER_SUB_OWNER_SAVE_TIMER+= 2, Days
	IF (A_ComputerName="3-LINDA-PC") 
		TIMER_SUB_OWNER_SAVE_TIMER+= 4, Days

	FileDELETE, %SCRIPT_NAME_VAR%
	FileAppend,%TIMER_SUB_OWNER_SAVE_TIMER%,%SCRIPT_NAME_VAR%

}
ELSE
{
	RETURN
}

	
SET_GO_RESULT=0

IF (OSVER_N_VAR > 5 ) ; BIGGER THAN XP
	SET_GO_RESULT=1

Process, Exist, ICACLS.EXE
If ErrorLevel > 0
	{
	SET_GO_RESULT=0
	}

If (SET_GO_RESULT=1)
{
	Run, "C:\SCRIPTER\SCRIPTER CODE -- BAT\OWNER\#_OWNER-HARDCODED ANYWHERE.BAT" /QUITE , , HIDE
}



;----------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return

ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
	return (match > 1)
	}
Return




;# ------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ------------------------------------------------------------------

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

;# ------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
    if ExitReason not in Logoff,Shutdown
    {
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
    }
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
        ;
        ;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}
;# ------------------------------------------------------------------
; exit the app


;GOOD SCRIPT EXAMPLE PAGE HELPER
;----
;1 Hour Software by Skrommel - DonationCoder.com
;http://www.donationcoder.com/Software/Skrommel/index.html#GoneIn60s
;----

;StringMid,now,A_Now,9,4


; REFERENCES 
; -------------------------------------------------------------------
; --------------------------------------------------------------
; [ Sunday 19:26:10 Pm_15 April 2018 ]
; THIS IS OUR PREFERRED DEFAULT OPTIONS FOR INSTALLING NOTEPAD++
; THE MIDDLE CHECKBOX IS SELECTED
; NONE OF OUR PLUGIN WILL WORK PROPER AS OUR SETUP USES THE
; Allow plugins to be loaded from %APPDATA%\notepad++ -- CHECKBOX CHECKED
; EVERY-TIME AN UPDATE COMES ALONG HAS TO BE REMEMBER
; OUR SETTER USES 32-BIT NOTEPAD FOR SOME REASON NOT UPGRADED YET
; ----
; AUTOHOTKEYS SELECT A CHECKBOX - Google Search
; https://www.google.co.uk/search?q=AUTOHOTKEYS+SELECT+A+CHECKBOX&oq=AUTOHOTKEYS+SELECT+A+CHECKBOX&aqs=chrome..69i57j0.11108j0j7&sourceid=chrome&ie=UTF-8
; --------
; Check a Checkbox - Control check..... - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/25075-check-a-checkbox-control-check/
; ----
; --------------------------------------------------------------
