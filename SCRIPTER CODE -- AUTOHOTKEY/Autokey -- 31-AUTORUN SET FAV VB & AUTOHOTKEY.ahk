;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 31-AUTORUN SET FAV VB & AUTOHOTKEY.ahk
;# __ 
;# __ Autokey -- 31-AUTORUN SET FAV VB & AUTOHOTKEY.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Sun 27-May-2018 04:48:00 ]
;# END         TIME [ Sun 27-May-2018 04:58:00 ] 10 MINUTE
;# LAST EDITOR TIME [ Sun 27-May-2018 04:58:00 ]
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; 001 ---------------------------------------------------------------
; AUTOEXEC BOOT SCRIPT
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Sun 27-May-2018 04:48:00
; TO   TIME __ Sun 27-May-2018 04:48:00 10 MINUTE
; -------------------------------------------------------------------
; 002 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Wed 30-May-2018 07:47:13
; TO   TIME __ Wed 30-May-2018 08:10:00 MAJOR CHANGE USE COMMAND LINE PARAM SAVE DUPLICATE CODE
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location Internet
;# ------------------------------------------------------------------
; Link to Folder of My AutoHotKeys Project Set
; https://drive.google.com/open?id=0BwoB_cPOibCPVmVYT1pKWUk4LVE
;# ------------------------------------------------------------------
; Link to Folder of My AutoHotKeys Project Set Dropbox
; https://www.dropbox.com/sh/h2ebk12dksaq7j3/AAD9Ow_SbBP33JKmuALRkO1_a?dl=0
;# ------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
; https://www.dropbox.com/s/w22p90h81uwni71/Autokey%20--%2021-AUTORUN.ahk?dl=0
;# ------------------------------------------------------------------


; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force

DetectHiddenWindows, on
SetStoreCapslockMode, off

GOSUB MAIN_ROUTINE

; -------------------------------------------------------------------

Exitapp

RETURN



MAIN_ROUTINE:






; -------------------------------------------------------------------
; -------------------------------------------------------------------
Element_1 := "D:\VB6\VB-NT\00_Best_VB_01\TIMEZONE MINI GUI DISPLAY\TIMEZONE MINI GUI DISPLAY.exe"
if FileExist(Element_1)
{
	; SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
; ;	Run, %Element_1%
;	SLEEP 3000
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------



; -------------------------------------------------------------------
; -------------------------------------------------------------------
Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk"
if FileExist(Element_1)
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run, %Element_1%
	SLEEP 3000
}


\\4-ASUS-GL522VW\4_ASUS_GL522VW_E_DRIVE:
; -------------------------------------------------------------------
; -------------------------------------------------------------------
Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN SECURE.ahk"
if FileExist(Element_1)
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run, %Element_1%
}
; -------------------------------------------------------------------
; -------------------------------------------------------------------


Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 79-BROWSER LOAD URL BOOT CHROME.ahk"
if FileExist(Element_1)
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run, %Element_1%
	SLEEP 3000
}

SET_GO_1=
IF (A_ComputerName="4-ASUS-GL522VW") 
	SET_GO_1=1
; IF (A_ComputerName="8-MSI-GP62M-7RD") 
	; SET_GO_1=1
IF (A_ComputerName="9-ASUS-G815LM") 
	SET_GO_1=1
IF SET_GO_1=1
{
	Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK"
	if FileExist(Element_1)
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		Run, %Element_1%
	}
}


SLEEP 3000

SET_GO_1=
IF (A_ComputerName="4-ASUS-GL522VW")
	SET_GO_1=1
IF (A_ComputerName="8-MSI-GP62M-7RD") 
	SET_GO_1=1
IF (A_ComputerName="9-ASUS-G815LM") 
	SET_GO_1=1
	
SET_GO_1=1
	
IF SET_GO_1=1
{

	Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk"
	if FileExist(Element_1)
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		Run, %Element_1%
	}
	; -------------------------------------------------------------------
	; -------------------------------------------------------------------

	Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk"
	if FileExist(Element_1)
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		Run, %Element_1%
		SLEEP 3000
	}
}
SLEEP 3000


SET_GO_1=
IF (A_ComputerName = "4-ASUS-GL522VW") 
	SET_GO_1=1
IF (A_ComputerName = "8-MSI-GP62M-7RD") 
	SET_GO_1=1
IF (A_ComputerName = "9-ASUS-G815LM") 
	SET_GO_1=1

SET_GO_1=1
	
IF SET_GO_1=1
{
	Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk"
	if FileExist(Element_1)
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		; MSGBOX "HERE"
		Run, %Element_1%
		SLEEP 3000
	}
}



SET_GO_1=
IF (A_ComputerName = "4-ASUS-GL522VW") 
	SET_GO_1=TRUE
IF (A_ComputerName = "8-MSI-GP62M-7RD") 
	SET_GO_1=TRUE
IF (A_ComputerName = "9-ASUS-G815LM") 
	SET_GO_1=TRUE
	
IF SET_GO_1
{	
	Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 58-AUTO REPEAT BROWSER FUNCTION SET.ahk"
	if FileExist(Element_1)
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		Run, %Element_1%
		SLEEP 3000
	}
}

SET_GO_1=TRUE

IF SET_GO_1
{	
	Element_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 75-GOODSYNC OPTIONS SET.ahk"
	if FileExist(Element_1)
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		Run, %Element_1%
		SLEEP 3000
	}
}



;--------------------------------------------------------------------
;AUTOHOTKEYS
;--------------------------------------------------------------------

FN_VAR_1 := "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
TEMP_VAR_1=%FN_VAR_1%
TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")
FN_VAR_1=%TEMP_VAR_3%

DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY
IFWINNOTEXIST %FN_VAR_1%
{
	; SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	; Run, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk
}


FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 04-SET CAP NUM & SCROLL TO LIKING - ONCE.ahk"	
if FileExist(FN_VAR)
	IF !WinExist(FN_VAR) 
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		Run, "%FN_VAR%"
		SLEEP 3000
	}

; -------------------------------------------------------------------
; ALL THE AUTOHOTKEYS DONE
; -------------------------------------------------------------------






; DetectHiddenWindows On
; FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 21-AutoRun.ahk"
; SplitPath, FN_VAR, name
; IFWINNOTEXIST, %NAME%
; IfExist, %FN_VAR%
; IF !WinExist(NAME) 
; {
	; ; MSGBOX %NAME% 
	; SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	; Run, "%FN_VAR%" /RUN_FAVS
; }
; PAUSE 
; -------------------------------------------------------------------

SetTitleMatchMode 3


FN_VAR:="C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2026.520.400_x64__8wekyb3d8bbwe\olk.exe"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run "%FN_VAR%"
	SLEEP 2000
	WinWait, ahk_class Outlook Host, , 15
	SLEEP 2000
	WinWait, ahk_class Chrome_WidgetWin_0, , 15
	SLEEP 2000
	WinWait, ahk_class Outlook Host, , 15
	SLEEP 2000
	; WinMinimize
	; SLEEP 4000
	; -----------------------------
	; ---- THIS ONE DOES HER
	; ---- WELL, IT DID AND THEN IT STOPPED
	; ---- TALK ABOUT THIS ONE LATER ---- 2026 MAY 
	; ---- FIXXER AGAIN -- RUNWAIT WAS WRONG -- RUN AND WAIT FOR EXIT DOZY
	;-----RUNNER MINIMIZE EXCEPT THE EXE -- MUCKS IT UP 
	; -----------------------------
	WinMinimize, ahk_exe %name%
	

}



FN_VAR:="D:\VB6\VB-NT\00_BEST_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run, "%FN_VAR%" ,, MIN
	WinWait, VB_KEEP_RUNNER, , 10
	WinMinimize
}


FN_VAR:="D:\VB6\VB-NT\00_BEST_VB_01\CLIPBOARD_VIEWER\ClipBoard Viewer.exe"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run, "%FN_VAR%" ,, MIN
	WinWait, Clipboard Viewer, , 10
	WinMinimize
}


FN_VAR:="D:\VB6\VB-NT\00_BEST_VB_01\URL Logger\URL Logger.exe"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run, "%FN_VAR%" ,, MIN
	WinWait, URL Logger, , 10
	WinMinimize
}


FN_VAR:="D:\VB6\VB-NT\00_BEST_VB_01\CLIPBOARD LOGGER\CLIPBOARD LOGGER.EXE"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run, "%FN_VAR%" ,, MIN
	WinWait, ClipBoard Logger, , 10
	WinMinimize
}



SET_GO_1=TRUE
IF (A_ComputerName = "2-ASUS-EEE") 
	SET_GO_1=
	
IF SET_GO_1
{	
	SetTitleMatchMode 2 
	FN_VAR:="C:\Program Files\Google\Chrome\Application\chrome.exe"
	IfNOTExist, %FN_VAR%
	FN_VAR:="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
	SplitPath, FN_VAR, name
	Process, Exist, %name%
	If Not ErrorLevel
	IfExist, %FN_VAR%
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		Run, "%FN_VAR%"
		SLEEP 4000
		WinWait, ahk_class Chrome_WidgetWin_1, , 10
		WinMinimize
		COUNTER_LOOP=0
		LOOP 
		{
		WinGet, HWND_10, ID, ahk_class Chrome_WidgetWin_1
		WinMinimize ahk_id %HWND_10%
		WinGet, state, MinMax, ahk_id %HWND_10%
		if (state = -1)
			BREAK
		COUNTER_LOOP+=1
		IF COUNTER_LOOP>100
			BREAK 
		SLEEP 1000
		}
	}
}


SET_GO_1=TRUE
IF (A_ComputerName = "2-ASUS-EEE") 
	SET_GO_1=
	
IF SET_GO_1
{	
	FN_VAR:="C:\Program Files\Mozilla Firefox\firefox.exe"
	SplitPath, FN_VAR, name
	Process, Exist, %name%
	If Not ErrorLevel
	IfExist, %FN_VAR%
	{
		SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		Run "%FN_VAR%"
		SLEEP 2000
		WinWait, ahk_class MozillaWindowClass, , 10
		COUNTER_LOOP_1=0
		COUNTER_LOOP_2=0
		LOOP 
			{
			WinGet, HWND_10, ID, ahk_class MozillaWindowClass
			WinMinimize ahk_id %HWND_10%
			WinGet, state, MinMax, ahk_id %HWND_10%
			if (state = -1)
				BREAK
			COUNTER_LOOP_1+=1
			IF COUNTER_LOOP_1>100
				BREAK 
			SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
			SLEEP 1000
			}
	}
}
; PAUSE




FN_VAR:="C:\Program Files (x86)\Siber Systems\AI RoboForm\identities.exe"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run "%FN_VAR%"
	WinWait, ahk_class RfEditor, , 10
	COUNTER_LOOP=0
	LOOP 
	{
	WinGet, HWND_10, ID, ahk_class RfEditor
	WinMinimize ahk_id %HWND_10%
	WinGet, state, MinMax, ahk_id %HWND_10%
	if (state = -1)
		BREAK
	COUNTER_LOOP+=1
	IF COUNTER_LOOP>100
		BREAK 
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	SLEEP 1000
	}
}

FN_VAR:="C:\Program Files\Siber Systems\GoodSync\GoodSync.exe"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run "%FN_VAR%"
	WinWait, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}, , 10
	COUNTER_LOOP=0
	LOOP 
	{
	WinGet, HWND_10, ID, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
	WinMinimize ahk_id %HWND_10%
	WinGet, state, MinMax, ahk_id %HWND_10%
	if (state = -1)
		BREAK
	COUNTER_LOOP+=1
	IF COUNTER_LOOP>100
		BREAK 
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	SLEEP 1000
	}
}

FN_VAR:="C:\Program Files (x86)\PicPick\picpick.exe"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run "%FN_VAR%"  /startup
	; WinWait, ahk_class TfrmOptions, , 5
	; WinMinimize
}

; ---- THIS ONE TAKER LONGER SAVE TILL LAST
; --------------------------------------------------------------------------
SetTitleMatchMode 2
FN_VAR:="C:\Program Files (x86)\Notepad++\notepad++.exe"
SplitPath, FN_VAR, name
Process, Exist, %name%
If Not ErrorLevel
IfExist, %FN_VAR%
{
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	Run "%FN_VAR%"
	WinWait, ahk_class Notepad++, 10
	COUNTER_LOOP=0
	LOOP
	{
	WinGet, HWND_10, ID, ahk_class Notepad++
	WinMinimize ahk_id %HWND_10% ; ---- RESULT GOOD ONE
	; WinMinimize Notepad++; ------------ TRY TITLE
	; WinMinimize, ahk_exe %name%  ; ---- THIS NOT WORK FOR OUR NOTEPAD++
	WinGet, state, MinMax, ahk_id %HWND_10%
	if (state = -1)
		BREAK
	COUNTER_LOOP+=1
	IF COUNTER_LOOP>500
		BREAK
	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	SLEEP 1000
	}
}

RETURN


