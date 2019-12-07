;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK
;# __ 
;# __ Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

;---------------------------------------------------------------
; DESCRIPTION
;---------------------------------------------------------------
; DETECT CAMERA BY VOLUME LABEL OF MEDIA CARD INSERTED
; AND THEN RUNS ACCORDING PROGRAM
; C:\SCRIPTER\SCRIPTER CODE -- VBS\
; VBS 29-COPY CAMERA PHOTO IMAGES.VBS
; AND THE VBS SCRIPTOR PROGRAM THEN GOES ON TO RUN OTHER PROGRAMS IN VB6-VBA 
; FOR RE-NAMER PHOTOS WITH EXIF INFO AND DIRECTORY'S INTO DAYS
;---------------------------------------------------------------

;# ------------------------------------------------------------------
; Location Internet
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
; 
;# ------------------------------------------------------------------

; 001
; --------------------------------------------------------------
; FROM  Thu 05-Jul-2018 08:45:29
; TO    Thu 05-Jul-2018 08:45:29
; --------------------------------------------------------------
; BEGIN CODE EXTRACTED FROM 
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
; TO ENABLE SHARE EASIER
; 
; --------------------------------------------------------------

#Warn
#NoEnv
#SingleInstance Force
;--------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; SCRIPT ============================================================

; ---------------------------------------------------------------
; I MADE MENU ITEM INTO INCLUDE FILE IN 3 PART 
; 01. INTRO SETUP MENU
; 02. THE MENU ROUTINE
; 03. ANY ROUTINE THE MENU USE
; ---------------------------------------------------------------
; SAVER OF RSI INJURY AND MORE ACCURATE
; THE INCLUDE FILE ARE SAME FOLDER
; ---------------------------------------------------------------
; FROM __ Sun 09-Jun-2019 07:03:00 __ Clipboard Count = 024
; TO   __ Sun 09-Jun-2019 17:48:00 __ Clipboard Count = 452 __ 10 HOURING 45 MINUTE
; ---------------------------------------------------------------
; Create the popup menu by adding some items to it.
; MenuHandler:
; ---------------------------------------------------------------
; #Include GO WITH FULL PATH AS SOME LAUNCHER DO NOT SET WORK PATH WHEN RUNNER
; RATHER THAN CHANGE THE WORKING PATH WITHIN-AH
; ---------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01-INCLUDE MENU 01 of 03.ahk

VAR_STORE_CAMERA_LABEL=

OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

IF OSVER_N_VAR<10
{
	Exitapp
}

GLOBAL O_HWND_ID_1
GLOBAL MSGBOX_CAMERA_CountDown_Generate_VAR

setTimer TIMER_PREVIOUS_INSTANCE,1
	
SETTIMER TIMER_DRIVE_GET_CAMERA,4000

RETURN

TIMER_DRIVE_GET_CAMERA:
DriveGet, list, list
COUNT_VAR_1:=StrLen(LIST)-1
COUNT_VAR_2=0
LOOP
{
	COUNT_VAR_2+=1
	VOLUME_LETTER:=SUBSTR(LIST,COUNT_VAR_2,1)":\"
	; VOLUME_LETTER:=%VOLUME_LETTER% 
	
	DriveGet, Label_CAMERA, label, %VOLUME_LETTER%
	
	;MSGBOX % label
	
	SET_GO=FALSE
	IF Label_CAMERA=MC - HX60V
		SET_GO=TRUE
	IF Label_CAMERA=MC - HX60V0
		SET_GO=TRUE
	IF Label_CAMERA=FUJI XP90
		SET_GO=TRUE
	IF Label_CAMERA=GOPRO
		SET_GO=TRUE
	IF Label_CAMERA=DUAL SCRN
		SET_GO=TRUE
	IF Label_CAMERA=DASH_CAMERA
		SET_GO=TRUE
	IF Label_CAMERA=4K-ULTRA-HD
		SET_GO=TRUE
	IF Label_CAMERA=4K-ULTRA-02
		SET_GO=TRUE
	IF Label_CAMERA=KEY_FOB
		SET_GO=TRUE
		
	IF SET_GO=TRUE
		BREAK

	IF COUNT_VAR_2>%COUNT_VAR_1%
		BREAK
}

; -------------------------------------------------------------------
; LOCK IT WORKER REMOVE CAMERA BEFORE DO ANOTHER
; -------------------------------------------------------------------
IF SET_GO=FALSE
	VAR_STORE_CAMERA_LABEL=

IF (!VAR_STORE_CAMERA_LABEL and SET_GO="TRUE")
{
	VAR_STORE_CAMERA_LABEL:=Label_CAMERA
	Secs_CAMERA := 800
	SetTimer, MSGBOX_CAMERA_CountDown, 1000
	GOSUB MSGBOX_CAMERA_CountDown_Generate_Line
	MsgBox, 4, Copy_Camera_Media_Card, %MSGBOX_CAMERA_CountDown_Generate_VAR%, %Secs_CAMERA%
	SetTimer, MSGBOX_CAMERA_CountDown, Off
	IfMsgBox Yes
	{
		VAR_STORE_CAMERA_LABEL:=Label_CAMERA
		FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 29-COPY CAMERA PHOTO IMAGES.VBS"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2000 , 100
			SoundBeep , 3000 , 100
			Run, "%FN_VAR%"
		}
	}
}


SetTitleMatchMode 3  ; Exactly
DetectHiddenText, Off

HWND_ID_1 := WinExist("This PC\FinePix XP90 XP91 XP95 ahk_class CabinetWClass")
IF (HWND_ID_1>0 and HWND_ID_1<>O_HWND_ID_1)
{
	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA - FUJI XP90 - GOODSYNC.BAT"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%"
	}
}
O_HWND_ID_1=%HWND_ID_1%






RETURN

MSGBOX_CAMERA_CountDown_Generate_Line:
MSGBOX_CAMERA_CountDown_Generate_VAR=Would you like to Copy Camera Media Card`n%Label_CAMERA%`nRespond in`n%Secs_CAMERA% Seconds
RETURN

MSGBOX_CAMERA_CountDown:
Secs_CAMERA -= 1
GOSUB MSGBOX_CAMERA_CountDown_Generate_Line
ControlSetText,Static1,%MSGBOX_CAMERA_CountDown_Generate_VAR%,Copy_Camera_Media_Card ahk_class #32770

Return

MenuHandler:
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-02-INCLUDE MENU 02 of 03.ahk
return

#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03-INCLUDE MENU 03 of 03.ahk



;# ------------------------------------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return

ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	DHW_2 := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % DHW_2
	return (match > 1)
	}
Return


; ------------------------------------------------------------------
EOF:                           ; on exit
Display.OnExit()               ; release and free the library
ExitApp     
; ------------------------------------------------------------------

; ------------------------------------------------------------------
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
; ---------------------------------------------------------------------
; exit the app
