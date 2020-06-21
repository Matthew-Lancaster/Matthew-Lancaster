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
<<<<<<< HEAD
; C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\
=======
; C:\SCRIPTER\SCRIPTER CODE -- VBS\
>>>>>>> 90ddd3dff472788f77096141bd1af892a766f0ff
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
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk

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

SetTitleMatchMode 2 
DetectHiddenText, Off


<<<<<<< HEAD
LABEL_CAMERA=

=======
>>>>>>> 90ddd3dff472788f77096141bd1af892a766f0ff

GLOBAL O_HWND_ID_1
GLOBAL MSGBOX_CAMERA_CountDown_Generate_VAR

SETTIMER TIMER_PREVIOUS_INSTANCE,1
	
TIMER_DRIVE_GET_CAMERA_FIRST_RUN=TRUE
SETTIMER TIMER_DRIVE_GET_CAMERA,1

SETTIMER TIMER_CHKDSK_OF_V_DRIVE_CAMERA_MEDIA_CARD_DELETE_ITEM,5000

SETTIMER TIMER_DELETE_EMPTY_2020_CyberShot_HX60V_MP_ROOT,10000

RETURN

; ---------------------------------------------------------------
; MODIFY MIDNIGHT TO 3 HOUR
; ---------------------------------------------------------------
; ROUTINE ADDITION HERE 
; DELETE CHKDSK FRAGMENT
; HERE CODE HALF HOUR
; TIMER_CHKDSK_OF_V_DRIVE_CAMERA_MEDIA_CARD_DELETE_ITEM:
; Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK
; TWO CODE LINKER
; Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:
; Fri 17-Jan-2020 14:40:00
; ---------------------------------------------------------------

TIMER_CHKDSK_OF_V_DRIVE_CAMERA_MEDIA_CARD_DELETE_ITEM:
	SETTIMER TIMER_CHKDSK_OF_V_DRIVE_CAMERA_MEDIA_CARD_DELETE_ITEM,300000

	VOLUME_LETTER:="V:\"
	
	DriveGet, LABEL_CAMERA, label, %VOLUME_LETTER%

	IF LABEL_CAMERA=MC - HX60V1
	{
		FN_VAR_DELETE:="V:\FOUND.*"    ; EXAMPLE FOUND.004
		IfExist, %FN_VAR_DELETE%
		{
			Loop, Files, %FN_VAR_DELETE% , D      ; 2 = DIRECTORY
			{
				FOLDER_NAME_CHKDSK = %A_LoopFileFullPath%
			}
			
			FileRemoveDir %FOLDER_NAME_CHKDSK%, 1
		}
	}
RETURN

; --------------------------------------------------------------------
; ---- CODE TIME WRITE -- Fri 17-Jan-2020 15:00:00
; ---- CODE TIME WRITE -- Fri 17-Jan-2020 15:20:00
; ---- CHECKER EVERY POSSIBILITY
; ---- MOST SUITABLE WRITE STYLE
; --------------------------------------------------------------------
TIMER_DELETE_EMPTY_2020_CyberShot_HX60V_MP_ROOT:

	SETTIMER TIMER_DELETE_EMPTY_2020_CyberShot_HX60V_MP_ROOT,10000 ; 10 SEC

	FILE_NAME_COUNTER=
	FILE_NAME_COUNT=
	
	FN_VAR_FILE_EXIST_AR_1:="D:\DSC\2015+SONY\2020 CyberShot HX60V\MP_ROOT"
	FN_VAR_FILE_EXIST_AR_2=%FN_VAR_FILE_EXIST_AR_1%\*.*
	IfExist, %FN_VAR_FILE_EXIST_AR_2%
	{
		Loop, Files, %FN_VAR_FILE_EXIST_AR_2%
		{
			FILE_NAME_COUNTER = %A_LoopFileFullPath%
			FILE_NAME_COUNT+=1
		}
	
		IF !FILE_NAME_COUNT
		{
			FileRemoveDir %FN_VAR_FILE_EXIST_AR_1% ,1
		}
	}

RETURN


TIMER_DRIVE_GET_CAMERA:
	
	IF TIMER_DRIVE_GET_CAMERA_FIRST_RUN
		SETTIMER TIMER_DRIVE_GET_CAMERA,10000
	IF TIMER_DRIVE_GET_CAMERA_FIRST_RUN
		TIMER_DRIVE_GET_CAMERA_FIRST_RUN=

		
	
	
	SET_ARRAY_1:=[]
	I:=0
	I+=1,SET_ARRAY_1[I]:="MC - HX60V"
	I+=1,SET_ARRAY_1[I]:="MC - HX60V0"
	I+=1,SET_ARRAY_1[I]:="FUJI XP90"
	I+=1,SET_ARRAY_1[I]:="GOPRO"
	I+=1,SET_ARRAY_1[I]:="DUAL SCRN"
	I+=1,SET_ARRAY_1[I]:="DASH_CAMERA"
	I+=1,SET_ARRAY_1[I]:="4K-ULTRA-HD"
	I+=1,SET_ARRAY_1[I]:="4K-ULTRA-02"
	I+=1,SET_ARRAY_1[I]:="KEY_FOB"
	I+=1,SET_ARRAY_1[I]:="MEM-BANK-AR"
	I+=1,SET_ARRAY_1[I]:="MEM-BANK-CAMERA"      ; -- WHICH ONE CODE AFTER USER HERE
	I+=1,SET_ARRAY_1[I]:="HIKVISION#1"
	I+=1,SET_ARRAY_1[I]:="MAGICMINDR5"
	I+=1,SET_ARRAY_1[I]:="MAGICMINDR7"
	I+=1,SET_ARRAY_1[I]:="FRONT DOOR"
	I+=1,SET_ARRAY_1[I]:="HI-LOOK-04"           ; WHITE 04 PORT
	I+=1,SET_ARRAY_1[I]:="HI-LOOK-08"           ; BLACK 08 PORT
	I+=1,SET_ARRAY_1[I]:="HARD_REC_8G"          ; ACTIONS HS USB FlashDisk USB ---- FAT32 -- Wed 04-Mar-2020
	I+=1,SET_ARRAY_1[I]:="RETEKESS"
<<<<<<< HEAD
	I+=1,SET_ARRAY_1[I]:="128GB-V-OPA"          ; ---------------------------------------------------------------
												; £54.99 HTTPS://TINYURL.COM/Y9CN3UG3
												; Spy Camera, OMOUP 10000mAh Power Bank with Camera,
												; 1080P HD Hidden Camera Nanny Cam with Motion Detection 
												; Night Vision Gravity Sensing Function
												; 17-JUN-2020 ORDER -- 18-JUN-2020 DELIVERY -- 20-JUN-2020 GET GO
												; NEW MODEL -- HAS EVERYTHING
												; AND OVER SPEC TAKE 128GB MC-CARD -- THAN SPECIFY MAXIMUM 64GB
												; ---------------------------------------------------------------
	I+=1,SET_ARRAY_1[I]:="HARD-RIPER2"          ; ACTIONS HS USB FlashDisk USB ---- FAT32 -- Sat 20-Jun-2020
												; Digital Voice Recorder, COOAU 8GB Rechargeable: Amazon.co.uk: Electronics 
												; https://www.amazon.co.uk/gp/product/B07G5BVWS8/ref=ppx_yo_dt_b_asin_title_o00_s00?ie=UTF8&psc=1
												; https://tinyurl.com/yanwzeyu
												; ---------------------------------------------------------------
	; I+=1,SET_ARRAY_1[I]:=""
	; I+=1,SET_ARRAY_1[I]:=""
	; I+=1,SET_ARRAY_1[I]:=""
	; I+=1,SET_ARRAY_1[I]:=""
	; I+=1,SET_ARRAY_1[I]:=""
=======
>>>>>>> 90ddd3dff472788f77096141bd1af892a766f0ff
	; I+=1,SET_ARRAY_1[I]:=""
	; I+=1,SET_ARRAY_1[I]:=""
	; I+=1,SET_ARRAY_1[I]:=""

	
	FIND_VOLUME_NAME_02=
	
	DriveGet, list, list
	COUNT_VAR_1:=StrLen(LIST)-1
	COUNT_VAR_2=0
	DRIVE_SET=
	LOOP, % StrLen(LIST)
	{
		VOLUME_LETTER:=SUBSTR(LIST,A_INDEX,1)
		DriveGet, LABEL_CAMERA, label, %VOLUME_LETTER%:
		DRIVE_SET=%DRIVE_SET%--%LABEL_CAMERA%__
	}
	SET_GO=
	Loop, % SET_ARRAY_1.MaxIndex()
	{
		FIND_VOLUME_NAME_01:="--" SET_ARRAY_1[A_Index] "__"
		FIND_VOLUME_NAME_02:=SET_ARRAY_1[A_Index]
		IF INSTR(DRIVE_SET,FIND_VOLUME_NAME_01)>0
<<<<<<< HEAD
		{
			LABEL_CAMERA:=SET_ARRAY_1[A_Index]
			SET_GO=TRUE
			BREAK                                ; -- FIRST FIND GIVE EXIT
		}
	}

=======
			SET_GO=TRUE
		IF SET_GO=TRUE
			BREAK                                ; -- FIRST FIND GIVE EXIT
	}


>>>>>>> 90ddd3dff472788f77096141bd1af892a766f0ff
	; -------------------------------------------------------------------
	; IF DOS SHELL BATCH RUN SWERVE THAT ONE
	; AS TRIGGER FROM CERTAIN SYNC JOB
	; Fri 31-Jan-2020 04:25:06
	; -------------------------------------------------------------------
	; HARD MAKE WORK WHEN SOMETIME RUN REQUEST IS SILENTLY
	; -------------------------------------------------------------------
	WINFIND_VAR:="BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT"
	IFWINEXIST %WINFIND_VAR% ahk_class ConsoleWindowClass
	{
		; MSGBOX BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT `nIS ALREADY RUNNER
		RETURN
	}
	; -------------------------------------------------------------------

	
	; MSGBOX % SET_GO
	; MSGBOX % VAR_STORE_CAMERA_LABEL
	; MSGBOX % FIND_VOLUME_NAME_02
	; PAUSE
	
	IF VAR_STORE_CAMERA_LABEL
	IF !SET_GO
		SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
	IF !SET_GO
		VAR_STORE_CAMERA_LABEL=
	; -------------------------------------------------------------------
	; LOCK IT WORKER REMOVE CAMERA BEFORE DO ANOTHER
	; -------------------------------------------------------------------
	IF SET_GO
	IF FIND_VOLUME_NAME_02
	IF VAR_STORE_CAMERA_LABEL<>%FIND_VOLUME_NAME_02%
	{
		VAR_STORE_CAMERA_LABEL=%FIND_VOLUME_NAME_02%
		SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
		VAR_STORE_CAMERA_LABEL:=FIND_VOLUME_NAME_02   ; -- STOP REPEAT ASK IT MSGBOX 
		
		Secs_CAMERA := 800
		SetTimer, MSGBOX_CAMERA_CountDown, 1000
		GOSUB MSGBOX_CAMERA_CountDown_Generate_Line

		MsgBox, 4, Copy_Camera_Media_Card, %MSGBOX_CAMERA_CountDown_Generate_VAR%, %Secs_CAMERA%
		SetTimer, MSGBOX_CAMERA_CountDown, Off

		IfMsgBox Yes
		{
<<<<<<< HEAD
			FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 29-COPY CAMERA PHOTO IMAGES.VBS"
=======
			FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 29-COPY CAMERA PHOTO IMAGES.VBS"
>>>>>>> 90ddd3dff472788f77096141bd1af892a766f0ff
			IfExist, %FN_VAR%
			{

				SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
				Run, "%FN_VAR%"
			}
		}
		SETTIMER TIMER_DRIVE_GET_CAMERA,OFF
		SETTIMER TIMER_DRIVE_GET_CAMERA,10000
	}

	SetTitleMatchMode 3    ; EXACT AR
	DetectHiddenText, Off

	HWND_ID_1 := WinExist("This PC\FinePix XP90 XP91 XP95 ahk_class CabinetWClass")
	IF (HWND_ID_1>0 and HWND_ID_1<>O_HWND_ID_1)
	{
		FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA - FUJI XP90 - GOODSYNC.BAT"
		IfExist, %FN_VAR%
		{
			SOUNDPLAY, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV
			; SoundBeep , 2500 , 100
			Run, "%FN_VAR%"
		}
	}
	O_HWND_ID_1=%HWND_ID_1%
RETURN


MSGBOX_CAMERA_CountDown_Generate_Line:
<<<<<<< HEAD
	; DO YOU WANT
	; SLURP WANT
=======
>>>>>>> 90ddd3dff472788f77096141bd1af892a766f0ff
	MSGBOX_CAMERA_CountDown_Generate_VAR=Would you like to Copy Camera Media Card`n%LABEL_CAMERA%`nRespond in`n%Secs_CAMERA% Seconds
RETURN

MSGBOX_CAMERA_CountDown:
	Secs_CAMERA -= 1
	GOSUB MSGBOX_CAMERA_CountDown_Generate_Line
	ControlSetText,Static1,%MSGBOX_CAMERA_CountDown_Generate_VAR%,Copy_Camera_Media_Card ahk_class #32770
Return




#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk



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
