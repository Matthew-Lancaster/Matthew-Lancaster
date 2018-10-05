;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
;# __ 
;# __ Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Fri 17:20:00 Pm_04 May 2018 ]
;# END         TIME [ Fri 17:40:00 Pm_04 May 2018 ]
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; 001 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Fri 04-May-2018 17:20:00
; TO   TIME __ Fri 04-May-2018 17:40:00 BEGINNER
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 002 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Thu 07-Jun-2018 16:21:43
; TO   TIME __ Thu 07-Jun-2018 16:27:00 ADD AHK SCRIPT ALREADY RUNNING
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 003 ---------------------------------------------------------------
; -------------------------------------------------------------------
; NOW WITH MORE INTELLIGENT MODIFIED DATE RELOADER 
; AND USER OF ARRAY'S
; AND RENAME FROM 
; Autokey -- 28-Autohotkeys Set AutoRun.ahk
; TO
; Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
; -------------------------------------------------------------------
; FROM TIME __ Wed 20-Jun-2018 21:41:37
; TO   TIME __ Wed 20-Jun-2018 22:28:00 _ 50 MINUTE LEARN
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
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
; 
;# ------------------------------------------------------------------


; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
;--------------------

SetStoreCapslockMode, off

DetectHiddenWindows, on
SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

;--------------------------------------------------------------------
;AUTOHOTKEYS
;--------------------------------------------------------------------

; Each array must be initialized before use:
FN_Array := []
DATE_MOD_Array := []

ArrayCount := 0
ArrayCount += 1
FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk"	
ArrayCount += 1
FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 02-SAVE AS KEY ENTER.ahk"
ArrayCount += 1
FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk"
ArrayCount += 1
FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 14-Brightness With Dimmer.ahk"
ArrayCount += 1
FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES MYSMS & FB GB.ahk"
ArrayCount += 1
FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 19-SCRIPT_TIMER_UTIL.ahk"
ArrayCount += 1
FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 32-BRUTE BOOT DOWN.ahk"

OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

IF OSVER_N_VAR>=10
{
	ArrayCount += 1
	FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK"
}

ArrayCount += 1
FN_Array[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"


; ArrayCount += 1
; FN_Array[ArrayCount] := 
; ArrayCount += 1
; FN_Array[ArrayCount] := 


Loop % ArrayCount
{
  Element := FN_Array[A_Index]

	IfExist, %Element%
	{
		FileGetTime, OutputVar, %Element%, M
		DATE_MOD_Array[A_Index] := OutputVar
		; MSGBOX % DATE_MOD_Array[A_Index]
	}
}

setTimer TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD,4000

RETURN

TIMER_SUB_AUTOHOTKEYS_ARRAY_RELOAD:

Loop % ArrayCount
{
  Element_1 := FN_Array[A_Index]
  Element_2 := DATE_MOD_Array[A_Index]

	IfExist, %Element_1%
	SET_GO=FALSE
	IF !WinExist(Element_1) 
		SET_GO=TRUE
	
	FileGetTime, OutputVar, %Element_1%, M
	IF OutputVar<>%Element_2%
	{
		SET_GO=TRUE
		DATE_MOD_Array[A_Index] := OutputVar
	}
		
	IF SET_GO=TRUE	
		{
			SoundBeep , 2000 , 100
			Run, "%Element_1%"
		}
}

RETURN
