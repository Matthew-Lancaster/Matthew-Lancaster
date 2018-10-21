;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
;# __ 
;# __ Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Fri 17:20:00 Pm_04 May 2018 ]
;# END         TIME [ Fri 17:40:00 Pm_04 May 2018 ]
;# LAST EDITOR TIME [ Sun 14:05:00 Pm_21 Oct 2018 ]
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


; -------------------------------------------------------------------
; 004 ---------------------------------------------------------------
; -------------------------------------------------------------------
; ADD CODE - MORE PRECISE USING EXACT - SetTitleMatchMode 3 ; EXACTING
; DEPENDENT ON VERSION NUMBER OF AHK IS SAME FOREVER UNLESS MODIFY
;
; ADD CODE - BETTER USAGE OF ARRAY - ARRAY HANDLING TOOK ALL THE TIME
; LEARN LOTS MORE LIKE REMOVE AND INSERT UNABLE TO CHANGE AN 
; ARRAY VALUE -- SEEM AN EFFORT AS HAVE TO PUSH WHOLE STACK 
; UP AND DOWN AGAIN _ AT LEAST IT AUTO
; YOU CHANGE CHANGE TH VLAU ETO NOTHING SIMPLE ENOUGH
; BUT ALSO PROBLEM AROSE THAT A SECOND ARRAY I WANTER REQUIRED THAT 
; ALL THE ITEM WERE INSERTED IN WITH NEW CODE DISCOVERY 
; DURING A LOOP LEARN
; " - AutoHotkey v1.1.26.01"
; YOU COULD TRIM VERSION NUMBER DOWN A BIT AND NOT USE EXACTING AGAIN
;
; WELL IT ONLY TOOK 10 MINUTE TO LOCATE VERSION NUMBER AND USE IT ON
;
; -------------------------------------------------------------------
; FROM TIME __ Sat 20-Oct-2018 01:58:30
; TO   TIME __ Sat 20-Oct-2018 04:30:00 _ 2 HOUR 30 MINUTE
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 005 ---------------------------------------------------------------
; -------------------------------------------------------------------
; ADD CODE _ MADE A CHANGE _ IF YOU KILL A PROCESS IT WON'T RELOAD 
; BECAUSE THE APP FILE PROGRAM WILL ALREADY BEEN SCANNED IN
; IT ONLY LOOKING FOR DATE CHANGER NOT IF IT DISAPPEARED
; THEN IF WORKING ON SOMETHING APP AHK AND KILL IT WON'T RELOAD AGAIN 
; QUICKLY _ SUSSED THAT LOOP GONE
; -------------------------------------------------------------------
; FROM TIME __ Sun 21-Oct-2018 13:58:58
; TO   TIME __ Sun 21-Oct-2018 14:05:00
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

DetectHiddenWindows, ON
SetTitleMatchMode 3  ; EXACTLY

SoundBeep , 2000 , 100

;--------------------------------------------------------------------
;AUTOHOTKEYS
;--------------------------------------------------------------------

; Each array must be initialized before use:
FN_Array_1 := []
FN_Array_2 := []
DATE_MOD_Array := []

ArrayCount := 0
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk"	
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 02-SAVE AS KEY ENTER.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 14-Brightness With Dimmer.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES MYSMS & FB GB.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 19-SCRIPT_TIMER_UTIL.ahk"
ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 32-BRUTE BOOT DOWN.ahk"
ArrayCount += 1


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
	FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK"
}

ArrayCount += 1
FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"

FN_Array_1[ArrayCount] := "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 41-Minimize Chrome Close & Close RButton.ahk"

; ArrayCount += 1
; FN_Array_1[ArrayCount] := 
; ArrayCount += 1
; FN_Array_1[ArrayCount] := 


; -------------------------------------------------------------------
; ADD THE TERMINATOR VERSION NUMBER AND THEN WE ARE ABLE TO USE EXACT 
; STRING MATCHING IN CASE NOTEPAD HAD IT
; -------------------------------------------------------------------


; WHEN GOT AN ARRAY VALUE 
; YOU CAN SET IT TO NOTHING 
; BUT UNABLE TO TAKE IT VALUE AND PUT BACK AGAIN
; MUST USE REMOVE AND INSERT SAME POSITION AGAIN
;
; MAYBE BABY


Loop % ArrayCount
{
	AHK_TERMINATOR_VERSION:=" - AutoHotkey v"A_AhkVersion
	
	; WORKING CODE FOR A REFERENCE IF EVER REMOVE AND INSERT
	; TEMP_VAR_1:=FN_Array_1[A_Index]
	; FN_Array_1.RemoveAt(A_Index)
	; TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
	; TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
	; TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")
	; FN_Array_1.InsertAt(A_Index,TEMP_VAR_3)

	TEMP_VAR_1:=FN_Array_1[A_Index]
	TEMP_VAR_2="%AHK_TERMINATOR_VERSION%"
	TEMP_VAR_3=%TEMP_VAR_1%%TEMP_VAR_2%
	TEMP_VAR_3:=StrReplace(TEMP_VAR_3, """" , "")
	
	FN_Array_2.InsertAt(A_Index,TEMP_VAR_3)
	
	; ---------------------------------------------------------------
	; FINALLY GOT THERE WITH MY ARRAY HANDLE-SHIP
	; NICE TIME
	; [ Saturday 03:54:40 Am_20 October 2018 ]
	; REMOVE AND INSERT IF WANT TO CHANGE
	; LOOK AT HOW BEAUTIFUL OUR CODE ARE
	; ALL THOSE STRUGGLING
	; TOOK SOME TAPPING
	; HOW A LOOK AT MY GOLDEN EYE TIME KEEPER
	; Sat 20-Oct-2018 01:58:30
	; Sat 20-Oct-2018 03:58:00 __ 2 HOUR-ING FOR A TINY BIT OF CODE
	;                             TOOK LONGER TO MAKE IT PROGRAMMER THE LANGUAGE
	;                             BECAUSE I'M A CRIPPLE HURTING COLLAR BONE
	; ---------------------------------------------------------------
	; POINTLESS IN THE END MIGHT AS WELL HOLD TWO ARRAY SET
	; FIND OUT NEW INFO CAN'T PUT VALUE INTO NEW ARRAY AT INDEX
	;
	; LEARN SOMETHING CAN'T DO AS EXPECT HERE CREATE A SECOND ARRAY 
	; BUT UNABLE TO ADD ANYTHING UNLESS INSERT AGAIN
	; ---------------------------------------------------------------
}
  

Loop % ArrayCount
{
  Element := FN_Array_1[A_Index]
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
	Element_1 := FN_Array_1[A_Index]
	Element_2 := DATE_MOD_Array[A_Index]
	Element_3 := FN_Array_2[A_Index]
	
	SET_GO=FALSE
	IfExist, %Element_1%
		IF (!WinExist(Element_3) and !Element_2)
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
			Run, %Element_1%
		}
}

RETURN
