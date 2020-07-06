;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 94-ALL_LOWER_THAN_NORMAL_PROCCES_PRIORITY_RESTORE.ahk
;# __ 
;# __ Autokey -- 94-ALL LOWER THAN NORMAL PROCCES PRIORITY RESTORE.ahk
;# __ 
;# BY __ Matt.Lan@Btinternet.com
;# __ 
;# Sun 26-Apr-2020 19:17:03
;# Sun 26-Apr-2020 23:58:00 -- 4 HOUR 40 MINUTE
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; WORK TIME
; -------------------------------------------------------------------
; SESSION 2 DAY 1
; Sun 26-Apr-2020 11:22:41
; Sun 26-Apr-2020 11:05:46 -- 12 HOUR 56 MINUTE
; -------------------------------------------------------------------
; SESSION 2 DAY 1
; Sun 26-Apr-2020 19:17:03
; Sun 26-Apr-2020 23:58:00 -- 4 HOUR 40 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; TO CREATE EFFORT FOR NEW MP3 UNIT
; -------------------------------------------------------------------
; 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; ------------------------------------------------------------------
; Location SOURCE
;---------------------------------------------------------------------
; ---------------------
; Matthew-Lancaster/Autokey -- 94-ALL_LOWER_THAN_NORMAL_PROCCES_PRIORITY_RESTORE.ahk
; ---------------------
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOHOTKEY/Autokey%20--%2094-ALL_LOW_PROCCES_PRIORITY_TO_NORMAL.ahk 
; ---------------------
; HTTPS://TINYURL.COM/SRZ69YL
; ---------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; CREDIT TO 
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; CREDIT TO **** AUTOHOTKEYS **** HELP PAGE -- SEARCH WORD -- Process
; -------------------------------------------------------------------
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; MAIN CREDIT MY OPERATION
; A MASSIVE EFFORT HARD WORK FIND READY MADE LEARN SCRIPT 
; LIKE HERE
; -------------------------------------------------------------------
; -------------------------------------------------------------------

#Warn
#NoEnv
#SingleInstance Force


RENAME_EXTENSION_SET_GO_EVERY_TIME=
O_LEN_STRING_HOLD_HWND=
RENAME_FILTER_EXTENSION_CASE_EXCLUDE=
RENAME_FILTER_EXTENSION_CASE_INCLUDE=
RENAME_EXTENSION_QUIET_WITH_AUDIO=
RENAME_EXTENSION_SET_DONE_QUIET=
; ------------------------------------------------ DECLARE VARIABLE




; GOSUB SUB_STRIP_THE_HANDBRAKE_EXE_BATCH_GENERATOR_ADDTION_FILENAME   ; ____ EXAMPLE __T1_C1 -- _T2_C1 -- _T3_C1 -- _T4_C1
; RETURN

; GOSUB SUB_MOVE_TO_DIRECTORY_STRUCTURE_FOLDER
; RETURN



REVERSE_OR_FORWARD=FORWARD
SOURCE_HDD_OR_PEN_DRIVE=PEN_DRIVE

; GOSUB SUB_SET_DATE_UNIT



; ONLY WHEN VIDEO
; GOSUB SUB_SET_DATE_FOLDER_ONLY_REMOTE_UNIT

; RUN HERE FOR REKETESS V115 

IF !FileExist("J:\M")
{
	MSGBOX ,4096,, NONE DRIVE INPUT ;
}
GOSUB RETREIVE_MODIFIED_DATE_SORTED_CONTENT_TO_LIST_FILE



RETURN


GOSUB SUB_RENAME_ERROR_WHEN_WRONG
GOSUB SUB_RENAME_ERROR_WHEN_WRONG_2
GOSUB SUB_RENAME_ERROR_WHEN_WRONG_2

RETURN


; -------------------------------------------------------------------
; -------------------------------------------------------------------
SUB_STRIP_THE_HANDBRAKE_EXE_BATCH_GENERATOR_ADDTION_FILENAME:
	Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.* ; ---- , R
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		StringUpper UPPER_OutExtension,OutExtension
		
		; -----------------------------------------------------------
		; -----------------------------------------------------------
		; HERE SHOULD BE INSTRREV
		; -----------------------------------------------------------
		; -----------------------------------------------------------
		; ; ____ EXAMPLE __T1_C1 -- _T2_C1 -- _T3_C1 -- _T4_C1
		FILENAME_:=SUBStr(OutNameNoExt,1,INSTR(OutNameNoExt,"_T",,-1)-1)   ;; REVERSE INSTR FINDER BY ,,-1
		R_PATH = %OutDir%\%FILENAME_%.%UPPER_OutExtension%
		; MSGBOX %A_LoopFileFullPath%`n%R_PATH%
		FileMove, %A_LoopFileFullPath%, %R_PATH%
	}
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
SUB_MOVE_TO_DIRECTORY_STRUCTURE_FOLDER:
	
	; ROUTINE WHEN AFTER CONVERT TO MP4
	; OR ROUTINE SAME AS MUSIC VERSION
	; ---------------------------------
	Loop, Files, F:\MP3-YX-510_02_TS\V\*.* , R
	{
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		
		IF INSTR(".MP3 .WAV .MP4 .WMV .AVI .MPG .MPEG .FLV",OutExtension)
		{
			FILENAME = F:\MP3-YX-510_02_TS_VIDEO\V2_4\%OutNameNoExt%.MP4
			R_PATH_2:=StrReplace(OutDir, "F:\MP3-YX-510_02_TS\V\", "F:\MP3-YX-510_02_TS_VIDEO\V2_4\")

			
			R_PATH_1=%R_PATH_2%\%OutNameNoExt%.MP4

			IF FileExist(FILENAME)
			{
				; MSGBOX %FILENAME%`n`n%R_PATH_1%`n`n%R_PATH_2%
				FileCreateDir, %R_PATH_2%
				FileMove, %FILENAME%, %R_PATH_1%
			}
				
			IF Mod(A_INDEX, 1000)=0 
				TOOLTIP % TS "`n" TS_2 "`n" FILENAME,100,100
		}
	}
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
SUB_SET_DATE_UNIT:

	WORK_DO_COUNT=0
	
	; ---------------------------------------------------------------
	; LET SWITCH WITH AH BEGIN
	; ---------------------------------------------------------------
	
	; "F:\MP3-YX-510_02_TS\M"
	; F:\MP3-YX-510_02_TS\M
	; F:\MP3-YX-510_02_TS_VIDEO\V2_4

	MSGBOX_OFF=TRUE
	
	
	IF REVERSE_OR_FORWARD=REVERSE
		SUBST_1_DATE_Y:= 1980 ; IF FORWARD SET
	IF REVERSE_OR_FORWARD=REVERSE
		SUBST_1_DATE_Y:= 19080 ; WHEN REVERSE SET
	SUBST_1_DATE_M:= 01
	SUBST_1_DATE_D:= 01
	
	TS=% SUBST_1_DATE_Y . SUBST_1_DATE_M . SUBST_1_DATE_D . 01 . 00 . 00
	
	INFO_DISPLAY_ONCE=TRUE
	
	FileList=
	
	IF SOURCE_HDD_OR_PEN_DRIVE=HDD
		DRIVE_NAME=F:\MP3-YX-510_02_TS\M\*.*
	IF SOURCE_HDD_OR_PEN_DRIVE=PEN_DRIVE
		DRIVE_NAME=I:\M\*.*
	
	Loop, Files, %DRIVE_NAME% , FDR
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		FILENAME = %OutDir%\%OutFILENAME%

		IF INSTR(FILEEXIST(FILENAME), "D")
		{
			FileSetTime, %TS% , %FILENAME% , M ,1     ; ,1 MEAN SET DATE FOLDER DEFAULT ARE 0 1 2 -- 0 ARE FILE ONLY -- 1 ARE FILE AND FOLDER
			FileSetTime, %TS% , %FILENAME% , C ,1     ; ,1 NOT REALLY REQUIRE DEFAULT SINGLE FOLDER ALLOW SET DATE
			CONTENT_NAME_SET=%FILENAME%
			
			IF REVERSE_OR_FORWARD=FORWARD
				TS+= 1, Days             ; FORWARD
			IF REVERSE_OR_FORWARD=REVERSE
				TS+= -1, Days            ; REVERSE
		}

		IF INSTR(".MP3 .WAV .MP4",OutExtension)
		{
			FileSetTime, %TS% , %FILENAME% , M
			FileSetTime, %TS% , %FILENAME% , C
			FileGetTime, TS_2,  %FILENAME%, M
			CONTENT_NAME_SET=%FILENAME%
			IF REVERSE_OR_FORWARD=FORWARD
				TS+= 1, Days             ; FORWARD
			IF REVERSE_OR_FORWARD=REVERSE
				TS+= -1, Days            ; REVERSE
			
			IF Mod(A_INDEX, 500)=0 
				TOOLTIP % 	SubStr(TS, 1, 8) "`n" SubStr(TS_2, 1, 8 ) "`n" FILENAME,100,100
		
			IF INFO_DISPLAY_ONCE
				IF !MSGBOX_OFF
					MSGBOX % TS_2 "`n" FILENAME
			INFO_DISPLAY_ONCE=
		}
	}

	IF !MSGBOX_OFF
		MSGBOX % TS_2 "`n" CONTENT_NAME_SET
	; --------------------------------------------------------
	; 2000 TO 2011 FOR MUSIC CHUNK
	; --------------------------------------------------------
	
	; --------------------------------------------------------
	; TS+= 1, Years               ; ---- NOT GOT YEAR PARAMETER
	; ROUND THE YEAR NUMBER VIDEO STARTING
	; SUBTRACT ONE YEAR TO DATE GIVE
	; AND NOTHING MONTH AND DAY
	; --------------------------------------------------------
	; THE MP3 PLAYER WANT MOST RECENT CONTENT PLAY FIRST
	; CORRECTION PREFER IS -- TIME BEFORE FIRST
	; CARE TOO MANY DATE ONLY ALLOW LESS THAN 2022
	; --------------------------------------------------------
	
	
	; ---------------------------------------------------------------
	; VIDEO ---------------------------------------------------------
	; ---------------------------------------------------------------
	TS:=SubStr(TS, 1, 4)
	IF REVERSE_OR_FORWARD=REVERSE
		TS+= -1   ; WHEN ROUND DOWN NOT REQUIRE SUBTRACT ONE YEAR -- FORWARD
		; TS+= -1 ; WHEN ROUND DOWN NOT REQUIRE SUBTRACT ONE YEAR -- REVERSE -- REM OUT WAY
	TS=% TS . 01 . 01 . 01 . 00 . 00

	INFO_DISPLAY_ONCE=TRUE
	
	IF SOURCE_HDD_OR_PEN_DRIVE=HDD
		DRIVE_NAME=F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.*
	IF SOURCE_HDD_OR_PEN_DRIVE=PEN_DRIVE
		DRIVE_NAME=I:\V\*.*

	Loop, Files, %DRIVE_NAME% , FDR
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		FILENAME = %OutDir%\%OutFILENAME%

		IF INSTR(FILEEXIST(FILENAME), "D")
		{
			FileSetTime, %TS% , %FILENAME% , M ,1     ; ,1 MEAN SET DATE FOLDER DEFAULT ARE 0 1 2 -- 0 ARE FILE ONLY -- 1 ARE FILE AND FOLDER 
			FileSetTime, %TS% , %FILENAME% , C ,1
			CONTENT_NAME_SET=%FILENAME%
			IF REVERSE_OR_FORWARD=FORWARD
				TS+= 1, Days     ; FORWARD
			IF REVERSE_OR_FORWARD=REVERSE
				TS+= -1, Days      ; REVERSE
		}
		
		IF INSTR(".MP3 .WAV .MP4 .WMV .AVI .MPG .MPEG .FLV",OutExtension)
		{
			FileSetTime, %TS% , %FILENAME% , M
			FileSetTime, %TS% , %FILENAME% , C
			FileGetTime, TS_2,  %FILENAME%, M
			CONTENT_NAME_SET=%FILENAME%
			IF REVERSE_OR_FORWARD=FORWARD
				TS+= 1, Days     ; FORWARD
			IF REVERSE_OR_FORWARD=REVERSE
				TS+= -1, Days      ; REVERSE
			; -------------------------------------------------------
			; -- COUNT GO BACKWARD NOT FORWARD DEVICE HERE ---- MP3-YX-510 ---- MP3 PLAYER
			; -------------------------------------------------------
			IF Mod(A_INDEX, 20)=0 
				TOOLTIP % SubStr(TS, 1, 8) "`n" SubStr(TS_2, 1, 8 ) "`n" FILENAME,100,100
			IF INFO_DISPLAY_ONCE
				IF !MSGBOX_OFF
					MSGBOX % TS_2 "`n" FILENAME
			INFO_DISPLAY_ONCE=
			}
	}
	
	IF !MSGBOX_OFF
		MSGBOX % TS_2 "`n" CONTENT_NAME_SET

	TOOLTIP
		
	; EXIT ROUTINE HERE OR NEXT ONE WHEN CONVERT IDEA
	; -----------------------------------------------
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------



; -------------------------------------------------------------------
; -------------------------------------------------------------------
SUB_SET_DATE_FOLDER_ONLY_REMOTE_UNIT:

	WORK_DO_COUNT=0
	OLD_OutDir=
	
	; ---------------------------------------------------------------
	; LET SWITCH WITH AH BEGIN
	; ---------------------------------------------------------------
	
	; "F:\MP3-YX-510_02_TS\M"
	; F:\MP3-YX-510_02_TS\M
	; F:\MP3-YX-510_02_TS_VIDEO\V2_4

	MSGBOX_OFF=TRUE
	
	REVERSE_OR_FORWARD=FORWARD
	
	SUBST_1_DATE_Y:= 1980 ; IF FORWARD SET
	SUBST_1_DATE_M:= 01
	SUBST_1_DATE_D:= 01
	
	TS=% SUBST_1_DATE_Y . SUBST_1_DATE_M . SUBST_1_DATE_D . 01 . 00 . 00
	
	INFO_DISPLAY_ONCE=TRUE
	
	FileList=
	; ---------------------------------------------------------------
	; MUSIC ONE REMOTE UNIT -- 001 OF 002 -- FILE FOLDER RECURSE
	; ---------------------------------------------------------------

	DRIVE_NAME=J:\M\*.*
	Loop, Files, %DRIVE_NAME% , FDR
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		FILENAME = %OutDir%\%OutFILENAME%

		; IF INSTR(FILEEXIST(FILENAME), "D")
		{
			IF REVERSE_OR_FORWARD=FORWARD
				TS+= 1, Days             ; FORWARD
			IF REVERSE_OR_FORWARD=REVERSE
				TS+= -1, Days            ; REVERSE
			FileSetTime, %TS% , %FILENAME% , M ,1     ; ,1 MEAN SET DATE FOLDER DEFAULT ARE 0 1 2 -- 0 ARE FILE ONLY -- 1 ARE FILE AND FOLDER
			FileSetTime, %TS% , %FILENAME% , C ,1     ; ,1 NOT REALLY REQUIRE DEFAULT SINGLE FOLDER ALLOW SET DATE
			CONTENT_NAME_SET=%FILENAME%
			
			; IF OutDir<>%OLD_OutDir%
			; {
				; TS+= -100, Days  
			; }
			; OLD_OutDir=%OutDir%

			IF Mod(A_INDEX, 500)=0 
				TOOLTIP % 	SubStr(TS, 1, 8) "`n" FILENAME,100,100
		}
	}

	
	; ---------------------------------------------------------------
	; MUSIC ONE REMOTE UNIT -- 002 OF 002 -- FILE FOLDER RECURSE
	; ---------------------------------------------------------------
	
	REVERSE_OR_FORWARD=REVERSE
	
	SUBST_1_DATE_Y:= 1980
	SUBST_1_DATE_M:= 01
	SUBST_1_DATE_D:= 01
	
	TS=% SUBST_1_DATE_Y . SUBST_1_DATE_M . SUBST_1_DATE_D . 01 . 00 . 00

	DRIVE_NAME=J:\M\*.*
	Loop, Files, %DRIVE_NAME% , DR
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		FILENAME = %OutDir%\%OutFILENAME%

		IF INSTR(FILEEXIST(FILENAME), "D")
		{
			FileSetTime, %TS% , %FILENAME% , M ,1     ; ,1 MEAN SET DATE FOLDER DEFAULT ARE 0 1 2 -- 0 ARE FILE ONLY -- 1 ARE FILE AND FOLDER
			FileSetTime, %TS% , %FILENAME% , C ,1     ; ,1 NOT REALLY REQUIRE DEFAULT SINGLE FOLDER ALLOW SET DATE
			CONTENT_NAME_SET=%FILENAME%
			
			IF REVERSE_OR_FORWARD=FORWARD
				TS+= 1, Days             ; FORWARD
			IF REVERSE_OR_FORWARD=REVERSE
				TS+= -1, Days            ; REVERSE
			IF Mod(A_INDEX, 500)=0 
				TOOLTIP % 	SubStr(TS, 1, 8) "`n" FILENAME,100,100
		}
	}

	
	
	RETURN
	
	
	; ---------------------------------------------------------------
	; VIDEO REMOTE UNIT -- 001 OF 002 -- FILE FOLDER RECURSE
	; ---------------------------------------------------------------
	
	
	REVERSE_OR_FORWARD=FORWARD
	
	SUBST_1_DATE_Y:= 2010 ; IF FORWARD SET
	SUBST_1_DATE_M:= 01
	SUBST_1_DATE_D:= 01
	
	TS=% SUBST_1_DATE_Y . SUBST_1_DATE_M . SUBST_1_DATE_D . 01 . 00 . 00

	DRIVE_NAME=I:\V\*.*
	Loop, Files, %DRIVE_NAME% , FDR
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		FILENAME = %OutDir%\%OutFILENAME%

		; IF INSTR(FILEEXIST(FILENAME), "D")
		{
		
			IF OutDir<>%OLD_OutDir%
			{
				TS+= -50, Days  
			}
			OLD_OutDir=%OutDir%
			FileSetTime, %TS% , %FILENAME% , M ,1     ; ,1 MEAN SET DATE FOLDER DEFAULT ARE 0 1 2 -- 0 ARE FILE ONLY -- 1 ARE FILE AND FOLDER
			FileSetTime, %TS% , %FILENAME% , C ,1     ; ,1 NOT REALLY REQUIRE DEFAULT SINGLE FOLDER ALLOW SET DATE
			CONTENT_NAME_SET=%FILENAME%
			
			IF REVERSE_OR_FORWARD=FORWARD
				TS+= 1, Days             ; FORWARD
			IF REVERSE_OR_FORWARD=REVERSE
				TS+= -1, Days            ; REVERSE
			IF Mod(A_INDEX, 500)=0 
				TOOLTIP % 	SubStr(TS, 1, 8) "`n" FILENAME,100,100
		}
	}

	; ---------------------------------------------------------------
	; VIDEO REMOTE UNIT -- 002 OF 002 -- FOLDER RECURSE
	; ---------------------------------------------------------------

	REVERSE_OR_FORWARD=REVERSE
	
	SUBST_1_DATE_Y:= 2000 ; IF FORWARD SET
	SUBST_1_DATE_M:= 01
	SUBST_1_DATE_D:= 01
	
	TS=% SUBST_1_DATE_Y . SUBST_1_DATE_M . SUBST_1_DATE_D . 01 . 00 . 00

	DRIVE_NAME=I:\V\*.*
	Loop, Files, %DRIVE_NAME% , DR
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		FILENAME = %OutDir%\%OutFILENAME%

		IF INSTR(FILEEXIST(FILENAME), "D")
		{
			FileSetTime, %TS% , %FILENAME% , M ,1     ; ,1 MEAN SET DATE FOLDER DEFAULT ARE 0 1 2 -- 0 ARE FILE ONLY -- 1 ARE FILE AND FOLDER
			FileSetTime, %TS% , %FILENAME% , C ,1     ; ,1 NOT REALLY REQUIRE DEFAULT SINGLE FOLDER ALLOW SET DATE
			CONTENT_NAME_SET=%FILENAME%
			
			IF REVERSE_OR_FORWARD=FORWARD
				TS+= 1, Days             ; FORWARD
			IF REVERSE_OR_FORWARD=REVERSE
				TS+= -1, Days            ; REVERSE
			IF Mod(A_INDEX, 500)=0 
				TOOLTIP % 	SubStr(TS, 1, 8) "`n" FILENAME,100,100
		}
	}


	TOOLTIP
		
	; EXIT ROUTINE HERE OR NEXT ONE WHEN CONVERT IDEA
	; -----------------------------------------------
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------



; -------------------------------------------------------------------
; -------------------------------------------------------------------
RETREIVE_MODIFIED_DATE_SORTED_CONTENT_TO_LIST_FILE:

	; Example #4: Retrieve file names sorted by modification date:
	FileList =
	IF SOURCE_HDD_OR_PEN_DRIVE=HDD
	{
		Loop, Files, J:\M\*.* , R
			FileList = %FileList%%A_LoopFileTimeModified%`t%A_LoopFileFullPath%`n
		; Loop, Files, F:\MP3-YX-510_02_TS\M\*.* , R
			; FileList = %FileList%%A_LoopFileTimeModified%`t%A_LoopFileFullPath%`n
		; Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.* , R
			; FileList = %FileList%%A_LoopFileTimeModified%`t%A_LoopFileFullPath%`n
	}
	FileList =
	IF SOURCE_HDD_OR_PEN_DRIVE=PEN_DRIVE
	{
		Loop, Files, J:\M\*.* , R
			FileList = %FileList%%A_LoopFileTimeModified%`t%A_LoopFileFullPath%`n
		; Loop, Files, I:\M\*.* , R
			; FileList = %FileList%%A_LoopFileTimeModified%`t%A_LoopFileFullPath%`n
		; Loop, Files, I:\V\*.* , R
			; FileList = %FileList%%A_LoopFileTimeModified%`t%A_LoopFileFullPath%`n
	}	
	
	FileList =
	Loop, Files, K:\M\*.* , R
		FileList = %FileList%%A_LoopFileTimeModified%`t%A_LoopFileFullPath%`n
		
	; K:\M\*.*
	VAR_PLAYER=V115_MUSIC__
	FILE_PLAY_SCRIPT_OUT=C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT --%VAR_PLAYER%__#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_MUSIC=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95__MP3-%VAR_PLAYER%__FULL_#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_HEADER=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95__MP3-%VAR_PLAYER%__HEADER_#NFS_EX_.TXT

	FileList =
	Loop, Files, J:\M\*.* , R
		FileList = %FileList%%A_LoopFileTimeModified%`t%A_LoopFileFullPath%`n
	; J:\M\*.*
	VAR_PLAYER=V115_MUSIC_MINDREADER__
	FILE_PLAY_SCRIPT_OUT=C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT --%VAR_PLAYER%__#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_MUSIC=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95__MP3-%VAR_PLAYER%__FULL_#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_HEADER=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95__MP3-%VAR_PLAYER%__HEADER_#NFS_EX_.TXT

	; BOOK BOO
	; M:\TALKBOOK\*.*
	; ---------------------------------------------------------------
	; FileList =
	; VAR_PLAYER=V115_MINDREADER__
	; FILE_PLAY_SCRIPT_OUT=C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT --%VAR_PLAYER%__#NFS_EX_.TXT
	; FILE_PLAY_SCRIPT_OUT_MUSIC=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95__MP3-%VAR_PLAYER%__FULL_#NFS_EX_.TXT
	; FILE_PLAY_SCRIPT_OUT_HEADER=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95__MP3-%VAR_PLAYER%__HEADER_#NFS_EX_.TXT
	; ---------------------------------------------------------------

	; BOOK BOO
	; L:\MUSIC
	; ---------------------------------------------------------------
	; FileList =
	; VAR_PLAYER=SANDISK_256G_MUSIC_MINDREADER__
	; FILE_PLAY_SCRIPT_OUT=C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT --%VAR_PLAYER%__#NFS_EX_.TXT
	; FILE_PLAY_SCRIPT_OUT_MUSIC=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95__MP3-%VAR_PLAYER%__FULL_#NFS_EX_.TXT
	; FILE_PLAY_SCRIPT_OUT_HEADER=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95__MP3-%VAR_PLAYER%__HEADER_#NFS_EX_.TXT
	; ---------------------------------------------------------------


	; VAR_PLAYER=V115_M_TB
	; FILE_PLAY_SCRIPT_OUT=C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT - MP3-%VAR_PLAYER%_FILE_SCRIPT_#NFS_EX_.TXT


	Sort, FileList  ; Sort by date.

	VAR_PLAYER=YX-510
	VAR_PLAYER=V115
	FILE_PLAY_SCRIPT_OUT=C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT - MP3-%VAR_PLAYER%_FILE_SCRIPT_#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_01=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95-SET_DATE_FILE__MP3-%VAR_PLAYER%_FILE_SCRIPT_#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_VIDEO=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95-SET_DATE_FILE__MP3-%VAR_PLAYER%_FILE_SCRIPT_V_#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_MUSIC=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95-SET_DATE_FILE__MP3-%VAR_PLAYER%_FILE_SCRIPT_M_#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_T_01=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95-SET_DATE_FILE__MP3-%VAR_PLAYER%_FILE_SCRIPT_TMP_#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_T_VIDEO=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95-SET_DATE_FILE__MP3-%VAR_PLAYER%_FILE_SCRIPT_V_TMP_#NFS_EX_.TXT
	FILE_PLAY_SCRIPT_OUT_T_MUSIC=C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95-SET_DATE_FILE__MP3-%VAR_PLAYER%_FILE_SCRIPT_M_TMP_#NFS_EX_.TXT
	IF FileExist(FILE_PLAY_SCRIPT_OUT_T_VIDEO)
		FileDelete, %FILE_PLAY_SCRIPT_OUT_T_VIDEO%
	IF FileExist(FILE_PLAY_SCRIPT_OUT_T_MUSIC)
		FileDelete, %FILE_PLAY_SCRIPT_OUT_T_MUSIC%
	IF FileExist(FILE_PLAY_SCRIPT_OUT_T_01)
		FileDelete, %FILE_PLAY_SCRIPT_OUT_T_01%

	Loop, Parse, FileList, `n
	{
		if A_LoopField =  ; Omit the last linefeed (blank item) at the end of the list.
			continue
		StringSplit, FileItem, A_LoopField, %A_Tab%  ; Split into two parts at the tab char.
		FileItem1 :=SubStr(FileItem1, 1, 8 )

		SplitPath, FileItem2, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive

		FileGetTime, VV,  %OutDir%, M
		VV :=SubStr(VV, 1, 8 )
		FILE_PLAY_SCRIPT_LINE=%VV% .. %FileItem1% .. %FileItem2%`n
		FILE_PLAY_SCRIPT_LINE=%FileItem1% .. %FileItem2%`n
		
		IF INSTR(".MP3 .WAV .MP4 .WMV .AVI .MPG .MPEG .FLV",OutExtension)
		{
		
			FileAppend,%FILE_PLAY_SCRIPT_LINE%,%FILE_PLAY_SCRIPT_OUT_T_01%
			IF INSTR(OutDir, "\M\")
				FileAppend,%FILE_PLAY_SCRIPT_LINE%,%FILE_PLAY_SCRIPT_OUT_T_MUSIC%
			IF INSTR(OutDir, "I:\V\")
				FileAppend,%FILE_PLAY_SCRIPT_LINE%,%FILE_PLAY_SCRIPT_OUT_T_VIDEO%
			IF INSTR(OutDir, "\V2_4\")
				FileAppend,%FILE_PLAY_SCRIPT_LINE%,%FILE_PLAY_SCRIPT_OUT_T_VIDEO%
		}
		
	}

	; FileList=
	; Loop, read, %FILE_PLAY_SCRIPT_OUT_T_VIDEO%
		; FileList = %FileList%%A_LoopReadLine%`n
	; Sort, FileList  ; SORT BY DATE OF FOLDER WHICH IS FIRST COLUMN
	; IF FileExist(FILE_PLAY_SCRIPT_OUT_T_VIDEO)
		; FileDelete, %FILE_PLAY_SCRIPT_OUT_T_VIDEO%
	; FileAppend,%FILELIST%,%FILE_PLAY_SCRIPT_OUTT__VIDEO%
	
	TOOLTIP 

	FileCopy, %FILE_PLAY_SCRIPT_OUT_T_01%, %FILE_PLAY_SCRIPT_OUT%,1
	FileCopy, %FILE_PLAY_SCRIPT_OUT_T_01%, %FILE_PLAY_SCRIPT_OUT_01%,1
	FileCopy, %FILE_PLAY_SCRIPT_OUT_T_MUSIC%, %FILE_PLAY_SCRIPT_OUT_MUSIC%,1
	FileCopy, %FILE_PLAY_SCRIPT_OUT_T_VIDEO%, %FILE_PLAY_SCRIPT_OUT_VIDEO%,1
	IF FileExist(FILE_PLAY_SCRIPT_OUT_T_VIDEO)
		FileDelete, %FILE_PLAY_SCRIPT_OUT_T_VIDEO%
	IF FileExist(FILE_PLAY_SCRIPT_OUT_T_MUSIC)
		FileDelete, %FILE_PLAY_SCRIPT_OUT_T_MUSIC%
	IF FileExist(FILE_PLAY_SCRIPT_OUT_T_01)
		FileDelete, %FILE_PLAY_SCRIPT_OUT_01%
	
 	RUN,C:\PROGRAM FILES (X86)\NOTEPAD++\NOTEPAD++.EXE -n1 "%FILE_PLAY_SCRIPT_OUT_01%"
	RUN,C:\PROGRAM FILES (X86)\NOTEPAD++\NOTEPAD++.EXE -n1 "%FILE_PLAY_SCRIPT_OUT_VIDEO%"
 	RUN,C:\PROGRAM FILES (X86)\NOTEPAD++\NOTEPAD++.EXE -n1 "%FILE_PLAY_SCRIPT_OUT_MUSIC%"
	
	MSGBOX ,4096,, DONE ....,2 ; PUT A MSGBOX SO NOTEPAD++ WILL LOAD THE CHANGE CONTENT -- HAS TO LOSE FOCUS BEFORE SUPPLY UPDATE AR
	#WinActivateForce, ahk_class Notepad++
	
	; -------------------------------------------------------------------
	; MSGBOX ,4100,, HERE MAKE ASK WITH YES NOT 
	; MSGBOX ,4096,,
	; -------------------------------------------------------------------
	; Group #4: Modality
	; To indicate the modality of the dialog box, add one of the following values:
	; Function
	; Decimal Value
	; Hex Value
	; System Modal (always on top) 4096 0x1000 
	; Task Modal 8192 0x2000 
	; Always-on-top (style WS_EX_TOPMOST)
	; (like System Modal but omits title bar icon) 262144 0x40000 
	; -------------------------------------------------------------------

	
	
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
SUB_RENAME_ERROR_WHEN_WRONG_2:
; -------------------------------------------------------------------
; ACCIDENT -- CHOP TRIM STRING MORE THAN WANT WITH SUB ROUTINE
; SUB_STRIP_THE_HANDBRAKE_EXE_BATCH_GENERATOR_ADDTION_FILENAME:
; -------------------------------------------------------------------

STRING_NOT_USE_AGAIN=

T_INDEX_1=0
T_INDEX_2=0
Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.*, F                          ; Include Files and Directories FD
    T_INDEX_1+=1
T_INDEX_2=0
Loop, Files, F:\MP3-YX-510_02_TS\V\*.*, FR                                 ; Include Files and Directories FD
    T_INDEX_2+=1

FileList1 =
Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.*, F                          ; Include Files and Directories FD
    FileList1 = %FileList1%%A_LoopFileTimeModified%`t%A_LoopFileName%`n
Sort, FileList1                                                             ; Sort by date.

FileList2 =
Loop, Files, F:\MP3-YX-510_02_TS\V\*.*, FR                                  ; Include Files and Directories FD AND RECURSE
    FileList2 = %FileList2%%A_LoopFileName%`t%A_LoopFileFullPath%`n
Sort, FileList2                                                             ; Sort by date.

T_INDEX=0

; -------------------------------------------------------------------
; FileList2 ---- PROPER 
; -------------------------------------------------------------------
Loop, Parse, FileList2, `n
{
    if A_LoopField =  ; Omit the last linefeed (blank item) at the end of the list.
        continue
	StringSplit, FileItem2, A_LoopField, %A_Tab%  ; Split into two parts at the tab char.

	X_INDEX=%A_INDEX%
	T_INDEX+=1
	
	; MSGBOX %A_INDEX% `n%FileItem21% `n%FileItem22%
	
	; ---------------------------------------------------------------
	; FileList1 ---- NOT PROPER 
	; ---------------------------------------------------------------
	Loop, Parse, FileList1, `n
	{
		if A_LoopField =  ; Omit the last linefeed (blank item) at the end of the list.
			continue

		StringSplit, FileItem1, A_LoopField, %A_Tab%  ; Split into two parts at the tab char.
		
		SplitPath, FileItem12, 1OutFILENAME, 1OutDir, 1OutExtension, 1OutNameNoExt
		SplitPath, FileItem21, 2OutFILENAME, 2OutDir, 2OutExtension, 2OutNameNoExt

		; MSGBOX %A_INDEX% `n%X_INDEX%`n%FileItem21% `n%FileItem22% `n%FileItem12%
		
		; ---------------------------------------------------------------
		; 1 NOT PROPER
		; 2 PROPER
		; ---------------------------------------------------------------
		IF 1OutNameNoExt<>%2OutNameNoExt%
		IF INSTR(2OutNameNoExt,1OutNameNoExt)
		{

			STRING_COMPARE_02=--%A_INDEX%
			IF INSTR(STRING_NOT_USE_AGAIN,STRING_COMPARE_02)=0
			{
			
				STRING_NOT_USE_AGAIN=%STRING_NOT_USE_AGAIN%--%A_INDEX%
				
				; -------------------------------------------------------
				; 1ST NEST LOOP -- FileList2 ---- PROPER
				; 2ND NEST LOOP -- FileList1 ---- NOT PROPER
				; -------------------------------------------------------
				; FileList1 ---- NOT PROPER  ---- A_INDEX ---- NOT PROPER
				; FileList2 ---- PROPER      ---- X_INDEX ---- PROPER
				; -------------------------------------------------------
				
				; MSGBOX %T_INDEX% --OF-- %T_INDEX_2% --OF-- %T_INDEX_1%`n%X_INDEX% -- PROPER`n%A_INDEX% -- NOT PROPER`n`n%FileItem21%`n%X_INDEX%`n`n%FileItem22%`n`n%FileItem12%`n%A_INDEX%
				
				FILE_RENAME1=F:\MP3-YX-510_02_TS_VIDEO\V2_4\%FileItem12%
				FILE_RENAME2=F:\MP3-YX-510_02_TS_VIDEO\V2_4\%2OutNameNoExt%.MP4
				IF FileExist(FILE_RENAME1)
				{
					; -----------------------------------------------
					; ALL DOUBLE CHECK 
					; -----------------------------------------------
					MSGBOX %FILE_RENAME1%`n%FILE_RENAME2%
				
					FileMove, %FILE_RENAME1%, %FILE_RENAME2%
				}
			}		
		}		
	}
}	
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------
	
	
; -------------------------------------------------------------------
; -------------------------------------------------------------------
SUB_RENAME_ERROR_WHEN_WRONG_3:

	; "F:\MP3-YX-510_02_TS_VIDEO\V2_4\00 2001 MEDIA HARDCORE\00 ROOT_DATE\1993 a17 -- OLD -- 1993\50Vcd - Old_01.MPG\50Vcd - Old_01.MP4"
	; "F:\MP3-YX-510_02_TS_VIDEO\V2_4\00 2001 MEDIA HARDCORE\00 ROOT_DATE\1993 a17 -- OLD -- 1993\50Vcd - Old_01.MP4"
	
	Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.*
    {
		SplitPath, A_LoopFileFullPath, 1OutFILENAME, 1OutDir, 1OutExtension, 1OutNameNoExt
		
		FP_1:=A_LoopFileFullPath
		
		Loop, Files, F:\MP3-YX-510_02_TS\V\*.* , R
		{
			SplitPath, A_LoopFileFullPath, 2OutFILENAME, 2OutDir, 2OutExtension, 2OutNameNoExt

			
			IF INSTR(1OutNameNoExt,2OutNameNoExt)
			{
			FILENAME = %1OutDir%\%2OutNameNoExt%.MP4
			IF FP_1<>%FILENAME%
			{
			MSGBOX % FP_1 "`n" FILENAME
			; FileMove, %FP_1%, %FILENAME%
			}
			}
		}
	}
	
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
SUB_RENAME_ERROR_WHEN_WRONG:

	; "F:\MP3-YX-510_02_TS_VIDEO\V2_4\00 2001 MEDIA HARDCORE\00 ROOT_DATE\1993 a17 -- OLD -- 1993\50Vcd - Old_01.MPG\50Vcd - Old_01.MP4"
	; "F:\MP3-YX-510_02_TS_VIDEO\V2_4\00 2001 MEDIA HARDCORE\00 ROOT_DATE\1993 a17 -- OLD -- 1993\50Vcd - Old_01.MP4"
	
	Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.* , R
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		IF INSTR(OutDir,OutNameNoExt)
		{
			; R_PATH:=StrReplace(OutDir,OutNameNoExt.MP4,"")
			R_PATH:=SUBStr(OutDir,1,INSTR(OutDir,OutNameNoExt)-2)
			
			; MSGBOX %R_PATH%\%OutFILENAME%
			; MSGBOX F:\MP3-YX-510_02_TS_VIDEO\V2_4\%OutFILENAME%
			
			FileMove, %A_LoopFileFullPath%, F:\MP3-YX-510_02_TS_VIDEO\V2_4\%OutNameNoExt%.MP4
			FileRemoveDir, %OutDir%
			
			; MSGBOX F:\MP3-YX-510_02_TS_VIDEO\V2_4\%OutNameNoExt%.MP4`n%R_PATH%%OutFILENAME%
			FileMove, F:\MP3-YX-510_02_TS_VIDEO\V2_4\%OutFILENAME%, %R_PATH%\%OutFILENAME%
			
			; MSGBOX % A_LoopFileFullPath "`n" R_PATH OutFILENAME
			
		}
	}
	
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; -------------------------------------------------------------------
HOW_TO_PREVENT_CREATION_OF_SYSTEM_VOLUME_INFORMATION_FOLDER_IN_WINDOWS_10_FOR_USB_FLASH_DRIVES___SUPER_USER_:
; IFEXIST, J:\M\01 SOUND EFFECT & TECHNO SAMPLES_REKETEKESS\BBC Micro.wav
; {
	; ; ---------------------------------------------------------------
	; ; How to prevent creation of "System Volume Information" folder in Windows 10 for USB flash drives? - Super User 
	; ; https://superuser.com/questions/1199823/how-to-prevent-creation-of-system-volume-information-folder-in-windows-10-for
	; ; ---------------------------------------------------------------
	; ; IDEA DELETE \SYSTEM VOLUME INFO FOLDER
	; ; AS PLAYER FIND FILE AN REPORT NOT RECOGNITION
	; ; USE IO-BIT-UNLOCKER
	; ; WHEN DELETE CREATE FILE REPLACE
	; ; ---------------------------------------------------------------
	
	; IF TRUE=FALSE
	; {
		; FileSetAttrib, -RHS, J:\System Volume Information\* ,1 ,1

		; FileDelete, J:\System Volume Information\WPSettings.dat
		; FileRemoveDir, J:\System Volume Information
		
			; Loop, J:\System Volume Information\*, ,1 
		; {
			; FileDelete, A_LoopFileFullPath
			; MSGBOX % A_LoopFileFullPath
		; }
		; FileRemoveDir, J:\System Volume Information,1
		; FileRecycle "J:\System Volume Information"
		; DirDelete "J:\System Volume Information", 1
	; }
; }

; IFEXIST, I:\M\01 SOUND EFFECT & TECHNO SAMPLES_REKETEKESS\BBC Micro.wav
; {
	; IF TRUE=FALSE
	; {
		; FileSetAttrib, -RHS, I:\System Volume Information\* ,1 ,1
		
		; ; ERROR: File ownership cannot be applied on insecure file systems;
		; ; there is no support for ACLs.
		; ; run, %comspec% /k takeown /r /f "I:\System Volume Information", , max
		
		; Loop, I:\System Volume Information\*, ,1 
		; {
			; FileDelete, A_LoopFileFullPath
			; MSGBOX % A_LoopFileFullPath
		; }
		; FileRemoveDir, I:\System Volume Information,1
		
	; }
	
	
; }
	
RETURN
; -------------------------------------------------------------------
; -------------------------------------------------------------------
