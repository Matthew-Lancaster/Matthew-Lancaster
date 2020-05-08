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





GOSUB SUB_STRIP_THE_HANDBRAKE_EXE_BATCH_GENERATOR_ADDTION_FILENAME

GOSUB SUB_MOVE_TO_DIRECTORY_STRUCTURE_FOLDER

RETURN
GOSUB SUB_SET_DATE_UNIT
GOSUB SUB_RENAME_ERROR_WHEN_WRONG
GOSUB SUB_RENAME_ERROR_WHEN_WRONG_2


RETURN



SUB_STRIP_THE_HANDBRAKE_EXE_BATCH_GENERATOR_ADDTION_FILENAME:

	Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.* ; ---- , R
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		StringUpper UPPER_OutExtension,OutExtension
		
		FILENAME_:=SUBStr(OutNameNoExt,1,INSTR(OutNameNoExt,"_T")-1)
		R_PATH = %OutDir%\%FILENAME_%.%UPPER_OutExtension%
;		MSGBOX %A_LoopFileFullPath%`n%R_PATH%
		
		FileMove, %A_LoopFileFullPath%, %R_PATH%
	}
	
RETURN


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
			R_PATH:=StrReplace(OutDir, "F:\MP3-YX-510_02_TS\V\", "F:\MP3-YX-510_02_TS_VIDEO\V2_4\")
			FileCreateDir, %R_PATH%
			
			R_PATH=%R_PATH%\%OutNameNoExt%.MP4

			IF FileExist(FILENAME)
			{
				; MSGBOX %R_PATH%
				FileMove, %FILENAME%, %R_PATH%
			}
				
			IF Mod(A_INDEX, 1000)=0 
				TOOLTIP % TS "`n" TS_2 "`n" FILENAME,100,100
		}
	}

RETURN



SUB_SET_DATE_UNIT:

	WORK_DO_COUNT=0
	
	; ---------------------------------------------------------------
	; LET SWITCH WITH AH BEGIN
	; ---------------------------------------------------------------
	
	; "F:\MP3-YX-510_02_TS\M"
	; F:\MP3-YX-510_02_TS\M
	; F:\MP3-YX-510_02_TS_VIDEO\V2_4
	
	DT1=DateSerial(2000,1,1)+TIMESERIAL(1,0,0)

	SUBST_1_DATE:= SubStr(A_LoopReadLine, 1, 8)
	SUBST_1_DATE_Y:= 2000
	SUBST_1_DATE_M:= 01
	SUBST_1_DATE_D:= 01
	
	TS=% SUBST_1_DATE_Y . SUBST_1_DATE_M . SUBST_1_DATE_D . 01 . 00 . 00
	
	Loop, Files, F:\MP3-YX-510_02_TS\M\*.* , R
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		FILENAME = %OutDir%\%OutFILENAME%

		IF INSTR(".MP3 .WAV .MP4",OutExtension)
		{
			FileSetTime, %TS% , %FILENAME% , M
			FileGetTime, TS_2, %FILENAME%, M
			TS+= 1, Days
			IF Mod(A_INDEX, 1000)=0 
				TOOLTIP % TS "`n" TS_2 "`n" FILENAME,100,100
		}
	}
	; --------------------------------------------------------
	; TS+= 1, Years              ; ---- NOT GOT YEAR PARAMETER
	; --------------------------------------------------------

	TS:=SubStr(TS, 1, 4)
	TS+= 1
	TS=% TS . 01 . 01 . 01 . 00 . 00

	Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.* , R
    {
		SplitPath, A_LoopFileFullPath, OutFILENAME, OutDir, OutExtension, OutNameNoExt, OutDrive
		FILENAME = %OutDir%\%OutFILENAME%

		IF INSTR(".MP3 .WAV .MP4 .WMV .AVI .MPG .MPEG .FLV",OutExtension)
		{
			FileSetTime, %TS% , %FILENAME% , M
			FileGetTime, TS_2, %FILENAME%, M
			TS+= 1, Days
			IF Mod(A_INDEX, 1000)=0 
				TOOLTIP % TS "`n" TS_2 "`n" FILENAME,100,100
		}
	}

	; EXIT ROUTINE HERE OR NEXT ONE WHEN CONVERT IDEA
	; -----------------------------------------------
RETURN


	
SUB_RENAME_ERROR_WHEN_WRONG_2:

	; "F:\MP3-YX-510_02_TS_VIDEO\V2_4\00 2001 MEDIA HARDCORE\00 ROOT_DATE\1993 a17 -- OLD -- 1993\50Vcd - Old_01.MPG\50Vcd - Old_01.MP4"
	; "F:\MP3-YX-510_02_TS_VIDEO\V2_4\00 2001 MEDIA HARDCORE\00 ROOT_DATE\1993 a17 -- OLD -- 1993\50Vcd - Old_01.MP4"
	
	Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2\*.* , R
    {
		SplitPath, A_LoopFileFullPath, 1OutFILENAME, 1OutDir, 1OutExtension, 1OutNameNoExt
		
		Loop, Files, F:\MP3-YX-510_02_TS_VIDEO\V2_4\*.*
		{
			SplitPath, A_LoopFileFullPath, 2OutFILENAME, 2OutDir, 2OutExtension, 2OutNameNoExt

			
			IF INSTR(1OutNameNoExt,2OutNameNoExt)
			{
			FILENAME = %2OutDir%\%1OutNameNoExt%.MP4
			
			MSGBOX % FILENAME
			FileMove, F:\MP3-YX-510_02_TS_VIDEO\V2_4\%OutFILENAME%, %FILENAME%
			}
		}
	}
	
RETURN



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
