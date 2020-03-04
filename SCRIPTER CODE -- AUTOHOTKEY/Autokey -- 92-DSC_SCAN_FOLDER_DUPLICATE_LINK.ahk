; -------------------------------------------------------------------
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
; Autokey -- 92-DSC_SCAN_FOLDER_DUPLICATE_LINK.ahk
;
; -------------------------------------------------------------------
; Wed 19-Feb-2020 22:30:00
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; TAKEN FROM AND RE-USER
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 84-RENAME_FILE_EXTENSION_CASE_INCLUDE.ahk
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; TIME SPEND HERE TODAY ALSO DAY BEFORE
; Mon 16-Sep-2019 22:01:51
; Tue 17-Sep-2019 00:51:04
; AND OTHER CODE UNTIL 2 AM
; AND WORK ENTERTAINER UNTIL 8 AM
; AND THEN UP AT 10 AM TUESDAY AFTER SUDDEN SLONK OUT ASLEEP
; WHERE START NEXT SESSION SOON GOT GOING
; EARLY POST DELIVER CAME - TWICE AND 1ST OPEN DOOR NOT FOR LETTER BOX
; A FEW MANY
; -------------------------------------------------------------------
; TIME SPEND HERE TODAY FOR LAST SESSION
; -------------------------------------------------------------------
; Tue 17-Sep-2019 10:48:12
; Tue 17-Sep-2019 14:32:28 -- 10 MINUTE TO GET LAST BACK TRACK OF TIME
; -------------------------------------------------------------------
; FINISH
; -------------------------------------------------------------------
; Tue 17-Sep-2019 14:32:28
; Tue 17-Sep-2019 15:18:00 
; -------------------------------------------------------------------


RENAME_EXTENSION_SET_GO_EVERY_TIME=
O_LEN_STRING_HOLD_HWND=
RENAME_FILTER_EXTENSION_CASE_EXCLUDE=
RENAME_FILTER_EXTENSION_CASE_INCLUDE=
RENAME_EXTENSION_QUIET_WITH_AUDIO=
RENAME_EXTENSION_SET_DONE_QUIET=

GOSUB TIMER_RENAME_FILE_EXTENSION_CASE_UPPER_OR_LOWER
	
RETURN

TIMER_RENAME_FILE_EXTENSION_CASE_UPPER_OR_LOWER:

	RENAME_DONE=
	RENAME_DONE_COUNT=0
	
	; ---------------------------------------------------------------
	; LET SWITCH WITH AH BEGIN
	; ---------------------------------------------------------------
	
	SET_ARRAY_1:=[]
	SET_ARRAY_2:=[]
	SET_ARRAY_3:=[]
	SET_ARRAY_4:=[]
	SET_ARRAY_5:=[]
	SET_ARRAY_6:=[]
	ArrayCount:=0
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\DSC"
	SET_ARRAY_2[ArrayCount]:="MP4 MPG MPEG AVI MOV"        ; THESE ARE CASE WANTER END RESULT WILL BE AS HERE
	SET_ARRAY_3[ArrayCount]:=

	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="F:\DSC\2015+SONY"
	SET_ARRAY_2[ArrayCount]:="MP4 MPG MPEG AVI MOV"        ; THESE ARE CASE WANTER END RESULT WILL BE AS HERE
	SET_ARRAY_3[ArrayCount]:=
	
	
	; ArrayCount+=1
	; SET_ARRAY_1[ArrayCount]:="C:\SCRIPTER\"
	; SET_ARRAY_2[ArrayCount]:="vbp"
	; SET_ARRAY_3[ArrayCount]:="NOT TALK"

	
	DELETE_BACKUP_COUNT=0
	
	IF RENAME_FILTER_EXTENSION_CASE_INCLUDE
		RENAME_FILTER_EXTENSION_CASE_INCLUDE=% Space_1 RENAME_FILTER_EXTENSION_CASE_INCLUDE Space_1
	STRINGUPPER, RENAME_FILTER_EXTENSION_CASE_INCLUDE,RENAME_FILTER_EXTENSION_CASE_INCLUDE
	

	IF !RENAME_EXTENSION_QUIET_WITH_AUDIO
		Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	
	StringCaseSense, On
	
	NEST_2ND_LOOP_MAXIMUM=% SET_ARRAY_2.MaxIndex()

	; BEFORE LOOP SET HERE -- SetTitleMatchMode 2  ; PARTIAL PATH
	; ---------------------------------------------------------------
	SetTitleMatchMode 2  ; PARTIAL PATH

	; ------------------------------------------------------
	; FIND IF WIND ACTIVE FOR NOTEPAD++ AND HERE SCRIPT NAME
	; AND THEN IF SHOULD SHOW TOOLTIP
	; ------------------------------------------------------
	SET_SHOW_TOOLTIP=
	IF !RENAME_EXTENSION_SET_DONE_QUIET
	{
		FILE_ScriptName=%A_ScriptName%
		StringGetPos, pos, FILE_ScriptName, _, R
		IF pos
			FILE_ScriptName:=SubStr(FILE_ScriptName,1, pos)
		IfWinActive, %FILE_ScriptName% ahk_class Notepad++
			SET_SHOW_TOOLTIP=TRUE
	}
	
	; ---------------------------------------------------
	; WONT DO IF ARRAY NOT BEEN SET TO LEVEL 
	; SET TO 10 STOP ERROR MESSAGE AT 7
	; FIND THE NUMBER OF HOW MANY ROW TO THE ARRAY 1 TO 5
	; ---------------------------------------------------
	Loop, 3
	{
	 IF SET_ARRAY_%A_Index%
		NEST_2ND_LOOP_MAXIMUM_NUMBER_ROW_ARRAY=%A_Index%
	}
	
	COMPARE_TO_FIND=
	
	Loop, % SET_ARRAY_1.MaxIndex()
	{
	
		PATH_1:=SET_ARRAY_1[A_Index]
		NOT_TALK:=SET_ARRAY_3[A_Index]
		
		; MAKE PATH HAVE RIGHT END BACKSLASH IF NOT FITTER
		; -----------------------------------------------------------
		StringRight, CHECKVAR, PATH_1, 1
		IF (CHECKVAR<>"\")	
			PATH_1=%PATH_1%\

		XV:=INSTR(PATH_1,"\", FALSE, 1,2)
		NewStr_1:=SubStr(PATH_1,1,XV-1)
		NewStr_2:=SubStr(PATH_1,XV)
		NEW_FOLDER_1=% NewStr_1 "_YT_1" NewStr_2
		NEW_FOLDER_2=% NewStr_1 "_YT_2" NewStr_2

		IfNOTExist, %NEW_FOLDER_2%
			FileCreateDir, %NEW_FOLDER_2%
		
		IF SET_SHOW_TOOLTIP
			TOOLTIP % "Loop SET_ARRAY_1 MAX" "`n" SET_ARRAY_1.MaxIndex() "`n" "Loop SET_ARRAY_1_INDEX" "`n" A_INDEX "`n" SET_ARRAY_2[A_Index]

		EXT_COMP_1:=SET_ARRAY_2[A_Index]

		; SPLITTER 
		; ---------------------------------------------------------------
		IfExist, %PATH_1%
		Loop, parse, EXT_COMP_1, " "
		{
			EXT_COMP_GET:=A_LoopField
			
			STRINGUPPER EXT_COMP_UP, EXT_COMP_GET

			; -----------------------------------------------------------
			; SPLITTER
			; -----------------------------------------------------------
			EXCLUDE_NOT_GO=
			IF RENAME_FILTER_EXTENSION_CASE_EXCLUDE
				Loop, parse, RENAME_FILTER_EXTENSION_CASE_EXCLUDE, " "
				IF A_LoopField=%EXT_COMP_UP%
					EXCLUDE_NOT_GO=TRUE
			; -----------------------------------------------------------
			; IF NOTHING %EXCLUDE_NOT_GO% THEN OPERATE ALL 
			; EXCEPT INCLUDE REQUIRE
			; DON'T USER EXCLUDE WITH INCLUDE OR INCLUDE MEAN NONE
			; -----------------------------------------------------------
			
			; -----------------------------------------------------------
			; INCLUDE_TAKE_ACTION_AND_EXLCUDE_LEFT_NOT
			; -----------------------------------------------------------
			IF RENAME_FILTER_EXTENSION_CASE_INCLUDE
				EXCLUDE_NOT_GO=
			SET_INCLUDE_GO=
			IF !RENAME_FILTER_EXTENSION_CASE_EXCLUDE
			IF RENAME_FILTER_EXTENSION_CASE_INCLUDE
				Loop, parse, RENAME_FILTER_EXTENSION_CASE_INCLUDE, " "
				IF A_LoopField=%EXT_COMP_UP%
					SET_INCLUDE_GO=TRUE
			IF !RENAME_FILTER_EXTENSION_CASE_INCLUDE
				SET_INCLUDE_GO=TRUE
			; -----------------------------------------------------------

			
			IF SET_SHOW_TOOLTIP
				TOOLTIP % "Loop SET_ARRAY_1 MAX" "`n" SET_ARRAY_1.MaxIndex() "`n" "Loop SET_ARRAY_1_INDEX" "`n" A_INDEX "`n" SET_ARRAY_2[A_Index] "`n" "EXCLUDE_NOT_GO" "`n" EXCLUDE_NOT_GO  "`n" "SET_INCLUDE_GO" "`n"  SET_INCLUDE_GO

			IF SET_INCLUDE_GO  ; ---- IF NOTHING %SET_INCLUDE_GO% THEN NOT OPERATE
			IF !EXCLUDE_NOT_GO
			IF EXT_COMP_GET
			{
				PATH_NAME_2=% PATH_1 "*." EXT_COMP_GET
				; *C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 84-RENAME_FILE_EXTENSION_CASE_INCLUDE.ahk - Notepad++ [Administrator]
				
				FILE_COUNTER_MAX=
				Loop, Files, %PATH_NAME_2% ,R
					FILE_COUNTER_MAX=%A_INDEX%
				IF !FILE_COUNTER_MAX
					FILE_COUNTER_MAX=0
				
				IF !RENAME_EXTENSION_SET_DONE_QUIET
				IF !NOT_TALK
				IF SET_SHOW_TOOLTIP
					TOOLTIP % "FILE_COUNTER_MAX" "`n" FILE_COUNTER_MAX "`n" "FILE_COUNTER_INDEX" "`n" FILE_COUNTER_INDEX "`n"  A_Index "`n" PATH_NAME_2 

				Loop, Files, %PATH_NAME_2% ,R
				{
					PATH_NAME_3=%A_LoopFileFullPath%
					FILE_COUNTER_INDEX=%A_INDEX%

					IF SET_SHOW_TOOLTIP
					IF Mod(A_INDEX, 100)=0 
					{
						TOOLTIP % "FILE_COUNTER_MAX" "`n" FILE_COUNTER_MAX "`n" "FILE_COUNTER_INDEX" "`n" FILE_COUNTER_INDEX "`n" A_Index "`n" PATH_NAME_2 
						Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
					}
					
					SplitPath, PATH_NAME_3, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
					PATH_EXT_NAME_4=%OutExtension%
					PATH_NAME_4_REPLACE=% OutDir "\" OutFileName
					
					XV:=INSTR(PATH_NAME_4_REPLACE,"\", FALSE, 1,2)
					NewStr_1:=SubStr(PATH_NAME_4_REPLACE,1,XV-1)
					NewStr_2:=SubStr(PATH_NAME_4_REPLACE,XV)
					NEW_FOLDER_1=% NewStr_1 "_YT_1" NewStr_2
					NEW_FOLDER_2=% NewStr_1 "_YT_2" NewStr_2
					NEW_FOLDER_4=% "D:\DSC_YT_1\"
					
					
					; TOOLTIP % "FILE_1" "`n" PATH_NAME_4_REPLACE "`n" "FILE_2" "`n" NEW_FOLDER_1
					
					; -----------------------------------------------
					; 4096 BRING TO FOREGROUND
					; 4100 BRING TO FOREGROUND + YES NOT
					; FOREGROUND + YES NOT
					; 4096 + 4
					; -----------------------------------------------
					; MSGBOX ,4100,,% "RENAME TO HAPPEN UPPER CASE .MP4 EXTENSION`n`n"
					; IFMSGBOX YES
					; -----------------------------------------------
					IF !NOT_TALK
						; MSGBOX, 4096,,%  FILE_COUNTER_INDEX " OF " FILE_COUNTER_MAX "`n`n" "RENAME TO HAPPEN CASE CONVERT _.___ EXTENSION`n`n" A_LoopFileFullPath "`n`n" PATH_NAME_4_REPLACE
					
					RENAME_DONE=% RENAME_DONE PATH_NAME_4_REPLACE "`n" NEW_FOLDER_1 "`n" 
					RENAME_DONE_COUNT+=1
					COMPARE_TO_FIND=% COMPARE_TO_FIND PATH_NAME_4_REPLACE "`n"
					
					SplitPath, PATH_NAME_4_REPLACE, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
					
					IFNOTEXIST, %OutDir%
						MSGBOX % "DEBUGGER ERROR _ RESULT RENAME TARGET FOLDER NOT EXIST AS SHOWN`n`nTARGET FOLDER`n-------------_-----`n" OutDir "\`n`nSUPPOSE FILENAME TARGET WHEN RENAME`n-----------------------------------_---------------`n" PATH_NAME_4_REPLACE
						
					; IF FILE_COUNTER_INDEX>180 THEN 
					; BREAK

					SplitPath, NEW_FOLDER_1, OutFileName, OutDir

					NEW_FOLDER_1_ADD=% NEW_FOLDER_1_ADD "`n" NEW_FOLDER_1
					
					IF INSTR(PATH_NAME_ADD,NEW_FOLDER_4)=0
					IfNOTExist, %NEW_FOLDER_4%
						FileCreateDir, %NEW_FOLDER_4%
					PATH_NAME_ADD=% PATH_NAME_ADD "----" NEW_FOLDER_4
						
						
					; FILEMOVE, %A_LoopFileFullPath%,%PATH_NAME_4_REPLACE%
					PATH_NAME_5=% OutDir "\" OutNameNoExt ".lnk"
					PATH_NAME_7=% NEW_FOLDER_1 ".lnk"
					PATH_NAME_7=% NEW_FOLDER_4 OutNameNoExt ".lnk"
					
					PATH_NAME_5_ADD=% PATH_NAME_5_ADD "`n" PATH_NAME_5
					PATH_NAME_7_ADD=% PATH_NAME_5_ADD "`n" PATH_NAME_7

					
					; MSGBOX % PATH_NAME_7

					; FOR THE SUB TREE STRUCTURE
					; -----------------------------------------------
					IfNOTExist, %PATH_NAME_5%
						FileCreateShortcut, %PATH_NAME_4_REPLACE%, %PATH_NAME_5%

					; WITHOUT SUB TREE STRUCTURE AND SINGLE FOLDER
					; -----------------------------------------------
					IfNOTExist, %PATH_NAME_7%
						FileCreateShortcut, %PATH_NAME_4_REPLACE%, %PATH_NAME_7%
				}
			}
		}
	}


	; NOW DELETE WHAT NOT WANTER
	; ---------------------------------------------------
	; ---------------------------------------------------
	; ---------------------------------------------------
	; ---------------------------------------------------
	; ---------------------------------------------------
	NEW_FOLDER_1_ADD=
	Loop, % SET_ARRAY_1.MaxIndex()
	{
		PATH_1:=SET_ARRAY_1[A_Index]
		NOT_TALK:=SET_ARRAY_3[A_Index]

		StringRight, CHECKVAR, PATH_1, 1
		IF (CHECKVAR<>"\")	
			PATH_1=%PATH_1%\

		EXT_COMP_1:=SET_ARRAY_2[A_Index]

		IfExist, %PATH_1%
		Loop, parse, EXT_COMP_1, " "
		{
			EXT_COMP_GET:=A_LoopField
			STRINGUPPER EXT_COMP_UP, EXT_COMP_GET
			SET_INCLUDE_GO=
			IF RENAME_FILTER_EXTENSION_CASE_INCLUDE
				Loop, parse, RENAME_FILTER_EXTENSION_CASE_INCLUDE, " "
					IF A_LoopField=%EXT_COMP_UP%
						SET_INCLUDE_GO=TRUE
			IF !RENAME_FILTER_EXTENSION_CASE_INCLUDE
				SET_INCLUDE_GO=TRUE
			IF SET_INCLUDE_GO                           ; ---- IF NOTHING %SET_INCLUDE_GO% THEN NOT OPERATE
			IF EXT_COMP_GET
			{
				NEW_FOLDER_1=% "D:\DSC_YT_1\"
				Loop, Files, %NEW_FOLDER_1% ,R
				{
					PATH_NAME_3=%A_LoopFileFullPath%
					; -----------------------------------------------
					; WHAT READ FOR SCAN TARGET FOLDER %PATH_NAME_3% 
					; IF NOT IN SOURCE FOLDER BY ONLY CHECK FILENAME NOT PATH TREE
					; -----------------------------------------------
					IF INSTR(COMPARE_TO_FIND,%A_LoopFileName%)=0
					{
						MSGBOX % "DELETE `n" PATH_NAME_3
						; FileDelete, PATH_NAME_3
						DELETE_BACKUP_COUNT+=1
					}

					IF SET_SHOW_TOOLTIP
					IF Mod(A_INDEX, 100)=0 
					{
						TOOLTIP % "FILE_COUNTER_MAX" "`n" FILE_COUNTER_MAX "`n" "FILE_COUNTER_INDEX" "`n" FILE_COUNTER_INDEX "`n" A_Index "`n" PATH_NAME_2 
						Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
					}
				}
			}
		}
	}
	
	; IF NOTHING THEN ACTOR
	; ---------------------------------------------------------------
	VAR_SUBSTR:=SUBSTR(RENAME_DONE,1,1000)
	IF RENAME_DONE
		MSGBOX, 4096,, % "RENAME DO COUNTER =`n" RENAME_DONE_COUNT "`n" VAR_SUBSTR "`n" "`n" "DELETE_OF_SYNC_DO =" "`n" DELETE_BACKUP_COUNT "`n" "NOT ANY DELETE YET REQUIRE CHECK AND REM LINE"

	StringCaseSense, OFF ; PUT BACK TO DEFAULT I WONDER

	SetTitleMatchMode=

	TOOLTIP

	Soundplay, %a_scriptDir%\NOT EXIST FILE.wav
	
RETURN

