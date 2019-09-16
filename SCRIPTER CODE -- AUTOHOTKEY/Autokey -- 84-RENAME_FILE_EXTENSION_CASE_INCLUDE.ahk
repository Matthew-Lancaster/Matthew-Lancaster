


; -------------------------------------------------------------------
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
; Autokey -- 78-RENAME_FILE_EXTENSION_CASE_INCLUDE.ahk

; THE INCLUDE FILE TO SHARE BETWEEN OTHER CODE
; HERE 
; 1.. Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; 2.. Autokey -- 78-RENAME_FILE_EXTENSION_CASE_02.ahk
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; [ Monday 22:08:00 Pm_16 September 2019 ]
; -------------------------------------------------------------------



TIMER_RENAME_FILE_EXTENSION_CASE_UPPER_OR_LOWER:

	; ---------------------------------------------------------------
	; WOULD OF BEEN EASIER USE THESE INSTEAD
	; ---------------------------------------------------------------
	; NICE CODE PLAY IDEA
	; ---------------------------------------------------------------
	; MAYBE A QUICKER TURN AROUND 
	; IF ONLY CHECK ACTIVE WINDOW
	; ---------------------------------------------------------------
	; WinGet, WNDCLASS, List, ahk_class wndclass_desked_gsk
	; WinGet, WNDCLASS, List, ahk_class CabinetWClass
	; ---------------------------------------------------------------

	RENAME_DONE=
	RENAME_DONE_COUNT=0
	
	; ---------------------------------------------------------------
	; THE BLOCK CHUNK CODE HERE OF WINGET
	; WILL DETECT ANY CLASS BEEN LOAD AND WHEN GONE
	; MAINLY THE GONE BIT
	; ---------------------------------------------------------------
	WinGet, TIMER_RENAME_EXT_MP4_HWND, ID, A
	WinGetCLASS, TIMER_RENAME_EXT_MP4_HWND_CLASS, A
	; ---------------------------------------------------------------
	; CLASS AND HWND TOGETHER
	; GOING TO BE SPLIT INTO AN ARRAY
	; ---------------------------------------------------------------
	TIMER_RENAME_EXT_MP4=% TIMER_RENAME_EXT_MP4_HWND "`n" TIMER_RENAME_EXT_MP4_HWND_CLASS

	; ---------------------------------------------------------------
	SET_THE_WINDOW_CLASSER=
	IF INSTR(TIMER_RENAME_EXT_MP4,"CabinetWClass")>0          ; EXPLORER CLASS
	SET_THE_WINDOW_CLASSER=TRUE
	IF INSTR(TIMER_RENAME_EXT_MP4,"wndclass_desked_gsk")>0    ; VB6 IDE CLASS
	SET_THE_WINDOW_CLASSER=TRUE
	; ---------------------------------------------------------------
	
	IF SET_THE_WINDOW_CLASSER=TRUE
	IF INSTR(STRING_HOLD_RENAME_EXT_MP4,TIMER_RENAME_EXT_MP4)=0
		STRING_HOLD_RENAME_EXT_MP4=% STRING_HOLD_RENAME_EXT_MP4 "`n" TIMER_RENAME_EXT_MP4
	
	; SPLIT ARRAY
	STRING_HOLD_HWND_2=
	Loop, parse, STRING_HOLD_RENAME_EXT_MP4, `n
	{
		IF A_LoopField
		IF A_LoopField<>CabinetWClass
		{
		WinGet STRING_HOLD_HWND_1,ID, ahk_id %A_LoopField%
		IF STRING_HOLD_HWND_1>0
			STRING_HOLD_HWND_2=% STRING_HOLD_HWND_2 A_LoopField "`n" 
		}
	}
	IF !O_LEN_STRING_HOLD_HWND
		O_LEN_STRING_HOLD_HWND=% STRLEN(STRING_HOLD_HWND_2)

	SET_GO_RENAME_EXT_MP4=
	LEN_STRING_HOLD_HWND_1:=STRLEN(STRING_HOLD_HWND_2)
	IF LEN_STRING_HOLD_HWND_1<%O_LEN_STRING_HOLD_HWND%
		SET_GO_RENAME_EXT_MP4=TRUE
	
	O_LEN_STRING_HOLD_HWND=%LEN_STRING_HOLD_HWND_1%
	
	; ---------------------------------------------------------------
	; DEBUG RUN ALWAYS REPEATER
	; ---------------------------------------------------------------
	; SET_GO_RENAME_EXT_MP4=TRUE
	
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	; ALL CHECKER READY TO GO
	; ---------------------------------------------------------------
	IF !SET_GO_RENAME_EXT_MP4
		RETURN
	; ---------------------------------------------------------------
	; LET SWITCH WITH AH BEGIN
	; ---------------------------------------------------------------
	
	SET_ARRAY_1:=[]
	SET_ARRAY_2:=[]
	SET_ARRAY_3:=[]
	SET_ARRAY_4:=[]
	ArrayCount:=0
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\DSC\" ; 2015+SONY\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\VI_ DSC ME\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:VI_ DSC ME 02\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\VIDEO\NOT\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\VIDEO\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\NOKIA E72\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\0 00 VIDEO 01\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\0 00 VIDEO 03 DOWNLOAD\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\0 00 VIDEO CCSS\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\0 00 VIDEO CCSS 02\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\0 00 VIDEO DOWNLOAD\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\0 00 MOBILE-1\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\0 00 MOBILE-2\"
	SET_ARRAY_2[ArrayCount]:="MP4"
	SET_ARRAY_3[ArrayCount]:=
	SET_ARRAY_4[ArrayCount]:="MPG"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\VB6\VB-NT\"
	SET_ARRAY_2[ArrayCount]:="vbp"
	SET_ARRAY_3[ArrayCount]:="NOT TALK"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="C:\SCRIPTER\"
	SET_ARRAY_2[ArrayCount]:="vbp"
	SET_ARRAY_3[ArrayCount]:="NOT TALK"


	; ---------------------------------------------------------------
	; FORMAT EXAMPLE -- WHAT TO EXCLUDE
	; ---------------------------------------------------------------
	; RENAME_FILTER_EXTENSION_CASE_EXCLUDE:="VBP MP4 MPG"
	; RENAME_FILTER_EXTENSION_CASE_EXCLUDE:="MP4 MPG"
	; ---------------------------------------------------------------
	RENAME_FILTER_EXTENSION_CASE_EXCLUDE=% " " RENAME_FILTER_EXTENSION_CASE_EXCLUDE " "
	STRINGUPPER, RENAME_FILTER_EXTENSION_CASE_EXCLUDE,RENAME_FILTER_EXTENSION_CASE_EXCLUDE
	
	
	; SET_ARRAY_1[ArrayCount]:="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\"
	; SET_ARRAY_1[ArrayCount]:="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION ARC\"

	Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	
	StringCaseSense, On
	
	Loop, % SET_ARRAY_1.MaxIndex()
	{
	B_Index=%A_Index%
	Loop, 2
	{

		IF A_Index=1
		EXT_COMP_GET:=SET_ARRAY_2[A_Index]
		IF A_Index=2
		EXT_COMP_GET:=SET_ARRAY_4[A_Index]

		; TOOLTIP % " -- " B_Index " -- " A_Index " -- " EXT_COMP_GET
		
		PATH_1:=SET_ARRAY_1[B_Index]
		NOT_TALK:=SET_ARRAY_3[B_Index]
		EXT_COMP_1=%EXT_COMP_GET%
		EXT_COMP_2:=% SUBSTR(EXT_COMP_GET,1,1)
		StringUPPER EXT_COMP_U_3,EXT_COMP_1
		StringUPPER EXT_COMP_U_2,EXT_COMP_2
		StringLOWER EXT_COMP_L_2,EXT_COMP_2
		
		; SPLITTER
		EXCLUDE_NOT_GO=
		Loop, parse, RENAME_FILTER_EXTENSION_CASE_EXCLUDE, " "
		{
		IF A_LoopField=%EXT_COMP_U_3%
			EXCLUDE_NOT_GO=TRUE
		}
		; IF NOTHING IN EXCLUDE THEN OPERATE
		IF !EXCLUDE_NOT_GO
		IF EXT_COMP_1
		IfExist, %PATH_1%
		{
			TOOLTIP % " -- " B_Index " -- " A_Index " -- " EXT_COMP_GET " -- " PATH_1
			PATH_NAME_2=%PATH_1%*.%EXT_COMP_1%
			Loop, Files, %PATH_NAME_2% ,R
			{
				PATH_NAME_3=%A_LoopFileFullPath%
				PATH_EXT_NAME_4:=% SUBSTR(A_LoopFileExt,1,1)
				; ---------------------------------------------------
				; IF WANT UPPER RESULT AND THEN DETECT LOWER
				; AND OTHER WAY
				; ---------------------------------------------------
				IF EXT_COMP_2=%EXT_COMP_U_2%
				{
					MSGBOX_I:=% "_." EXT_COMP_U_3 " GO UPPER CASE"
					SET_MODE_OPPOSITE_CASE=%EXT_COMP_L_2%
				}	
				IF EXT_COMP_2=%EXT_COMP_L_2%
				{	
					MSGBOX_I:=% "_." EXT_COMP_U_3 " GO LOWER CASE"
					SET_MODE_OPPOSITE_CASE=%EXT_COMP_U_2%
				}
				
				IF PATH_EXT_NAME_4=%SET_MODE_OPPOSITE_CASE%
				{
					PATH_REPLACE_EXT_5:=SubStr(PATH_NAME_3,1,StrLen(PATH_NAME_3)-3)
					PATH_REPLACE_EXT_5=%PATH_REPLACE_EXT_5%%EXT_COMP_1%
					Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
					; -----------------------------------------------
					; 4096 BRING TO FOREGROUND
					; 4100 BRING TO FOREGROUND + YES NOT
					; FOREGROUND + YES NOT
					; 4096 + 4
					; -----------------------------------------------
					; MSGBOX ,4100,,% "RENAME TO HAPPEN UPPER CASE .MP4 EXTENSION`n`n"
					; IFMSGBOX YES
					; -----------------------------------------------
					; IF !NOT_TALK
						; MSGBOX, 4096,,% "RENAME TO HAPPEN CASE CONVERT _.___ EXTENSION`n`n"MSGBOX_I "`n`n" A_LoopFileFullPath "`n`n" PATH_REPLACE_EXT_5
					
					RENAME_DONE=% RENAME_DONE PATH_REPLACE_EXT_5 "`n"
					RENAME_DONE_COUNT+=1
					
					FILEMOVE, %A_LoopFileFullPath%,%PATH_REPLACE_EXT_5%
				}
			}
		}
	}
	}
	
	; IF NOTHING THEN ACTOR
	; ---------------------------------------------------------------
	IF !RENAME_DONE_QUIET
	IF RENAME_DONE
		MSGBOX, 4096,, % "RENAME DONE COUNTER =`n" RENAME_DONE_COUNT "`n" RENAME_DONE

	StringCaseSense, OFF ; PUT BACK TO DEFAULT I WONDER

RETURN

