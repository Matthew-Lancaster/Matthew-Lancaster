;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
;# __ 
;# __ Autokey -- 82-RENAME VB PROJECT VBP LOWER CASE.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; ONLY ONE THING 
; HERE IS LAUNCHER PART OF CODE
; ONLY TO ASK QUESTION CONFIRM 
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; HERE IS SIMILAR TO PROJECT DONE AFTER AND BETTER
; -------------------------------------------------------------------
; Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; TIMER_RENAME_EXTENSION_OF_MP4_TO_UPPER:
; LIKE HERE LINK BOTH IN CASE MORE PROGRESS TO BE DONE
;
; THIS ONE BEEN FINISH AND NOW HERE IS MADE WITHOUT USE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Autokey -- 75-GOODSYNC OPTIONS SET.ahk
; DETECT_VB_IDE_BEEN_RUN_MULTIPLE_AND_DO_TASK_AFTER:
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; FROM __ Tue 10-Sep-2019 18:40:50
; TO   __ Tue 10-Sep-2019 19:45:00
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; SESSION 02
; -------------------------------------------------------------------
; TOP UP ADD ARRAY AS CLOSE THE SESSION
; -------------------------------------------------------------------
; FROM __ Tue 10-Sep-2019 23:50:00
; TO   __ Tue 10-Sep-2019 23:50:00
; -------------------------------------------------------------------

#noEnv
; #persistent
#singleInstance force
detectHiddenWindows, on
setWorkingDir %a_scriptDir%
; #NoTrayIcon

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

SoundBeep , 1500 , 400
Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav

GOSUB GO_ROUTINE_2

RETURN

GO_ROUTINE_2:
	
	SET_ARRAY_1:=[]
	ArrayCount:=0
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="D:\VB6\VB-NT\"
	ArrayCount+=1
	SET_ARRAY_1[ArrayCount]:="C:\SCRIPTER\"
	
	StringCaseSense, On
		
	Loop, % SET_ARRAY_1.MaxIndex()
	{
		Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		PATH_1:=SET_ARRAY_1[A_Index]
		IfExist, %PATH_1%
		{
			PATH_NAME_2=%PATH_1%*.VBP
		
			Loop, Files, %PATH_NAME_2%,R
			{
				FileList = %FileList%%A_LoopFileName%`n
				VB_3=%A_LoopFileFullPath%
				VB_4=%A_LoopFileExt%
				IF VB_4=VBP
				{
					VB_5:=SubStr(VB_3,1,StrLen(VB_3)-3)"vbp"
					Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
					MSGBOX % VB_3 "`n`n" VB_5
					FILEMOVE, %A_LoopFileFullPath%,%VB_5%
				}
			}
		}
	}

RETURN




GO_ROUTINE:

	; VB_1:="D:\VB6\VB-NT\00_Best_VB_01\"
	VB_1:="D:\VB6\VB-NT\00_Best_VB_01\"

	; CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    ; IF INSTR(DIR(CODER_VBP_FILE_NAME_2,".VBP")
	; {
	SET_ARRAY_1 := []
	ArrayCount := 0
	I=%ArrayCount%
	I+=1
	VB_2=%VB_1%VB_KEEP_RUNNER\VB_KEEP_RUNNER.VBP
	SET_ARRAY_1[I]:=VB_2

	Loop, SET_ARRAY_1.MaxIndex()
	{
		VB_2:=SET_ARRAY_1[I]
		
		FileList =
		Loop, Files, %VB_2%
			FileList = %FileList%%A_LoopFileName%`n
		MSGBOX % FileList
	}	
EXITAPP
	
PAUSE
	

RETURN


; -------------------------------------------------------------------
; REFERENCE FIND CREDIT
; -------------------------------------------------------------------
; ----
; AHK BEST WAY TO KILL EXPLORER - Google Search
; https://www.google.com/search?q=AHK+BEST+WAY+TO+KILL+EXPLORER
; --------
; Restart Explorer (the official way) - Scripts and Functions - AutoHotkey Community
; https://autohotkey.com/board/topic/93906-restart-explorer-the-official-way/
; ----
