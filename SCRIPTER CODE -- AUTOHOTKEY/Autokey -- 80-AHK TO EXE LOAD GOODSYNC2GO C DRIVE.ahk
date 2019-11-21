;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
;# __ 
;# __ Autokey --80-AHK TO EXE LOAD GOODSYNC2GO D DRIVE.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; THESE ARE DUAL PAIR FOR GOODSYNC2GO C AND D DRIVE
; D DRIVE WAS FIRST AS WANTED ICON
; AND THEN REALISE THAT WANT FOR C DRIVE ALSO
; OTHER WISE CHANGE ICON IS DONE BY PROPERTY
; AND ALSO BAT2EXE WAS USER BEFORE
; BUT IT CALL UP DOS PROMPT FOR LITTLE WHILE ANNOYER
; -------------------------------------------------------------------
; FROM __ Wed 04-Sep-2019 09:28:00
; TO   __ Wed 04-Sep-2019 10:21:38
; TO   __ Wed 04-Sep-2019 12:58:00 LITTLE SPLIT INTO DUAL COPY
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

GOSUB GO_ROUTINE

EXITAPP

GO_ROUTINE:

	DRIVE_LETTER=
	DL=C
	
	DELETE_AND_RECREATE_COMPILED_EXE_FOR_AUTOSCRIPT=
	
	; REQUIRE TO FIND IF SCRIPT WANT RECOMPILE TO EXE
	; ---------------------------------------------------------------
	FileGetTime, MOD_DATE_SCRIPT_1, %a_scriptNAME%, M
	IfExist, %a_scriptDir%\DATE_OF_AHK_SCRIPT.txt
		FileRead, MOD_DATE_SCRIPT_2, %a_scriptDir%\DATE_OF_AHK_SCRIPT.txt
	
	IF INSTR(MOD_DATE_SCRIPT_2,MOD_DATE_SCRIPT_1)=0
	{
		FileDelete, %a_scriptDir%\DATE_OF_AHK_SCRIPT.txt
		FileAppend, %MOD_DATE_SCRIPT_1%, %a_scriptDir%\DATE_OF_AHK_SCRIPT.txt
		DELETE_AND_RECREATE_COMPILED_EXE_FOR_AUTOSCRIPT=TRUE
		; MSgBOX % MOD_DATE_SCRIPT_2 "`n" MOD_DATE_SCRIPT_1
	}
	; ---------------------------------------------------------------
	
	SoundBeep , 1500 , 400
	Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	
	; ---------------------------------------------------------------
	FN_VAR_TMP_FILE=C:\SCRIPTOR DATA\VB_KEEP_RUNNER_IS_GOODSYNC2GO_RUNNER\
	PATH_NAME_2=%FN_VAR_TMP_FILE%*.TXT
	
	PATH := % PATH_NAME_2 "`n`n"
	LOOP, %PATH_NAME_2%
	{
		Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		PATH := % PATH A_LoopFileName "`n"
		; -----------------------------------------------------------
		; IF DL=C AND THEN ONLY INTEREST IS C AND NOT D BOTH 
		; SAME COMPUTER
		; OTHERWISE IS D AND D NOT ANY COMPUTER
		; -----------------------------------------------------------
		IF DL=C
			IF INSTR(A_LoopFileName,A_ComputerName)>0
			IF INSTR(A_LoopFileName,"D_HDD_GOODSYNC2GO_RUNNER")>0
			NUMBER_V +=1
		IF DL=D
			NUMBER_V +=1
	}
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
	IF NUMBER_V>0
	{
		T1=VB_KEEP_RUNNER AND AHK
		T2=Autokey -- 80-AHK TO EXE LOAD GOODSYNC2GO D DRIVE.ahk
		T3=DETECT D_HDD_GOODSYNC2GO IS RUNNER
		T4=%PATH%
		T5=NOT TO RUN TWO
		T20=% T1 "`n`n" T2 "`n`n" T3 "`n`n" T4 "`n" T5
		MSGBOX % T20
	}
	; ---------------------------------------------------------------

	IF DL=D
		ICON_GOT=GoodSync-inst_155.ico
	IF DL=C
		ICON_GOT=GoodSync-inst_308.ico
	
	; DEBUG SOMETHING
	; GOODSYNC2GO.EXE -- FILENAME HAS BEEN CHANGER TO GOODSYNC.EXE
	; BY GOODSYNC COMPANY
	; ON THE C DRIVE VERSION
	; Mon 04-Nov-2019 06:47:10
	
	IF A_Is64bitOS=1
	FN_VAR_EXE=%DL%:\GOODSYNC\x64\GoodSync2Go.exe
	IF A_Is32bitOS=0
	FN_VAR_EXE=%DL%:\GOODSYNC\x86\GoodSync2Go.exe

	AttributeString := FileExist(FN_VAR_EXE)
	IF !AttributeString
	{
		StringReplace, FN_VAR_EXE, FN_VAR_EXE, \GoodSync2Go.exe,\GoodSync.exe, All
	}

	AttributeString := FileExist(FN_VAR_EXE)
	IF !AttributeString
	{
		MSGBOX % FN_VAR_EXE "`n`n NOT EXIST TO FIND AH" ; ,vbMsgBoxSetForeground 
		RETURN
	}

	IF !NUMBER_V
	{
		GOODSYNC2GO_D_RUN=TRUE
		WinGet Path, ProcessPath, ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}
		IF INSTR(Path,"%DL%:\GoodSync\x64\GoodSync2Go.exe")
			GOODSYNC2GO_D_RUN=TRUE

		IF GOODSYNC2GO_D_RUN=TRUE
			Run, %FN_VAR_EXE% ; ,,MIN
		
		SoundBeep , 2500 , 100
		Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	}
	
	IF INSTR(a_scriptNAME,"AHK.EXE")
		RETURN
	
	SCRIPT_1_IN_2=%a_scriptNAME%
	STRINGUPPER, SCRIPT_1_IN_2,SCRIPT_1_IN_2
	; --------------------------------------------------------------------
	; "C:\GOODSYNC JOB CONTROL\GoodSync2Go AHK2EXE - C\AUTOKEY -- 80-AHK TO EXE LOAD GOODSYNC2GO C DRIVE.AHK.EXE"
	; "D:\GOODSYNC JOB CONTROL\GoodSync2Go AHK2EXE - D\AUTOKEY -- 80-AHK TO EXE LOAD GOODSYNC2GO D DRIVE.AHK.EXE"
	; --------------------------------------------------------------------
	SCRIPT_2_OUT_=%DL%:\GOODSYNC JOB CONTROL\GoodSync2Go AHK2EXE - %DL%\%SCRIPT_1_IN_2%.EXE
	FN_VAR_EXE=%SCRIPT_2_OUT_%
	
	IF DELETE_AND_RECREATE_COMPILED_EXE_FOR_AUTOSCRIPT=TRUE
		FileDelete, %FN_VAR_EXE%
	
	IfNOTExist, %FN_VAR_EXE%
	{
		FN_VAR_FOLDER=%DL%:\GOODSYNC JOB CONTROL\GoodSync2Go AHK2EXE - %DL%
		IfExist, %FN_VAR_FOLDER%
		{
			FN_VAR_EXE=C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe
			IfExist, %FN_VAR_EXE%
			{
				SCRIPT_1_IN__="%A_ScriptFullPath%"
				SCRIPT_2_OUT_="%SCRIPT_2_OUT_%"
				SCRIPT_3_ICON="%DL%:\GOODSYNC JOB CONTROL\GoodSync2Go AHK2EXE - %DL%\%ICON_GOT%"
				
				Run, %FN_VAR_EXE% /in %SCRIPT_1_IN__% /out %SCRIPT_2_OUT_% /icon %SCRIPT_3_ICON%
				
				SoundBeep , 2500 , 100
				Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
			}
		
		}
	}

RETURN


