; =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
;# __ 
;# __ Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
; =============================================================

; -------------------------------------------------------------------
; SESSION UP 
; 
; Mon 26-Aug-2019 13:10:47
; Mon 26-Aug-2019 15:00:10
; -------------------------------------------------------------------

TIMER_VB_EXE_APPLICATION_ERROR_MSGBOX:

	; -------------------------------------------
	; ---------------------------
	; VB_KEEP_RUNNER.EXE - Application Error
	; ---------------------------
	; The application failed to initialize properly (0xc0000033). Click on OK to terminate the application. 
	; ---------------------------
	; OK   
	; ---------------------------
	; -------------------------------------------
	; D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe is not a valid Win32 application.
	; -------------------------------------------

	; IFWINEXIST Vb6 Loader ahk_exe Shell VBasic 6 Loader.exe
	
	SetTitleMatchMode 2  
	; ---------------------------------------------------------------

	
	; ---------------------------------------------------------------
	; ADD THE TITLE NEXT WILL BE USER MATCH FOR EXE WITH THAT THING
	; ---------------------------------------------------------------

	FN_Array_1 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="- Application Error"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="VB6\VB-NT\"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="MSGBOX COUNTDOWN DELAY.ahk"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="MSGBOX COUNTDOWN DELAY_02.ahk"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="ClipBoard Logger"

	; ---------------------------------------------------------------
	; WHEN SEARCH FOR EXE NAME IT USER TITLE FOR MATCH GO
	; NOT CASE SENSITIVE
	; ---------------------------------------------------------------
	FN_Array_2 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_2[ArrayCount]:="VB_KEEP_RUNNER.EXE"
	ArrayCount += 1
	FN_Array_2[ArrayCount]:="ClipBoard Logger"
	ArrayCount += 1
	FN_Array_2[ArrayCount]:="MSGBOX COUNTDOWN DELAY.ahk"
	ArrayCount += 1
	FN_Array_2[ArrayCount]:="MSGBOX COUNTDOWN DELAY_02.ahk"

	; ---------------------------------------------------------------
	; GIVE THE EXE NAME THAT WILL RELAUNCH WHEN WINDOW POPPED OUT THE WAY
	; THE INDEX NUMBER HAS TO MATCH THE ARRAY ABOVE FN_Array_2
	; ---------------------------------------------------------------
	FN_Array_4 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_4[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
	ArrayCount += 1
	FN_Array_4[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.EXE"
	ArrayCount += 1
	FN_Array_4[ArrayCount]:=""
	ArrayCount += 1
	FN_Array_4[ArrayCount]:=""

	; ---------------------------------------------------------------
	; GIVE SOME INFO THAT WILL BE IN THE STATIC1 OR STATIC2 CONTROL 
	; FIELD OF MSGBOX 
	; NOT CASE SENSITIVE - I LET YOU HAVE IT EASIER
	; ---------------------------------------------------------------
	FN_Array_3 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="The application failed to initialize properly" ; WIN 10
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="is not a valid Win32 application."             ; WIN XP
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="Failed attempt to launch program or document"  ; WIN XP ANOTHER TYPE OF BUG WHEN CORRUPTED EXE FILE -- THIS PROGRAM FOR COUNT DOWN MSGBOX SO NOT STUCK ON SCREEN AND ALLOW CONTINUE WHEN COPY OVER DO
	; STARVED OF DRINK WHILE THIS CODE AWKARD - BIT RUSTY AUTOHOTKEYS ARRAY COME ALONG NICELY
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="The application was unable to start correctly" ; WIN 07
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="Run-time error" ; WIN XP
	; ArrayCount += 1
	; FN_Array_3[ArrayCount]:="Object doesn't support the Property or method" ; WIN XP
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="Object doesn" ; WIN XP
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="support the Property or method" ; WIN XP
	; ---------------------------------------------------------------
	; LINE THAT HAVE ODD CHARACTER HERE DON'T WORK -- doesn't
	; Run-time error -- WAS THE ALTERNATIVE IN THIS PARTICULAR ONE
	; AND WITH SPLITTER
	; ---------------------------------------------------------------

	
	SET_GO_GS=FALSE
	Loop % FN_Array_1.MaxIndex()
	{
		VAR_IN_NAME_1:=FN_Array_1[A_Index]
		IFWINEXIST %VAR_IN_NAME_1% ahk_class #32770
		{
			SET_GO_GS=TRUE
			VAR_IN_NAME_4=%VAR_IN_NAME_1%
		}
	}

	; TEST DEBUG
;	IF SET_GO_GS=TRUE
;			TOOLTIP %VAR_IN_NAME_1%

	IF SET_GO_GS=TRUE
	{
		SET_GO_GS=FALSE
		WinGetTitle, OutputVar_1, %VAR_IN_NAME_4% ahk_class #32770
		Loop % FN_Array_2.MaxIndex()
		{
			; WHEN SEARCH FOR EXE NAME IT ONLY LOOK FOR TITLE
			; -----------------------------------------------
			VAR_IN_NAME_2:=FN_Array_2[A_Index]
			VAR_IN_NAME_4:=FN_Array_4[A_Index]
			StringUpper, VAR_IN_NAME_2, VAR_IN_NAME_2
			StringUpper, OutputVar_1, OutputVar_1
			IF INSTR(OutputVar_1,VAR_IN_NAME_2)>0 
			{
				SET_GO_GS=TRUE
				RELAUNCH_PATH_VAR=%VAR_IN_NAME_4%
			}
		}
	}

	IF SET_GO_GS=FALSE
		SHOW_COUNTDOWN_ACTION=FALSE
	
	; TEST DEBUG
	; IF SET_GO_GS=TRUE
		; TOOLTIP %SET_GO_GS%
		; TOOLTIP %RELAUNCH_PATH_VAR%
	
	IF SET_GO_GS=TRUE
	{
		; SEARCH STATIC2 
		; MOST COMMON -- BUT MSGBOX DO HAVE CHANGE SOMETIME
		; -----------------------------------------------------------
		ControlGettext, MSGBOX_INFO, Static2, %VAR_IN_NAME_4% ahk_class #32770
		SET_GO_02=FALSE
		Loop % FN_Array_3.MaxIndex()
		{
			VAR_IN_NAME_3:=FN_Array_3[A_Index]
			StringUpper, VAR_IN_NAME_3, VAR_IN_NAME_3
			StringUpper, MSGBOX_INFO, MSGBOX_INFO
			IF INSTR(MSGBOX_INFO,VAR_IN_NAME_3)>0
				SET_GO_02=TRUE
		}
		; SEARCH STATIC1
		; -----------------------------------------------------------
		IF SET_GO_02=FALSE
		{
			ControlGettext, MSGBOX_INFO, Static1, %VAR_IN_NAME_4% ahk_class #32770
			Loop % FN_Array_3.MaxIndex()
			{
				VAR_IN_NAME_3:=FN_Array_3[A_Index]
				StringUpper, VAR_IN_NAME_3, VAR_IN_NAME_3
				StringUpper, MSGBOX_INFO, MSGBOX_INFO
				IF INSTR(MSGBOX_INFO,VAR_IN_NAME_3)>0
					SET_GO_02=TRUE
			}
		}
		IF SET_GO_02=TRUE
		{
			ControlGetText CONTROL_TEXT_01,Button1,%VAR_IN_NAME_4% ahk_class #32770
			IF Secs_MSGBOX_08<1
			IF Secs_MSGBOX_08_RUN_ONCE=TRUE
			{	
				Secs_MSGBOX_08_RUN_ONCE=FALSE
				Secs_MSGBOX_08=0
				; NA [v1.0.45+]: May improve reliability. See reliability below.
				VAR_IN_NAME_8=%VAR_IN_NAME_4% ahk_class #32770
				LOOP, 1000
				{
					ControlClick, Button1,%VAR_IN_NAME_4%,,,, NA x10 y10 
					IFWinNOTExist %VAR_IN_NAME_8%
					{
						BREAK
					}
					SLEEP 500
				}
				SOUNDBEEP 4000,300
				VAR_DONE_ESCAPE_KEY=TRUE
				SLEEP 1000
				IfExist, %RELAUNCH_PATH_VAR%
				{
					Run, "%RELAUNCH_PATH_VAR%", , MIN
				}
			}
			
			; TOOLTIP % StrLen(CONTROL_TEXT_01)
			
			IF INSTR(CONTROL_TEXT_01,"OK")>0
			IF StrLen(CONTROL_TEXT_01)=4
			{
				Secs_MSGBOX_08=20
				SOUNDBEEP 5000,200
				SHOW_COUNTDOWN_ACTION=TRUE
			}

			IF INSTR(CONTROL_TEXT_01,"OK")>0
			IF StrLen(CONTROL_TEXT_01)=2
			{
				Secs_MSGBOX_08=20
				SOUNDBEEP 5000,200
				SHOW_COUNTDOWN_ACTION=TRUE
			}

			IF CONTROL_TEXT_01=OK
			{
				Secs_MSGBOX_08=20
				SOUNDBEEP 5000,200
				SHOW_COUNTDOWN_ACTION=TRUE
				Secs_MSGBOX_08_RUN_ONCE=TRUE
			}

			IF Secs_MSGBOX_08>0 	
				Secs_MSGBOX_08-=1

			IF Secs_MSGBOX_08<1
			IF Secs_MSGBOX_08_RUN_ONCE=FALSE
			{
				Secs_MSGBOX_08_RUN_ONCE=TRUE
				CONTROL_TEXT_03=%CONTROL_TEXT_01%X
				LOOP, 88
				{
					; SLEEP 500
					CONTROL_TEXT_02=OK  %A_INDEX%X
					IF INSTR(CONTROL_TEXT_03,CONTROL_TEXT_02)>0
					{
						Secs_MSGBOX_08=20 ; %A_INDEX%
						SHOW_COUNTDOWN_ACTION=TRUE
						BREAK
					}
				}
			}
			IF SHOW_COUNTDOWN_ACTION=TRUE
				ControlSetText,Button1,OK  %Secs_MSGBOX_08%, %VAR_IN_NAME_4% ahk_class #32770
		}
	}

RETURN

