; =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE.ahk
;# __ 
;# __ Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
; =============================================================

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
	
	SetTitleMatchMode 2  ; Specify Full path
	; ---------------------------------------------------------------

	FN_Array_1 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="- Application Error"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="MSGBOX COUNTDOWN DELAY.ahk"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="MSGBOX COUNTDOWN DELAY_02.ahk"

	FN_Array_2 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_2[ArrayCount]:="MSGBOX COUNTDOWN DELAY.ahk"
	ArrayCount += 1
	FN_Array_2[ArrayCount]:="MSGBOX COUNTDOWN DELAY_02.ahk"

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

	
	SET_GO_GS=FALSE
	Loop % FN_Array_1.MaxIndex()
	{
		VAR_IN_NAME_1:=FN_Array_1[A_Index]
		IFWINEXIST %VAR_IN_NAME_1% ahk_class #32770
		{
			SET_GO_GS=TRUE
			VAR_IN_NAME_4=%VAR_IN_NAME_1%
			; TOOLTIP %VAR_IN_NAME_4% " -- " %VAR_IN_NAME_1%
		}
	}

	IF SET_GO_GS=TRUE
	{
		SET_GO_GS=FALSE
		WinGetTitle, OutputVar_1, %VAR_IN_NAME_4% ahk_class #32770
		Loop % FN_Array_2.MaxIndex()
		{
			VAR_IN_NAME_2:=FN_Array_2[A_Index]
			IF INSTR(OutputVar_1,VAR_IN_NAME_2)>0 
			{
				SET_GO_GS=TRUE
				RELAUNCH_PATH_VAR:="D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
			}
		}
	}

	IF SET_GO_GS=TRUE
	{
		ControlGettext, MSGBOX_INFO, Static2, %VAR_IN_NAME_4% ahk_class #32770
		SET_GO_02=FALSE
		Loop % FN_Array_3.MaxIndex()
		{
			VAR_IN_NAME_3:=FN_Array_3[A_Index]
			IF INSTR(MSGBOX_INFO,VAR_IN_NAME_3)>0
				SET_GO_02=TRUE
				
		}
		ControlGettext, MSGBOX_INFO, Static1, %VAR_IN_NAME_4% ahk_class #32770
		Loop % FN_Array_3.MaxIndex()
		{
			VAR_IN_NAME_3:=FN_Array_3[A_Index]
			IF INSTR(MSGBOX_INFO,VAR_IN_NAME_3)>0
				SET_GO_02=TRUE
				
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
			
			IF INSTR(CONTROL_TEXT_01,"OK")>0
			IF StrLen(CONTROL_TEXT_01)=4
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

