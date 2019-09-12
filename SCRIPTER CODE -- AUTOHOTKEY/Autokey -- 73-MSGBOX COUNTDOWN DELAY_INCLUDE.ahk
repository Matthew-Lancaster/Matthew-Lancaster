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

 
; -------------------------------------------------------------------
; SESSION 002 
; -------------------------------------------------------------------
; BETWEEN THE INCLUDE CODE HERE AND THE PROJECT INCLUDE WITHER
; TAKE ME ALL DAY
; MANY ARRAY ADDED FOR YOUR USER REQUIRING
; BUT STILL NOT YET AND EASIER TO DO
; PUT THAT STATIC1 AND AND STATIC2 HAS MORE ARRAY TO MULTIPLE SEARCH
; VERY NECESSARY TO REQUIRE 2 THERE
; AS NOT ALLOWED TO INCLUDE FULL STOP 
; 
; USEFUL ROUTINE
; -------------------------------------------------------------------
; Fri 30-Aug-2019 10:50:51
; Fri 30-Aug-2019 15:50:00 -- 10 TILL 4 -- SIX HOUR
; -------------------------------------------------------------------


; ' Sub ICACLS_ERROR_APPLYING_SECURITY()
; ' -------------------------------------------------------------------
; ' "Error Applying Security"
; ' THIS CODE IS OF -- D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe
; ' -------------------------------------------------------------------


IF INSTR(A_ScriptName,"_INCLUDE.ahk")>0 
	EXITAPP


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
	FN_Array_1[ArrayCount]:="Application Error"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="VB6\VB-NT\"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="MSGBOX COUNTDOWN DELAY.ahk"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="MSGBOX COUNTDOWN DELAY_02.ahk"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="ClipBoard Logger"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="Vb6 Loader"
	ArrayCount += 1
	FN_Array_1[ArrayCount]:="Microsoft Visual Basic" ; MSCOMCTL.OCX' could not be loaded
	
	; MESSENGER BOX ANSWER TEXT -- LOOK ABOVE -- FN_Array_1
	FN_Array_5_1 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_5_1[ArrayCount]:="OK"
	ArrayCount += 1
	FN_Array_5_1[ArrayCount]:="OK"
	ArrayCount += 1
	FN_Array_5_1[ArrayCount]:="OK"
	ArrayCount += 1
	FN_Array_5_1[ArrayCount]:="OK"
	ArrayCount += 1
	FN_Array_5_1[ArrayCount]:="OK"
	ArrayCount += 1
	FN_Array_5_1[ArrayCount]:="OK"
	ArrayCount += 1
	FN_Array_5_1[ArrayCount]:="No" ; "Microsoft Visual Basic"

	; MESSENGER BOX ANSWER CONTROL NAME -- LOOK ABOVE -- FN_Array_1 & FN_Array_5_1
	FN_Array_5_2 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_5_2[ArrayCount]:="Button1"
	ArrayCount += 1
	FN_Array_5_2[ArrayCount]:="Button1"
	ArrayCount += 1
	FN_Array_5_2[ArrayCount]:="Button1"
	ArrayCount += 1
	FN_Array_5_2[ArrayCount]:="Button1"
	ArrayCount += 1
	FN_Array_5_2[ArrayCount]:="Button1"
	ArrayCount += 1
	FN_Array_5_2[ArrayCount]:="Button1"
	ArrayCount += 1
	FN_Array_5_2[ArrayCount]:="Button2" ; "Microsoft Visual Basic"

	; COUNTER START VALUE
	FN_Array_5_3 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_5_3[ArrayCount]:=20
	ArrayCount += 1
	FN_Array_5_3[ArrayCount]:=20
	ArrayCount += 1
	FN_Array_5_3[ArrayCount]:=20
	ArrayCount += 1
	FN_Array_5_3[ArrayCount]:=20
	ArrayCount += 1
	FN_Array_5_3[ArrayCount]:=20
	ArrayCount += 1
	FN_Array_5_3[ArrayCount]:=8
	ArrayCount += 1
	FN_Array_5_3[ArrayCount]:=8
	
	; KILL AFTER COUNTDOWN
	FN_Array_5_4 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_5_4[ArrayCount]:=
	ArrayCount += 1
	FN_Array_5_4[ArrayCount]:=
	ArrayCount += 1
	FN_Array_5_4[ArrayCount]:=
	ArrayCount += 1
	FN_Array_5_4[ArrayCount]:=
	ArrayCount += 1
	FN_Array_5_4[ArrayCount]:=
	ArrayCount += 1
	FN_Array_5_4[ArrayCount]:=
	ArrayCount += 1
	FN_Array_5_4[ArrayCount]:="TRUE"
	
	; ---------------------------------------------------------------
	; WHEN SEARCH FOR EXE NAME IT USER TITLE FOR MATCH ABOVE 
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
	ArrayCount += 1
	FN_Array_2[ArrayCount]:=""
	ArrayCount += 1
	FN_Array_2[ArrayCount]:="" ; Shell VBasic 6 Loader.exe"
	ArrayCount += 1
	FN_Array_2[ArrayCount]:=""
	

	; ---------------------------------------------------------------
	; GIVE THE EXE NAME THAT WILL RELAUNCH WHEN WINDOW POPPED OUT THE WAY
	; THE INDEX NUMBER HAS TO MATCH THE ARRAY ABOVE FN_Array_2 NOT CORRECT 
	; ANSWER FN_Array_1
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
	ArrayCount += 1
	FN_Array_4[ArrayCount]:=""
	ArrayCount += 1
	FN_Array_4[ArrayCount]:=""
	ArrayCount += 1
	FN_Array_4[ArrayCount]:="D:\VB6\VB-NT\00_Best_VB_01\Shell VBasic 6 Loader\Shell VBasic 6 Loader.exe"
	FN_Array_4[ArrayCount]:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 58-VB6 CORRECT MSCOMCTL.OCX _2.2_ _TO_ _2.1_.VBS"

	; ---------------------------------------------------------------
	; GIVE SOME INFO THAT WILL BE IN THE STATIC1 OR STATIC2 CONTROL 
	; FIELD OF MSGBOX 
	; NOT CASE SENSITIVE - I LET YOU HAVE IT EASIER
	; DON'T INCLUDE DOT
	; ---------------------------------------------------------------
	FN_Array_3 := []
	ArrayCount := 0
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="The application failed to initialize properly" ; WIN 10
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="is not a valid Win32 application"             ; WIN XP
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
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="The application was unable to start" ; WIN 07 STARTER
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="#2.2# -- MSCOMCTL.OCX" ; WIN 10 "Vb6 Loader"
	; IT FIND HERE AND THEN HAS QUESTION APPLICATION TO RUN HERE ABOVE - FN_Array_4
	; #2.2# -- MSCOMCTL.OCX
	; WRONG VERSION -- AUTO CHANGED TO
	; #2.1# -- MSCOMCTL.OCX
	ArrayCount += 1
	FN_Array_3[ArrayCount]:="could not be loaded--Continue Loading Project" ; WIN 10  ; MSCOMCTL.OCX' could not be loaded--Continue Loading

	; ArrayCount += 1
	; FN_Array_3[ArrayCount]:="MSCOMCTL" ; WIN 10 "Vb6 Loader"
	; ---------------------------------------------------------------
	; LINE THAT HAVE ODD CHARACTER HERE DON'T WORK -- doesn't
	; Run-time error -- WAS THE ALTERNATIVE IN THIS PARTICULAR ONE
	; AND WITH SPLITTER
	; ---------------------------------------------------------------
	
	; ---------------------------------------------------------------
	; ARRAY NOT TO DO CROSS MIX ANYMORE
	; EACH ARRAY INDEX ONLY
	; IF MULTIPLE SEARCH TERM OPEN NEW ARRAY
	; ---------------------------------------------------------------
	
	SET_GO_GS=FALSE
	Loop % FN_Array_1.MaxIndex()
	{
		VAR_IN_NAME_1:=FN_Array_1[A_Index]
		IFWINEXIST %VAR_IN_NAME_1% ahk_class #32770
		{
			WINGET, VAR_HWND_10, ID, %VAR_IN_NAME_1% ahk_class #32770
			IF O_VAR_HWND_10<>%VAR_HWND_10%
			{			
				Secs_MSGBOX_08=0
				SHOW_COUNTDOWN_ACTION=FALSE
			}
			O_VAR_HWND_10=%VAR_HWND_10%
			
			SET_GO_GS=TRUE
			VAR_IN_NAME_4:=FN_Array_4[A_Index]     ; GET THE EXE WITH ARRAY POINTER
			VAR_IN_NAME_2:=FN_Array_2[A_Index]
			VAR_IN_NAME_5_1:=FN_Array_5_1[A_Index] ; BUTTON TEXT
			VAR_IN_NAME_5_2:=FN_Array_5_2[A_Index] ; BUTTON CONTROL NAME
			VAR_IN_NAME_5_3:=FN_Array_5_3[A_Index] ; COUNTER START VALUE
			VAR_IN_NAME_5_4:=FN_Array_5_4[A_Index] ; KILL AFTER COUNT DOWN
			RELAUNCH_PATH_VAR=%VAR_IN_NAME_4%
			INDEX_GO=%A_Index%
			SET_GO_GS=TRUE
			BREAK
		}
	}

	IF SET_GO_GS=FALSE
		SHOW_COUNTDOWN_ACTION=FALSE
	
	; TEST DEBUG
	; IF SET_GO_GS=TRUE
		
	IF SET_GO_GS=TRUE
	{
		; SEARCH STATIC2 AND THEN STATIC1
		; MOST COMMON -- BUT MSGBOX DO HAVE CHANGE SOMETIME
		; -----------------------------------------------------------
		ControlGettext, MSGBOX_INFO, Static2, %VAR_IN_NAME_1% ahk_class #32770
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
			ControlGettext, MSGBOX_INFO, Static1, %VAR_IN_NAME_1% ahk_class #32770
			Loop % FN_Array_3.MaxIndex()
			{
				VAR_IN_NAME_3:=FN_Array_3[A_Index]
				StringUpper, VAR_IN_NAME_3, VAR_IN_NAME_3
				StringUpper, MSGBOX_INFO, MSGBOX_INFO
				IF INSTR(MSGBOX_INFO,VAR_IN_NAME_3)>0
					SET_GO_02=TRUE
			}
		}      
		

		; READY TO GO PLAY WITH TIMER AT MSGBOX
		IF SET_GO_02=TRUE
		{

			; 1ST HAS TIMER BEEN SET BEFORE BY ANOTHER RUN OF CODE THAT WAS STOP
			; IF SO GRAB THAT TIME AND CONTINUE COUNTDOWN ON
			; WITH A CHOICE START FROM COUNT VALUE AGAIN EXAMPLE 20
			; -------------------------------------------------------

			ControlGetText CONTROL_TEXT_01,%VAR_IN_NAME_5_2%,%VAR_IN_NAME_1% ahk_class #32770

			; IF STRING WANT MATCH WHEN CHECK 1ST FEW CHAR
			IF INSTR(SubStr(CONTROL_TEXT_01,1,2),VAR_IN_NAME_5_1)>0
			; IF ORIGINAL STRING BEFORE CHANGE BEFORE START CODE
			IF INSTR(VAR_IN_NAME_5_1,CONTROL_TEXT_01)=0
			IF SHOW_COUNTDOWN_ACTION=FALSE
			IF Secs_MSGBOX_08<1
			{
				Secs_MSGBOX_08=%VAR_IN_NAME_5_3%
				SOUNDBEEP 5000,200
				SHOW_COUNTDOWN_ACTION=TRUE

				CONTROL_TEXT_03=%CONTROL_TEXT_01%X
				LOOP, 88
				{
					; IF A_INDEX=0 ; DO NOT HAPPEN START BEGIN AT 1
					; ---------------------------------------------
					CONTROL_TEXT_02=%VAR_IN_NAME_5_1%  %A_INDEX%X

					IF INSTR(CONTROL_TEXT_03,CONTROL_TEXT_02)>0
					{
						Secs_MSGBOX_08=20 ; %A_INDEX%
						Secs_MSGBOX_08=%A_INDEX%
						Secs_MSGBOX_08=%VAR_IN_NAME_5_3%
						SHOW_COUNTDOWN_ACTION=TRUE
						BREAK
					}
				}
			}


			; IF STRING WANT MATCH WHEN CHECK 1ST FEW CHAR
			IF INSTR(CONTROL_TEXT_01,%VAR_IN_NAME_5_1%)>0
			IF SHOW_COUNTDOWN_ACTION=FALSE
			IF Secs_MSGBOX_08<1
			{
				Secs_MSGBOX_08=%VAR_IN_NAME_5_3%
				SOUNDBEEP 5000,200
				SHOW_COUNTDOWN_ACTION=TRUE
			}
			
			; IS THE MSGBOX CONTROL IN NATURAL STATE AND REQUIRE NUMBER GO 
			; ALREADY DO THIS ABOVE -- IF NUMBER NOT WITH NAUGHT
			; PROBLEM MULTI TASK TIMER
			; IF ONE TAKEN HOLD LET DO COUNTER
			; -------------------------------------------------------
			; IF INSTR(%VAR_IN_NAME_5_1%,TRIM(CONTROL_TEXT_01))>0
			; {
				; Secs_MSGBOX_08=%VAR_IN_NAME_5_3%
				; SOUNDBEEP 5000,200
				; SHOW_COUNTDOWN_ACTION=TRUE
			; }
			
			; TEST DEBUG
			; TOOLTIP % SET_GO_GS "`n" VAR_IN_NAME_1  "`n" VAR_IN_NAME_4 "`n" VAR_IN_NAME_2 "`n" VAR_IN_NAME_5_1 "`n" VAR_IN_NAME_5_2 "`n" Secs_MSGBOX_08 "`n" CONTROL_TEXT_01			; LOOK FOR RESTART CODE AND LEFT DIGIT CHANGE THERE

			; IS VARIABLE COUNTER ABOVE VALUE 1 
			; AND THEN SUBTRACT NUMBER FROM IT BY 1 LESS	
			; -------------------------------------------------------
			IF Secs_MSGBOX_08>0 	
				Secs_MSGBOX_08-=1

			; ; DISPLAY NUMBER ON MSGBOX BUTTON
			; ; -------------------------------------------------------
			IF SHOW_COUNTDOWN_ACTION=TRUE
			{
				ControlSetText,%VAR_IN_NAME_5_2%,%VAR_IN_NAME_5_1%  %Secs_MSGBOX_08%, %VAR_IN_NAME_1% ahk_class #32770
			}
			
			; HAS COUNTDOWN BEEN GO AND REACH VALUE 0 
			; DO THE STUFF THEN -- POP THE BUTTON OFF & 
			; LAUNCH THE EXE IF REQUIRE
			; -------------------------------------------------------
			
			ControlGetText CONTROL_TEXT_01,%VAR_IN_NAME_5_2%,%VAR_IN_NAME_1% ahk_class #32770
			IF Secs_MSGBOX_08<1
			IF SHOW_COUNTDOWN_ACTION=TRUE
			{	
				SHOW_COUNTDOWN_ACTION=FALSE
				Secs_MSGBOX_08=0
				WINGET, VAR_IN_NAME_8, ID, %VAR_IN_NAME_1% ahk_class #32770
				LOOP, 1000
				{
					IF !VAR_IN_NAME_5_4
					{
						; NA [v1.0.45+]: May improve reliability.
						; -------------------------------------------
						ControlClick, %VAR_IN_NAME_5_2%,%VAR_IN_NAME_4%,,,, NA x10 y10 
					}
					IF VAR_IN_NAME_5_4=TRUE 
					{
						WINGET, VAR_PID, PID, %VAR_IN_NAME_1% ahk_class #32770
						IF VAR_PID>0 
							Process, Close, %VAR_PID%
					}
					WINGET HWND, ID, ahk_id %VAR_IN_NAME_8%
					IF !HWND
						BREAK
					SLEEP 500
				}
				SOUNDBEEP 4000,300
				VAR_DONE_ESCAPE_KEY=TRUE
				SLEEP 1000
				IfExist, %VAR_IN_NAME_4%
				{
					Run, "%VAR_IN_NAME_4%", , MIN
				}
			}

			
		}
	}

RETURN













