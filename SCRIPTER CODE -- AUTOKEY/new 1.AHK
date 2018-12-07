
  
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

; -------------------------------------------------------------------
;WANT SELECT ALL LINE AND PASTE ONTO IT
;WANT COPY ON IT OWN
;WANT HOLD CTRL UNTIL ASK IT STOP FOR LINK URL IN WEB PAGE
; -------------------------------------------------------------------
;WANT COPY TEXT AND REPEAT PASTE IT DOWN A LINE HOME DOWN PASTE PUT REMARK IN
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location OnLine
;--------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2001-F10%20__%20HOTKEY%20__%20PRINT%20SCREEN.ahk
; ----
;# ------------------------------------------------------------------


;
;--------------------
#SingleInstance force   

		FN_VAR_2=
		FN_VAR:="C:\Program Files\WordWeb\wweb32.exe"
		IfExist, %FN_VAR%
		FN_VAR_2:=FN_VAR
		FN_VAR:="C:\Program Files (X86)\WordWeb\wweb32.exe"
		IfExist, %FN_VAR%
		FN_VAR_2:=FN_VAR
		
		MSGBOX % FN_VAR_2 
		
		IfExist, %FN_VAR_2%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR_2%" -startup , , HIDE
		}
