; =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
;# __ 
;# __ Autokey -- 73-MSGBOX COUNTDOWN DELAY_INCLUDE_02.ahk
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

SUB_ROUTINE_INCLUDE:

IF INSTR(A_ScriptName,"_INCLUDE.ahk")>0 
	EXITAPP
	Element_1 = %A_ScriptFullPath%
	Element_22 := SUBSTR(Element_1,1,INSTR(Element_1,".ahk")-1)
	Element_2 = %Element_22%_INCLUDE.ahk
	Element_3 = %Element_22%_INCLUDE.ahk

	IfExist, %Element_1%
		FileGetTime, DATE_MOD_1, %Element_1%, M
	IfExist, %Element_2%
		FileGetTime, DATE_MOD_2, %Element_2%, M
	IfExist, %Element_3%
		FileGetTime, DATE_MOD_3, %Element_3%, M

	IF O_DATE_MOD_1
	IF O_DATE_MOD_1<>%DATE_MOD_1%
	{
		title := " - AutoHotkey v" A_AhkVersion
		DetectHiddenWindows, On
		WinGet, id, List, % A_ScriptFullPath . title
		IF  id > 1
			SLEEP 2000
		SCRIPTOR_OWN_PID=% DllCall("GetCurrentProcessId")
		Loop, %id%
		{
			WinGet, PID_8, PID, % "ahk_id " id%A_Index% 
			IF PID_8
			If PID_8 <> %SCRIPTOR_OWN_PID%
			{
				Process, Close, %PID_8% 
				; SOUNDBEEP 1200,40
			}
		}
	}
	
	IF O_DATE_MOD_1
	IF O_DATE_MOD_1<>%DATE_MOD_1% 
	{
		SLEEP 4000
 		Run, %Element_1%
	}

	IF O_DATE_MOD_2
	IF O_DATE_MOD_2<>%DATE_MOD_2% 
		Run, %Element_1%

	IF O_DATE_MOD_3
	IF O_DATE_MOD_3<>%DATE_MOD_3% 
		Run, %Element_1%

	O_DATE_MOD_1=%DATE_MOD_1%
	O_DATE_MOD_2=%DATE_MOD_2%
	O_DATE_MOD_3=%DATE_MOD_3%

RETURN













