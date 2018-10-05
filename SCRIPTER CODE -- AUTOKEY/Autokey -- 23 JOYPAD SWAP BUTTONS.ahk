;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 23 JOYPAD SWAP BUTTONS.ahk
;# __ 
;# __ Autokey -- 23 JOYPAD SWAP BUTTONS.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# Thu 19-Apr-2018 01:19:10 BEGIN
;# Thu 19-Apr-2018 01:19:10 END
;# __ 
;  =============================================================

#Warn
#NoEnv
#SingleInstance Force
;--------------------
#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
;--------------------

;# ----------------------------------------------------------------------------
;# ----------------------------------------------------------------------------

; SCRIPT ======================================================================

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

setTimer TIMER_ONE,10
setTimer TIMER_PREVIOUS_INSTANCE,4000


#include %A_ScriptDir%\VJoy_lib.ahk

VJoy_Init()
nButtons := VJoy_GetVJDButtonNumber(iInterface)



~*3Joy3::
loop
{
   If !GetKeyState("3Joy3","p")
{
      break
}
else
   {
    VJoy_SetBtn(1, 1, 3)
    Sleep, 100
    VJoy_SetBtn(0, 1, 3)
      sleep, 500
      }
}
Return

;2Joy4::2Joy5

TIMER_ONE:
    Loop 16  ; Query each joystick number to find out which ones exist.
    {
        GetKeyState, JoyName, %A_Index%JoyName
        if (JoyName = "Microsoft PC-joystick driver")
        {
			TOOLTIP, %JoyName%
            ;2Joy4::2Joy5
			;2Joy5::2Joy4
			;2Joy6::2Joy8
			;2Joy8::2Joy6
            break
        }
    }
; TOOLTIP, %A_ThisHotKey% was pressed.

if (not JoyName = "Microsoft PC-joystick driver")
{
suspend 
pause
}

return



	
;Ctrl enter unable to be swapped back to Enter so use SendInput
;--------------------------------------------------------------
;GroupAdd, GroupNameTest_2, mysms
;#IfWinActive ahk_group GroupNameTest_2
;^enter::
;sendinput {ENTER}
;SetCapsLockState , Off
;SoundBeep , 2500 , 100
;Return

;}	
	


TIMER_PREVIOUS_INSTANCE:

if ScriptInstanceExist()
{
	Exitapp
}
return

ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	dhw := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % dhw
	return (match > 1)
	}
Return




;# ----------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ----------------------------------------------------------------------

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

;# ----------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
    if ExitReason not in Logoff,Shutdown
    {
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
    }
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
        ;
        ;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}
;# ------------------------------------------------------------------------------
; exit the app



;Tooltip

; sets title matching to search for "containing" instead of "exact"
;SetTitleMatchMode, 2
	
;Tooltip % ALLOW_VAR

;If ALLOW_VAR = "True" 
;^enter::b
	
;}

;If WinActive("ahk_class Chrome_WidgetWin_1") ALLOW := "True"

;if WinActive("ahk_class Chrome_WidgetWin_1")
;{
;	if WinActive("A" mysms)
;	ALLOW := "False"
;}
;#IfWinactive ahk_class Chrome_WidgetWin_1 
;#IfWinActive mysms ALLOW := "False"
;#IfWinActive mysms ALLOW := "False"


; ----------------------------------------------------------------
; With group here mysms will be in the exclude
; If chrome exist but not if another window in chrome
; ----------------------------------------------------------------
;GroupAdd, GroupNameTest_1, ahk_class Chrome_WidgetWin_1, , , mysms
;GroupAdd, GroupNameTest_2, mysms
;#IfWinActive ahk_group GroupNameTest_1
;enter::^enter
;return
	
;Ctrl enter unable to be swapped back to Enter so use SendInput 
;#IfWinActive ahk_group GroupNameTest_2
;^enter::sendinput {CapsLock}{ENTER}
;^enter::
;return


;ALLOW_VAR := "False"
;If WinActive("ahk_class" "Chrome_WidgetWin_1") 
;{
;ALLOW_VAR := "True"
;}
;MSGBOX % ALLOW_VAR
;if WinActive("mysms") 
;{
;ALLOW_VAR := "False"
;}
;MSGBOX % ALLOW_VAR

;#IfWinactive ahk_class Chrome_WidgetWin_1 
;enter::
;#IfWinActive mysms 
;^enter::SoundBeep , 2500 , 100

;If ALLOW_VAR = "True" 
;{
;sendinput, {ENTER}
;}

;RETURN

