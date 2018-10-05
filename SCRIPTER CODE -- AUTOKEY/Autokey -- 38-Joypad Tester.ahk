;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 38-Joypad Tester.ahk
;# __ 
;# __ Autokey -- 38-Joypad Tester.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

;---------------------------------------------------------------
; DESCRIPTION
;---------------------------------------------------------------
;---------------------------------------------------------------

;# ------------------------------------------------------------------
; Location OnLine
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set Google Drive
; Possible Censorship of Code Detected By Google as Malicious Happen Here
; unlike DropBox that has All Available
; https://drive.google.com/open?id=0BwoB_cPOibCPTnRZZVFuRFpHOTg
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set DropBox
; https://www.dropbox.com/sh/ntghoncyb8py1tf/AACWYrfkVn9PlqpYzNNSMcpMa?dl=0
;--------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
; 
;# ------------------------------------------------------------------

; 001
; --------------------------------------------------------------
; FROM  
; TO    
; --------------------------------------------------------------
; 
; --------------------------------------------------------------

#Warn
#NoEnv
#SingleInstance Force
;--------------------
#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
;OR      Exitapp  TO EXIT FROM #Persistent
;Exitapp CALLS ONTO ExitFunc
;--------------------

; C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\JOYPAD_WORK\CvJoyInterface.ahk
; C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\JOYPAD_WORK\VJoy_lib.ahk
; #include %A_ScriptDir%\JOYPAD_WORK\VJoy_lib.ahk
#include %A_ScriptDir%\JOYPAD_WORK\CvJoyInterface.ahk
vJoyInterface := new CvJoyInterface()
myStick := vJoyInterface.Devices[1]

F10::
	; On press of F10, hold button 1
	myStick.SetBtn(1,1)
	Return

F10 up::
	; On release of F10, release button 1
	myStick.SetBtn(0,1)
	Return

; ----
; Joystick Test - Script Example | AutoHotkey
; https://autohotkey.com/docs/scripts/JoystickTest.htm
; ----

; July 6, 2005: Added auto-detection of joystick number.
; May 8, 2005 : Fixed: JoyAxes is no longer queried as a means of
; detecting whether the joystick is connected.  Some joysticks are
; gamepads and don't have even a single axis.

; If you want to unconditionally use a specific joystick number, change
; the following value from 0 to the number of the joystick (1-16).
; A value of 0 causes the joystick number to be auto-detected:
JoystickNumber = 0

; END OF CONFIG SECTION. Do not make changes below this point unless
; you wish to alter the basic functionality of the script.

; Auto-detect the joystick number if called for:
if JoystickNumber <= 0
{
    Loop 16  ; Query each joystick number to find out which ones exist.
    {
        GetKeyState, JoyName, %A_Index%JoyName
        if JoyName <>
        {
            JoystickNumber = %A_Index%
            break
        }
    }
    if JoystickNumber <= 0
    {
        MsgBox The system does not appear to have any joysticks.
        ExitApp
    }
}

; #SingleInstance
SetFormat, float, 03  ; Omit decimal point from axis position percentages.
GetKeyState, joy_buttons, %JoystickNumber%JoyButtons
GetKeyState, joy_name, %JoystickNumber%JoyName
GetKeyState, joy_info, %JoystickNumber%JoyInfo
Loop
{
    buttons_down =
    Loop, %joy_buttons%
    {
        GetKeyState, joy%A_Index%, %JoystickNumber%joy%A_Index%
        if joy%A_Index% = D
            buttons_down = %buttons_down%%A_Space%%A_Index%
    }
    GetKeyState, JoyX, %JoystickNumber%JoyX
    axis_info = X%JoyX%
    GetKeyState, JoyY, %JoystickNumber%JoyY
    axis_info = %axis_info%%A_Space%%A_Space%Y%JoyY%
    IfInString, joy_info, Z
    {
        GetKeyState, JoyZ, %JoystickNumber%JoyZ
        axis_info = %axis_info%%A_Space%%A_Space%Z%JoyZ%
    }
    IfInString, joy_info, R
    {
        GetKeyState, JoyR, %JoystickNumber%JoyR
        axis_info = %axis_info%%A_Space%%A_Space%R%JoyR%
    }
    IfInString, joy_info, U
    {
        GetKeyState, JoyU, %JoystickNumber%JoyU
        axis_info = %axis_info%%A_Space%%A_Space%U%JoyU%
    }
    IfInString, joy_info, V
    {
        GetKeyState, JoyV, %JoystickNumber%JoyV
        axis_info = %axis_info%%A_Space%%A_Space%V%JoyV%
    }
    IfInString, joy_info, P
    {
        GetKeyState, joyp, %JoystickNumber%JoyPOV
        axis_info = %axis_info%%A_Space%%A_Space%POV%joyp%
    }
    ToolTip, %joy_name% (#%JoystickNumber%):`n%axis_info%`nButtons Down: %buttons_down%`n`n(right-click the tray icon to exit)
    Sleep, 100
}
return