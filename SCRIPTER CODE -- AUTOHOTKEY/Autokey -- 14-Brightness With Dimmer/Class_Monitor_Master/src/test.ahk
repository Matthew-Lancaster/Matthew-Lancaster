; GLOBAL SETTINGS ===============================================================================================================

#Warn
#NoEnv
#SingleInstance Force

#Include Class_Monitor.ahk                                                   ; include the class here

global cnt := DllCall("user32.dll\GetSystemMetrics", "Int", 80)              ; get the number of display monitors on a desktop
global tmp := []                                                             ; initialize an Array

OnExit, EOF                                                                  ; set onexit to the label "EOF"
Display := New Monitor()                                                     ; initialize / start the class

; SCRIPT ========================================================================================================================

loop % cnt                                                                   ; loop throw all monitors
    tmp[A_Index] := Display.GetMonitorBrightness(1).CurrentBrightness        ; get the current brightness for all monitors


#Numpad4::                                                                   ; hotkey: win + numpad4 / win + numpad left
#NumpadLeft::
    loop % tmp.MaxIndex()                                                    ; loop throw all monitors
        Display.SetMonitorBrightness(A_Index, --tmp[A_Index])                ; decrease the brightness by 1 for all monitors
return

#Numpad6::                                                                   ; hotkey: win + numpad6 / win + numpad right
#NumpadRight::
    loop % tmp.MaxIndex()                                                    ; loop throw all monitors
        Display.SetMonitorBrightness(A_Index, ++tmp[A_Index])                ; increase the brightness by 1 for all monitors
return

#Numpad5::                                                                   ; hotkey: win + numpad5
    loop % tmp.MaxIndex()                                                    ; loop throw all monitors
        Display.SetMonitorBrightness(A_Index, tmp[A_Index] := 50)            ; set the brightness back to 50 for all monitors
return

; EXIT ==========================================================================================================================

EOF:                                                                         ; on exit
Display.OnExit()                                                             ; release and free the library
ExitApp      