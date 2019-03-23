#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

;Autokey -- 13-LOGITECH_SPEAKER_AUDIO 1Hz.ahk

;---------------------------------------------------------
;WANT 
;---------------------------------------------------------

;
;--------------------
#SingleInstance force
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------



;---------------------------------------------------------
; CODE INITIALIZE SOUND EFFECT LEARN
;---------------------------------------------------------
SoundBeep , 1500 , 400
;soundSetWaveVolume, 100 

SUSPEND
PAUSE

SetTimer, Player, 60000

Return

Player:
	;SoundBeep , 2000 , 100
	SoundPlay, C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 13-LOGITECH_SPEAKER_AUDIO 1Hz.wav
    
	;SoundPlay, %A_WinDir%\Media\ding.wav
	;SoundPlay, C:\Windows\Media\ding.wav
	
	;SoundPlay *-1  ; Simple beep. If the sound card is not available, the sound is generated using the speaker.
Return


;----------------------------------------------------------------------
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
;----------------------------------------------------------------------


;----
; HOW TO STOP THE LOGITECH SPEAKER GOING INTO SLEEP MODE - Google Search
;https://www.google.co.uk/search?num=50&rlz=1C1CHBF_en-GBGB759GB759&q=HOW+TO+STOP+THE+LOGITECH+SPEAKER+GOING+INTO+SLEEP+MODE&spell=1&sa=X&ved=0ahUKEwje9rrbvZbWAhWDY1AKHVhZCdEQvwUIJSgA&biw=1536&bih=646
;--------
; Prevent Speaker's Standby Mode [Solved] - Peripherals - Audio
;http://www.tomsguide.com/answers/id-2417519/prevent-speaker-standby-mode.html
;----