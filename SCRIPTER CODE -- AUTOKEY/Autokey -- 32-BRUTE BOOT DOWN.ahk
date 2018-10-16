;====================================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 32-BRUTE BOOT DOWN.ahk
;# __ 
;# __ Autokey -- 32-BRUTE BOOT DOWN.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;# START TIME [ Fri 22:19:40 Pm_08 Jun 2018 ]
;# END   TIME [ Sat 00:10:00 Pm_09 Jun 2018 ] 1 HOUR 40 MINUTE
;# __ 
;====================================================================

;# ------------------------------------------------------------------
; DESCRIPTION 
;# ------------------------------------------------------------------
; ANY PROGRAMS THAT ARE KNOWN TO HAVE BOOT DOWN PROBLEM
; WITH MESSENGER AT SHUTDOWN __ END NOW
; THE BUTTON IS PRESSED FOR THEM AFTER GRACE SHORT TIME
; LEAVING ANY PROGRAM THAT REQUIRE A GRACEFUL SHUT DOWN TO BE LEFT ALONE
; SYSTEMEXPLORER HAS A SHUTDOWN FAULT SOME COMPUTER
; FILLEZILLA SERVER HAS A SHUTDOWN FAULT SOME COMPUTER
;
; GOOGLESYNCDRIVE REQUIRE GRACEFUL SHUTDOWN
; GOODSYNC REQUIRE GRACEFUL SHUT DOWN
;
; THIS PROGRAM WILL NOT GET KNOCKED OUT BY SHUTDOWN REQUEST
; ONLY UNTIL THE OTHER PROGRAMS ARE GONE
; YOU CAN EXIT WITH RIGHT CLICK TO CLOSE 
; ONLY SHUTDOWN REQUEST MODE IS INTERCEPTED
; SHUTDOWN OR LOGG OFF
;
; THIS PROGRAM IS CALLED FROM AT BOOTER ONLY IF REQUIRED 
; C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 21-AUTORUN.ahk
;
; AT SHUTDOWN WINDOWS REPORT THE WINDOW_TITLE AS PROBLEM PROGRAM
; BUT NOT ALWAYS GIVES AWAY THE .EXE TALKING ABOUT
; USE THIS CODE TO TIDY UP AND BAD SHUTDOWN PROGRAMS
; WITH THE REM LINES LEFT IN TO CONVERT WINDOW_TITLE TO EXE PATH
;
; FINALLY THE PROCESS IS MORE SPEEDY 
;# ------------------------------------------------------------------

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
; https://www.dropbox.com/s/bsvrrmyknxr7pvd/Autokey%20--%2018-NORTON%20CONTROL%20BOOTER.ahk?dl=0
;# ------------------------------------------------------------------

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force
; --------------------
#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
; --------------------
; -------------------------------------------------------------------

GLOBAL EXIT_APP_VAR

EXIT_APP_VAR=FALSE
I_COUNT=0

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

DetectHiddenWindows, ON

SLEEP 4000

settimer MAIN_RUNNER, 1000

RETURN

MAIN_RUNNER:

SET_GO=TRUE
If WinExist("SystemExplorer")
	SET_GO=FALSE
If WinExist("CAsyncSocketEx Helper Window")
	SET_GO=FALSE
	
I_COUNT+=1
IF I_COUNT<480
	SET_GO=FALSE

; IF SET_GO=TRUE
; {
	; EXIT_APP_VAR=TRUE
	; SoundBeep , 2000 , 100
	; SoundBeep , 2500 , 100
	; EXITAPP
; }


;---------------------------------------------------
; THESE PAIR WON'T REALLY GET RUN BUT LEFT IN ANYWAY
; AS WORK TO FIND OUT FIRST OFF
; PRIORITY IS IN THE EXITAPP
;---------------------------------------------------

;--------------------------------------
;C:\PStart\Progs\#_PortableApps\PortableApps\SystemExplorerPortable\App\SystemExplorer\SystemExplorer.exe	
;--------------------------------------
IfWinExist End Program - SystemExplorer	
{
	; Run, "TASKKILL.exe" /F /IM SystemExplorer.exe /T , , HIDE
	Sleep 8000
	SoundBeep , 2500 , 100
	ControlClick, &End Now, End Program - SystemExplorer
}

;------------------------------
;FileZilla Server Interface.exe
;CAsyncSocketEx Helper Window
;------------------------------
IfWinExist End Program - CAsyncSocketEx Helper Window
{
	;WinGet, path, ProcessName, CAsyncSocketEx Helper Window
	Sleep 18000
	SoundBeep , 2500 , 100
	ControlClick, &End Now, End Program - CAsyncSocketEx Helper Window
}	

RETURN



;# ------------------------------------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

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

;# ------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ------------------------------------------------------------------

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

;# ------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
    if ExitReason not in Logoff,Shutdown
    {
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
    }
    if ExitReason in Logoff,Shutdown
    {
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
		
		;---------------------------------------------------------------------
		; TASKKILL.exe WON'T WORK CAN'T LAUNCH ANOTHER PROGRAM DURING SHUTDOWN
		;---------------------------------------------------------------------
		;Run, "TASKKILL.exe" /F /IM SystemExplorer.exe /T , , MIN
		;Run, "TASKKILL.exe" /F /IM FileZilla Server Interface.exe /T , , MIN
		
		;---------------------------------------------------------------------
		; WaitClose DOSEN;T WORK HER REM LINE
		; MUST BE KILLED INSTANTLY
		;---------------------------------------------------------------------
		; Process, WaitClose, ahk_exe SystemExplorer.exe, 2
		; Process, WaitClose, ahk_exe FileZilla Server Interface.exe, 2
		;---------------------------------------------------------------------
		
		;---------------------------------------------------------------------
		; THE PROBLEM IS ABOUT REQUIRING THIS WAY
		; IS IF THIS PROGRAM GET REQUEST TO SHUTDOWN BEFORE THE OTHER FEW COUPLE
		; AND THEN A HOLD UP PROBLEM
		; SO HAVE TO BE KILLED AS SOON AS SHUTDOWN REQUEST COMES ALONG
		;---------------------------------------------------------------------
		
		Process, Close, SystemExplorer.exe
		Process, Close, FileZilla Server Interface.exe
		
		;---------------------------------------------------------------------
		; THE RETURN 1 IS STILL USED BUT THE APP CLOSES ITSELF 
		; PROBABLY BECAUSE THE COUPLE OF PROGRAM ARE FOUND NOT TO EXIST
		; BUT EVEN WITH A MSGBOX IN THE WAY IT STILL GET CLOSE
		;---------------------------------------------------------------------
		
		IF EXIT_APP_VAR=FALSE 
			{
			; MSGBOX % EXIT_APP_VAR
			return 1
			}
	}
	
	; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
        ;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}

; -----------------------------------------------------------------
; exit the app