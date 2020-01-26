;Autokey Project to Use Enter Key on (Save As) Window Dialogue of Chrome.exe & Exit On Chrome Gone
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;
;-----------------------------------------------------------------------------------------------------------------------------------------
; Code to KILLER PROCESS SELECTED HARDCODED SCRIPTER AFTER TIMER AT BOOT OR WHEN

;-----------------------------------------------------------------------------------------------------------------------------------------
; Thu 25 Aug 2016 01:47:51 -- Length Time 1 hr 45 Min -- Finish Code

;-----------------------------------------------------------------------------------------------------------------------------------------
; 5TH Project in a Line of AutoKey -- First On Blogger Here
;-----------------------------------------------------------------------------------------------------------------------------------------

;
;--------------------
#SingleInstance force
;--------------------
;Here --- SingleInstance force -- Reduce Question Ask When Reload From Edit and More
;--------------------

; -------------------------------------------------------------------
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

; ---------------------------------------------------------------
; I MADE MENU ITEM INTO INCLUDE FILE IN 3 PART 
; 01. INTRO SETUP MENU
; 02. THE MENU ROUTINE
; 03. ANY ROUTINE THE MENU USE
; ---------------------------------------------------------------
; SAVER OF RSI INJURY AND MORE ACCURATE
; THE INCLUDE FILE ARE SAME FOLDER
; ---------------------------------------------------------------
; FROM __ Sun 09-Jun-2019 07:03:00 __ Clipboard Count = 024
; TO   __ Sun 09-Jun-2019 17:48:00 __ Clipboard Count = 452 __ 10 HOURING 45 MINUTE
; ---------------------------------------------------------------
; Create the popup menu by adding some items to it.
; MenuHandler:
; ---------------------------------------------------------------
; #Include GO WITH FULL PATH AS SOME LAUNCHER DO NOT SET WORK PATH WHEN RUNNER
; RATHER THAN CHANGE THE WORKING PATH WITHIN-AH
; ---------------------------------------------------------------
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-01_INCLUDE MENU 01 of 03.ahk



;----------------------------------------------------------------------
;My code begin
;----------------------------------------------------------------------

;WANT THIS CODE TO DO ON TIMER WHEN LOAD

;WANT CODE TO DO PROCESS PRIORITY CHANGING



Loop {
		x += 1

		sleep 1000

		if (x > 10)
		{
			SoundBeep
			break
		}
   }
;}



process=notepad.exe 
Process, Exist, %process%
if	pid :=	ErrorLevel
{
	Loop 
	{
		WinClose, ahk_pid %pid%, , 5	; will wait 5 sec for window to close
		if	ErrorLevel	; if it doesn't close
			Process, Close, %pid%	; force it 
		Process, Exist, %process%
	}	Until	!pid :=	ErrorLevel
}



While EverLoop=EverLoop
	;LOOK LIKE MUST BE CORRECT ORDER WAY AROUND
	;PROBLEM HERE LINE BELOW NOT FIND CORRECTLY THAT CHROME IS ACTIVE
	;MOVE ON NOT ENOUGH TIME
	;REM IT OUT
	;IfWinActive ahk_class Chrome_WidgetWin_1
	;	{
	;---------------------------------------------------------------
	;BACK TO SQUARE ONE
	;---------------------------------------------------------------
	;2ND TRY REPLACE WITH USE HERE INSTEAD
	;NOT INTO MULTI COUNT SCANNING
		
	;----------------------------------------------------------------------------------------------------------------
	{
	
	WinGet nChromeWindows, Count, ahk_class Chrome_WidgetWin_1
	If(nChromeWindows > 0)
	{
	
		if WinActive("Save As" ahk_class #32770)
        {
			SoundBeep
			{
			;--------------------------------------
			; For Some Reason It Maybe I Think Best to Find This Window
			; Because to Use WinActive With Look for Process Maybe Not Quite Right -- Wanting to Focus Put on the Window HWND Handle Than Process ID PID
			; And the Step Method 1.. 2.. 3.. is Speed Save As at 1st is Quicker and Poll a Process Will Use More CPU
			; process Tick's Ticker
			; Even Tough It is Done Every 500 Mill Sec Later -- More Code With Counter Can Make That Less
			;--------------------------------------
			; Modified - Is Not Much Difference to Check Below the Window Title and Class Name
			; in Term of Lead On to and The Process ID and Process Usage CPU Ticker Hungry Check to Do Both
			; The Two << Title and Class >> Name Both Same In CPU Ticker Check Term
			; Two Not same as Process Hungry CPU Process ID Check
			; WinActive("Save As" ahk_class #32770)
			
			
			;NOT WORK METHOD
			;WinActive("Save As" ahk_class #32770)
			;------------------------------
			
			WinActivate  ; Activate the window it found.
			
			;NOT WORK METHOD
			;WinActivate ("Save As" ahk_class #32770)
			
			;Send {Enter}
			;Send {Enter Down}
			;Sleep 200
			;Send {Enter Up}
			;Sendinput !+S

			Sendinput {Enter}

			}
		}
	}

	Var_Count += 1
		
	if Var_Count > 500
		{

		;500  Milli Sec with 500 -- is 250 Second without Chrome Exist and then Exit
		;1000 Milli Sec with 500 -- is 500 Second without Chrome Exist and then Exit		
		;-------------------------------------------------------------------------------------
		;-------------------
		Var_Count = 0
		;Up and Down Numeric 0
		;-------------------
		;Less Frequent Check on a Process ID a CPU Ticker Usage is High -- Enumerate the Processes
		;Double Check and Continuous Run Through With Var_Count Loop Odd Check and was Run at CPU 0.04%
		;A Continuous KeyBoard Check Run CPU Ticker At 0.01< and Less and None CPU Usage -- Process Explorer Result
		;---------------------------------------------------------------------------------------
		;Result State is Half CPU Ticker Usage at 0.02%
		;Ticker Meaning How Many Instruction Flop to Perform the Operation Command CPU Flop Thinker the Flip Flop Logic IC Integrated Circuit Terminology
		;---------------------------------------------------------------------------------------

		;------------------------------------------------------------
		;--------------
		;Sleep, 1000 ; tend to use more CPU Usage with Smaller Sleep 
		;Sleep 500
		; -----------  at 1000 Get 0.02 and 0.03% on The Asus Gl522VW Intel 7
		; -----------  Slight Less Than 0.03 and 0.04 at 500 millisecond

		;SEND KEY QUICKER SLEEP LONGER WAIT DON'T TRY REPEAT PRESSER
		Sleep 2000

		;--------------
		;It Don't Make Any Difference to Sleep 500 or 1000 Millisecond to the CPU 0.02% Usage But Take It Out and CPU Ticker 10%
		;Get a Quicker Response to Press Save As Button on 500 Than 1000 Sleeper
		;------------------------------------------------------------
		;EVER LOOP END
	}
}

;----------------------------------------------------------------------
;----------------------------------------------------------------------
;My Code Block Done
;----------------------------------------------------------------------
;----------------------------------------------------------------------


;---------------------------------------------------------
;WANT SELECT ALL LINE AND PASTE ONTO IT
;WANT COPY ON IT OWN
;WANT HOLD CTRL UNTIL ASK IT STOP FOR LINK URL IN WEB PAGE
;---------------------------------------------------------
;
;--------------------
#SingleInstance force
;--------------------
;--------------------
^F4::
		x = 0
		SoundBeep
	   
		;WinGetTitle, active_title, A
		
		;MSGBOX %O_active_title%
		;MSGBOX %active_title%
		;---------------------------------------------------------------------------------------------------------------------
		;1ST ANSWER FOUND BUT NOT GOOD ENOUGH BY WANT NUMERIC HWND INSTEAD
		;---------------------------------------------------------------------------------------------------------------------
		;GOOGLE SEARCH -- BODY SEARCH
		;--------------------------------------------------
		;FINAL GOOGLE SERCH -- AUTOHOTKEYS o_active_title = active_title
		;----
		;<SOLVED> IfWinActive, ahk_class, Title (exact name) Co - Ask for Help - AutoHotkey Community
		;https://autohotkey.com/board/topic/74614-solved-ifwinactive-ahk-class-title-exact-name-co/
		;----
		
		WinGetTitle, active_title, A
		IF active_title <> O_active_title
		{
			SoundBeep
			break
		}
		O_active_title = %active_title%
		
		;------------------------------------------
		;CHROME ACTIVE FIND
		;------------------------------------------
		;NEXT ANSWER IN THIS STAGE OF 2ND BUT NOT TOTALLY
		;IS WANTING NUMERIC OF HWND CHANGE IDEA
		;------------------------------------------
		IfWinNotActive ahk_class Chrome_WidgetWin_1
		{
			;SoundBeep
			break
		}
		;--------------------------------------------------------------------
		;ALL CONDITION LOOK THEN OKAY TO PROCEED 
		;--------------------------------------------------------------------
	    Loop {
				x += 1
				;Send,{PgDn} 
				Send,{Down}

				sleep 40

				if GetKeyState("RButton")
				{
					SoundBeep
					break
				}
				if (x > 4000000)
				{
					SoundBeep
					break
				}
	   }
;}



;-------------------------------------------------------------------------------------------   
;UN USEFUL CODE BLOCK ABOVE
;WinGet nChromeWindows, Count, ahk_class Chrome_WidgetWin_1
;If(nChromeWindows > 0)
;-------------------------------------------------------------------------------------------   
;		if WinNOTActive(ahk_exe chrome.exe)
;		{
;			SoundBeep
;			break
;		}
;-------------------------------------------------------------------------------------------   
   
   
;-------------------------------------------------------------------------
;REPLACE F10 TO DO CONTROL PRINT SCREEN
;FOR CLIPBOARD SCREEN SHOT -- CODE PIC___ WON'T HOT KEY F0 ON WIN 7 AND UP
;-------------------------------------------------------------------------
;F10::
;SEND ^{PrintScreen}

Return

;-------------------------------------------------------------
;MAKE WIN + Y WORK TO DO A SCROLL LOCK FOR WINDOWSE STOP CLEAR
;-------------------------------------------------------------
;NOT WORKING TO DO A WIN + KEY TO RELACE
;LIKE AUTOKEY USE WIN + A -- TO STOP A SPY
;-------------------------------------------------------------
;^::

;#Y::^L

;#Y::^ScrollLock

;#Y::
;SEND ^L
;SEND ScrollLock
;-------------------------------------------------------------

;---------------------------------------------------
;CTRL F FIND
;---------------------------------------------------
;F6::
;   Send,{ctrl down}f{ctrl up}{ENTER}
;   Send,{ctrl down}f{ctrl up}
;---------------------------------------------------
;CTRL F FIND AND ENTER ON HIGH-LIGHTED SELECTED TEXT
;---------------------------------------------------
;F6::
;   Send,{ctrl down}f{ctrl up}{ENTER}

MenuHandler:
#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-02_INCLUDE MENU 02 of 03.ahk
return

#Include C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 00-03_INCLUDE MENU 03 of 03.ahk



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

